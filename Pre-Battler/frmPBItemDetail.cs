using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using General;
using Infragistics.Win.UltraWinGrid;


namespace Pre_Battler
{
    public partial class frmPBItemDetail : Form
    {
        DataView gridView;
        DataTable gridTable;
        DataTable rankDistinct;
        int sessionID;
        bool bulkupdate = false;
        string SessionName; //Why the hell does this get set back to blank after the form loads???
        string userName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;

        public frmPBItemDetail(string SITEID, string SKUID, string SHIPTOS, string SESSION)
        {
            InitializeComponent();

            string SiteID = SITEID;
            string SkuID = SKUID;
            string ShipTos = SHIPTOS;
            SessionName = SESSION;

            LoadGrid(SkuID, SiteID, ShipTos, SessionName);

            setStoreCount();
            updateSKUCounter();
        }
        #region Major functions
        private void LoadGrid(string SiteID, string SkuID, string ShipTos, string SessionName)
        {
            //////////////////////////////////////////////////////////////
            //
            //  Function: LoadGrid
            //  ------------------
            //
            //  This was originally meant to handle the loading of the main grid, however as time marched on
            //  it became a collection of tasks needed to handle the data that's shuffled around on the form load
            //
            //////////////////////////////////////////////////////////////

            //Get all the for the specfied stores and site
            DataSet ds = new DataSet();
            string query = "usp_PB_SelectAvailable @SKULevelID=N'" + SkuID + "', @SiteID=" + SiteID + ", @ShipTos=N'" + ShipTos + "', @SessionName=N'" + SessionName + "', @userName=N'" + userName + "'";
            ds = Global.GetData(query);

            //Set up grid
            gridView = new DataView(ds.Tables[0]);
            gridTable = gridView.ToTable();

            gridTable.Columns.Add("WOH", typeof(System.Decimal)); //Add column for weeks on hand
            gridTable.Columns.Add("TotalShelves", typeof(System.Decimal)); //Add column for total shelves
            gridTable.Columns.Add("TotalOH", typeof(System.Int32)); //Add column for total OH
            calculateShelves();
            calculateTotalOH();
            UpdateWOH("WOH last week"); //Add column for WOH and update using sold last week

            ugrdItemDetail.DataSource = gridTable.DefaultView;

            //Set active session ID info
            SessionName = gridTable.Rows[0]["SessionName"].ToString();
            this.Text = "Distribute SKUs - " + SessionName;
            lblSite.Text = "Ship Site: " + gridTable.Rows[0]["SiteDesc"].ToString();
            sessionID = Convert.ToInt32(gridTable.Rows[0]["sessionID"]);

            //create a custom view for distinct SKUs, put a view all at the top
            DataView SKUview = new DataView(ds.Tables[0]);
            //SKUview.Sort = "SKUDesc ASC";
            SKUview.Sort = "SKUSize ASC, SKUDesc ASC";
            DataTable SKUdistinct = SKUview.ToTable(true, "SKUDesc");

            DataRow AllRow = SKUdistinct.NewRow();
            AllRow[0] = "View All";
            SKUdistinct.Rows.InsertAt(AllRow, 0);

            //create a custom view for store ranking
            DataView rankView = new DataView(ds.Tables[0]);
            rankView.Sort = "Rank ASC";
            rankDistinct = rankView.ToTable(true, "Rank");

            rankDistinct.Columns.Add("Quantity", typeof(System.Int32));
            //Set all quantities to zero. Future, set to the previous value.
            foreach (DataRow r in rankDistinct.Rows)
            {
                r["Quantity"] = 0;
            }

            //Set the datasources on the combo box and rank grid
            cmbSKUs.DataSource = SKUdistinct.DefaultView;
            cmbSKUs.ValueMember = "SKUDesc";
            cmbSKUs.DisplayMember = "SKUDesc";
            cmbSKUs.SelectedIndex = 0;
            ugrdRanks.DataSource = rankDistinct;

            Global.GridLayout(ugrdItemDetail, 2, Properties.Settings.Default.grdItemDetailLayout); //Load grid layout
            GridLayout(); //Apply formatting to grid and override settings we don't want them to change
            updateSummary(0);
        }

        private void updateSQL(int id = 0, int qty = 0)
        {
            //If the function was passed a specific record to update, lets do something efficiently
            if (id > 0)
            {
                Global.GetData("usp_PB_UpdateDetail @id=" + id + ", @qty=" + qty);
            }
            else //else, just update the whole damn thing.
            {
                string tableQ = "create table #tmp (id integer not null, qty integer not null)";
                string updateQ = "update sd set sd.qtyRequested = tmp.qty from PB_SessionDetail sd inner join #tmp tmp " +
                    "on sd.id=tmp.id where tmp.qty<>sd.qtyRequested; ";//drop table #tmp;
                DataView tmpvw = new DataView(gridTable);
                DataTable tmptbl = tmpvw.ToTable("Table", false, "id", "qtyRequested");
                Global.BulkUpdate(tmptbl, "#tmp",tableQ, updateQ);
            }
        }

        private void UpdateCellMath(int RowNum)
        {
            UltraGridRow updatedRow = ugrdItemDetail.Rows[RowNum];
            //DataRow updatedRow = gridTable.Rows[RowNum]; This won't work, we need the rows from the grid.
            string prodId = updatedRow.Cells["prodId"].Value.ToString();//this is the row we changed

            //Round to nearest pack qty
            decimal rounded = Convert.ToInt32(updatedRow.Cells["qtyRequested"].Value);
            int packQty = Convert.ToInt32(updatedRow.Cells["PackQty"].Value);
            rounded = Math.Round(rounded / packQty) * packQty;


            int NewDemand = 0;
            //Function does some basic math to recalculate the available quantities
            //basically steps through the rows that match the currently selected SKU and
            //recalculates the available inventory

            foreach (DataRow dr in gridTable.Rows)
            {
                if (dr["Id"].ToString() == updatedRow.Cells["Id"].Value.ToString())
                    dr["qtyRequested"] = rounded;

                if (dr["prodId"].ToString() == prodId)
                {
                    NewDemand += Convert.ToInt32(dr["qtyRequested"]); //step through the rows, if the SKU matches, add the requested qty up for a new demand number
                }
            }

            //The new available amount is a bunch of math taking our newdemand in account
            int newAvailable = Convert.ToInt32(updatedRow.Cells["qtyInStock"].Value) - NewDemand - Convert.ToInt32(updatedRow.Cells["qtyOnOrder"].Value) - Convert.ToInt32(updatedRow.Cells["qtyCommitted"].Value);

            foreach (DataRow dr in gridTable.Rows)
            {
                if (dr["prodId"].ToString() == prodId)
                {
                    dr["available"] = newAvailable; //update the available amount for all rows with matching SKU
                }
            }

            if (!bulkupdate)
            {
                calculateShelves(Convert.ToInt32(updatedRow.Cells["Id"].Value.ToString()));
                calculateTotalOH(Convert.ToInt32(updatedRow.Cells["Id"].Value.ToString()));
            }
        }

        private void calculateShelves(int row = 0)
        {
            /// Currently get a stack overflow if we try to target the actual changed row. The issue is that in order to target the row,
            /// you have to actually target the row in the grid, not the datatable. The row number the grid passes you changes based on 
            /// how it's been filtered, and doesn't always match the underlying table. That in itself isn't a problem, but when you change 
            /// the data in the grid, it fires off the cell math function, which in turn fires this off again, which updates the cell again, 
            /// which starts the whole process over. Doing it like this isn't causing much of a speed hit at the moment, so we keep as is.


            foreach (DataRow dr in gridTable.Rows)
            {
                if (row > 0)
                {
                    if (dr["Id"].ToString() == row.ToString())
                        dr["TotalShelves"] = Convert.ToDecimal(dr["qtyPerShelf"]) > 0 ?
                                (Convert.ToDecimal(dr["qtyRequested"]) / Convert.ToDecimal(dr["qtyPerShelf"])).ToString("0.##")
                                : "0";
                }
                else
                {
                    dr["TotalShelves"] = Convert.ToDecimal(dr["qtyPerShelf"]) > 0 ?
                                (Convert.ToDecimal(dr["qtyRequested"]) / Convert.ToDecimal(dr["qtyPerShelf"])).ToString("0.##")
                                : "0";
                }
            }
        }

        private void calculateTotalOH(int row = 0)
        {
            foreach (DataRow dr in gridTable.Rows)
            {
                if (row > 0)
                {
                    if (dr["Id"].ToString() == row.ToString())
                        dr["TotalOH"] = Convert.ToInt32(dr["qtyRequested"]) + Convert.ToInt32(dr["OnHand"]);
                }
                else
                {
                    dr["TotalOH"] = Convert.ToInt32(dr["qtyRequested"]) + Convert.ToInt32(dr["OnHand"]);
                }
            }
        }

        private void UpdateWOH(string WOHMethod)
        {
            //Updates the WOH column depending on what option is passed in the input string
            decimal sold;
            decimal onhand;
            decimal woh;

            foreach (DataRow dr in gridTable.Rows)
            {
                sold = 0;
                onhand = 0;
                woh = 0;

                onhand = Convert.ToInt32(dr["OnHand"].ToString());
                //Only update WOH if there is an onhand quantity, else leave it set to zero
                if (onhand > 0)
                {
                    switch (WOHMethod)
                    {
                        case "WOH last week":
                            sold = Convert.ToInt32(dr["LW_Sold"].ToString());
                            break;
                        case "WOH this week":
                            sold = Convert.ToInt32(dr["CW_Sold"].ToString());
                            break;
                        case "WOH 2 weeks ago":
                            sold = Convert.ToInt32(dr["TW_Sold"].ToString());
                            break;
                        case "WOH 2 week average":
                            sold = (Convert.ToInt32(dr["TW_Sold"].ToString()) + Convert.ToInt32(dr["LW_Sold"].ToString())) / 2;
                            break;
                    }

                    woh = sold > 0 ? onhand / sold : 999;
                }
                dr["WOH"] = woh.ToString("0.#");
            }
        }

        private void GridLayout()
        {
            //A whole shitload of predefined layout stuff for the grid.
            //Not completely redundant with the addition of the column choose, but a little less useful now

            //A probably more robust way of handling this should be developed so we can add columns without recompiling the software.

            ugrdItemDetail.DisplayLayout.Bands[0].Columns["id"].Hidden = true;
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["sessionId"].Hidden = true;
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["shipId"].Hidden = true;
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["sessionName"].Hidden = true;
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["prodId"].Hidden = true;
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["qtyPerShelf"].Hidden = true;
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["SiteDesc"].Hidden = true;
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["PackQty"].Hidden = true;

            ugrdItemDetail.DisplayLayout.Bands[0].Columns["id"].CellActivation = Activation.Disabled;
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["sessionId"].CellActivation = Activation.Disabled;
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["shipId"].CellActivation = Activation.Disabled;
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["sessionName"].CellActivation = Activation.Disabled;
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["prodId"].CellActivation = Activation.Disabled;
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["SiteDesc"].CellActivation = Activation.Disabled;
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["companyName"].CellActivation = Activation.Disabled;
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["SKUDesc"].CellActivation = Activation.Disabled;
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["SKUSize"].CellActivation = Activation.Disabled;
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["SKU"].CellActivation = Activation.Disabled;
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["available"].CellActivation = Activation.Disabled;
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["qtyInStock"].CellActivation = Activation.Disabled;
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["qtyOnOrder"].CellActivation = Activation.Disabled;
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["qtyCommitted"].CellActivation = Activation.Disabled;
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["Rank"].CellActivation = Activation.Disabled;
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["Zone"].CellActivation = Activation.Disabled;
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["YTDShipped"].CellActivation = Activation.Disabled;
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["CW_Sold"].CellActivation = Activation.Disabled;
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["LW_Sold"].CellActivation = Activation.Disabled;
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["TW_Sold"].CellActivation = Activation.Disabled;
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["YTDSold"].CellActivation = Activation.Disabled;
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["OnHand"].CellActivation = Activation.Disabled;
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["InTransit"].CellActivation = Activation.Disabled;
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["TotalShelves"].CellActivation = Activation.Disabled;
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["TotalOH"].CellActivation = Activation.Disabled;
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["SiteDesc"].CellActivation = Activation.Disabled;
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["YTDDumps"].CellActivation = Activation.Disabled;
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["Product"].CellActivation = Activation.Disabled;
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["PackQty"].CellActivation = Activation.Disabled;
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["Region"].CellActivation = Activation.Disabled;
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["Allocated"].CellActivation = Activation.Disabled;

            ugrdRanks.DisplayLayout.Bands[0].Columns["Rank"].CellActivation = Activation.Disabled;

            ugrdItemDetail.DisplayLayout.Bands[0].Columns["qtyRequested"].Header.Caption = "Quantity";
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["companyName"].Header.Caption = "Store Name";
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["SKUDesc"].Header.Caption = "SKU Desc";
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["SKUSize"].Header.Caption = "SKU Size";
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["SKU"].Header.Caption = "SKU #";
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["available"].Header.Caption = "Qty Avail";
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["qtyInStock"].Header.Caption = "Qty On Hand";
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["qtyOnOrder"].Header.Caption = "Qty NOW";
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["qtyCommitted"].Header.Caption = "Qty Committed";
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["YTDShipped"].Header.Caption = "Shipped YTD";
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["CW_Sold"].Header.Caption = "Sold CW";
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["LW_Sold"].Header.Caption = "Sold LW";
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["TW_Sold"].Header.Caption = "Sold L2W";
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["YTDSold"].Header.Caption = "Sold YTD";
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["OnHand"].Header.Caption = "On Hand";
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["TotalShelves"].Header.Caption = "Total Shelves";
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["TotalOH"].Header.Caption = "Total Inv OH";
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["YTDDumps"].Header.Caption = "YTD Dumped";
            ugrdItemDetail.DisplayLayout.Bands[0].Columns["Product"].Header.Caption = "Product Code";
        }

        private void updateSummary(int row)
        {
            lblStore.Text = ugrdItemDetail.Rows[row].Cells["companyName"].Text;//Set store name
            lblSKU.Text = ugrdItemDetail.Rows[row].Cells["SKUDesc"].Text;//Set SKU name

            //Update store units
            object sumUnitsByStore = gridTable.Compute("Sum(qtyRequested)", "shipId=" + ugrdItemDetail.Rows[row].Cells["shipId"].Text);
            lblStoreUnits.Text = "Total Units: " + sumUnitsByStore.ToString();
            //Update SKU units
            object sumUnitsBySKU = gridTable.Compute("Sum(qtyRequested)", "prodId=" + ugrdItemDetail.Rows[row].Cells["prodId"].Text);
            lblSKUUnits.Text = "Total Units: " + sumUnitsBySKU.ToString();

            //Update session units
            object sumUnitsTotal = gridTable.Compute("Sum(qtyRequested)", "");
            lblSessionUnits.Text = "Total Requested: " + sumUnitsTotal.ToString();

            //Update store shelves
            object sumShelvesByStore = gridTable.Compute("Sum(TotalShelves)", "shipId=" + ugrdItemDetail.Rows[row].Cells["shipId"].Text);
            lblStoreShelf.Text = "Total Shelf: " + sumShelvesByStore.ToString();
            //Update SKU shelves
            object sumShelvesBySKU = gridTable.Compute("Sum(TotalShelves)", "prodId=" + ugrdItemDetail.Rows[row].Cells["prodId"].Text);
            lblSKUShelf.Text = "Total Shelf: " + sumShelvesBySKU.ToString();
        }
        #endregion
        #region BS Fucntions
        private void setupPGBar(int records)
        {
            pgBar.Value = 0;
            pgBar.Maximum = records;
            pgBar.Visible = true;
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            if (cmbSKUs.SelectedIndex > 0)
            {
                cmbSKUs.SelectedIndex -= 1;
            }
        }

        private void cmbWOH_SelectedIndexChanged(object sender, EventArgs e)
        {
            //We uhhh, update the WOH column using the rantazzle function somewhere else
            UpdateWOH(cmbWOH.Text);
        }

        private void ugrdRanks_KeyDown(object sender, KeyEventArgs e)
        {
            Global.GridNavigation(ugrdRanks, e);
        }

        private void updateSKUCounter()
        {
            //fantastic function that does the amazing task of counting where we are in the SKU list and updating a label!
            //A FREAKING LABEL!!
            string maxSku = (cmbSKUs.Items.Count - 1).ToString();
            string curSku = cmbSKUs.SelectedIndex.ToString();

            lblSKUCounter.Text = "SKU " + curSku + "/" + maxSku;
        }

        private void ugrdItemDetail_KeyDown(object sender, KeyEventArgs e)
        {
            Global.GridNavigation(ugrdItemDetail, e);
        }

        private void btnForward_Click(object sender, EventArgs e)
        {
            if (cmbSKUs.SelectedIndex < cmbSKUs.Items.Count - 1)
            {
                cmbSKUs.SelectedIndex += 1;
            }
        }

        private void cmbSKUs_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Changes the active SKU desc filter, so they can use back and next buttons to navigate.
            ugrdItemDetail.DisplayLayout.Bands[0].ColumnFilters["SKUDesc"].FilterConditions.Clear();
            string filter = cmbSKUs.SelectedValue.ToString();
            FilterCondition gridFilter = new FilterCondition(FilterComparisionOperator.Equals, filter);

            if (cmbSKUs.SelectedIndex > 0)
                ugrdItemDetail.DisplayLayout.Bands[0].ColumnFilters["SKUDesc"].FilterConditions.Add(gridFilter);
            updateSKUCounter();
        }

        private void setStoreCount()
        {
            //Function to count the number of stores in the grid
            DataView tempView = new DataView(gridTable);
            DataTable tempTable = tempView.ToTable(true, "shipId");

            lblTotalStores.Text = "Total Stores: " + tempTable.Rows.Count.ToString();

            tempTable.Clear();
        }

        private void lblExpand_Click(object sender, EventArgs e)
        {
            //Increase the height of the rank panel for extra visibility
            pnlRank.Height = pnlRank.Height == 124 ? 460 : 124;
        }

        private void ugrdItemDetail_AfterCellUpdate(object sender, CellEventArgs e)
        {

            UpdateCellMath(e.Cell.Row.Index);
            if (!bulkupdate)
                updateSQL(Convert.ToInt32(e.Cell.Row.Cells["id"].Text), Convert.ToInt32(e.Cell.Row.Cells["qtyRequested"].Text));
        }

        private void sAVELAYOUTToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Call global4k's gridsave to make a layout in to a byte stream in to a byte array in to a string
            Properties.Settings.Default.grdItemDetailLayout = Global.GridLayout(ugrdItemDetail, 1);
            Properties.Settings.Default.Save();
        }

        private void sELECTCOLUMNSToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Show column chooser
            ugrdItemDetail.ShowColumnChooser();
        }

        private void ugrdItemDetail_AfterCellActivate(object sender, EventArgs e)
        {
            updateSummary(ugrdItemDetail.ActiveRow.Index);
        }

        private void ugrdItemDetail_AfterRowActivate(object sender, EventArgs e)
        {
            updateSummary(ugrdItemDetail.ActiveRow.Index);
        }
        #endregion
        #region Distribution Methods

        private void btnFull_Click(object sender, EventArgs e)
        {

            this.Cursor = Cursors.WaitCursor;
            bulkupdate = true;
            int rowindex = 0;
            decimal rounded = 0;
            int packQty = 0;
            DataRow rank;
            setupPGBar(ugrdItemDetail.Rows.GetFilteredInNonGroupByRows().Count());

            //we went through some steps to speed this shit up. The grid responds a ton faster if you turn painting off, updated the dataset, then refresh the grid
            ugrdItemDetail.BeginUpdate();
            ugrdItemDetail.SuspendRowSynchronization();
            ugrdItemDetail.SuspendSummaryUpdates();

            foreach (UltraGridRow r in ugrdItemDetail.Rows.GetFilteredInNonGroupByRows())
            {
                //Get rank row and index of dataset
                rank = rankDistinct.Select("Rank = '" + r.Cells["Rank"].Value + "'")[0];
                rowindex = r.ListIndex;
                //Round to nearest pack qty
                rounded = Convert.ToInt32(rank["Quantity"]);
                packQty = Convert.ToInt32(gridTable.Rows[rowindex]["PackQty"]);
                rounded = Math.Round(rounded / packQty) * packQty;

                gridTable.Rows[rowindex]["qtyRequested"] = rounded; //This is fast
                //r.Cells["qtyRequested"].Value = rank["Quantity"]; This is slow
                pgBar.PerformStep();
            }
            //end of speed enhancements
            ugrdItemDetail.ResumeRowSynchronization();
            ugrdItemDetail.ResumeSummaryUpdates(true);
            ugrdItemDetail.EndUpdate();

            ugrdItemDetail.UpdateData();
            calculateShelves();
            calculateTotalOH();
            pgBar.Visible = false;
            updateSQL();
            bulkupdate = false;
            this.Cursor = Cursors.Default;
        }

        private void btnWOH_Click(object sender, EventArgs e)
        {
            //Calculating the needed quantity to fullfill WOH is relativly simple...
            //We need to know the desired WOH, the actual WOH, and the quantity sold in the period 
            //used for the actual WOH calculation. Actual WOH is the units on hand / units sold in a certain 'week' period
            DataRow rank;
            int requestedWOH;
            int sold;
            int onhand;
            int woh;
            int demand;
            int rowindex = 0;
            decimal rounded = 0;
            int packQty = 0;
            string WOHMethod = cmbWOH.Text;

            setupPGBar(ugrdItemDetail.Rows.GetFilteredInNonGroupByRows().Count());

            //we went through some steps to speed this shit up. The grid responds a ton faster if you turn painting off, updated the dataset, then refresh the grid
            ugrdItemDetail.BeginUpdate();
            ugrdItemDetail.SuspendRowSynchronization();
            ugrdItemDetail.SuspendSummaryUpdates();

            this.Cursor = Cursors.WaitCursor;
            bulkupdate = true;
            foreach (UltraGridRow r in ugrdItemDetail.Rows.GetFilteredInNonGroupByRows())
            {
                //reset vars
                sold = 0;
                onhand = 0;
                woh = 0;
                demand = 0;
                requestedWOH = 0;
                rowindex = r.ListIndex;

                //Check rank table and find the target WOH
                rank = rankDistinct.Select("Rank = '" + r.Cells["Rank"].Value + "'")[0];
                requestedWOH = Convert.ToInt32(rank["Quantity"].ToString()); //We have a requested amount now

                //Get the on hand
                onhand = Convert.ToInt32(gridTable.Rows[rowindex]["OnHand"]);

                //Calculate the sold
                if (onhand > 0)
                {
                    switch (WOHMethod)
                    {
                        case "WOH last week":
                            sold = Convert.ToInt32(gridTable.Rows[rowindex]["LW_Sold"]);
                            break;
                        case "WOH this week":
                            sold = Convert.ToInt32(gridTable.Rows[rowindex]["CW_Sold"]);
                            break;
                        case "WOH 2 weeks ago":
                            sold = Convert.ToInt32(gridTable.Rows[rowindex]["TW_Sold"]);
                            break;
                        case "WOH 2 week average":
                            sold = (Convert.ToInt32(gridTable.Rows[rowindex]["TW_Sold"]) + Convert.ToInt32(gridTable.Rows[rowindex]["LW_Sold"])) / 2;
                            break;
                    }

                    woh = sold > 0 ? onhand / sold : 999; //And now we have WOH
                }

                demand = requestedWOH - woh > 0 ? (requestedWOH - woh) * sold : 0;
                //Round to nearest pack qty
                rounded = demand;
                packQty = Convert.ToInt32(gridTable.Rows[rowindex]["PackQty"]);
                rounded = Math.Round(rounded / packQty) * packQty;

                gridTable.Rows[rowindex]["qtyRequested"] = rounded;

                pgBar.PerformStep();
            }
            //end of speed enhancements
            ugrdItemDetail.ResumeRowSynchronization();
            ugrdItemDetail.ResumeSummaryUpdates(true);
            ugrdItemDetail.EndUpdate();

            ugrdItemDetail.UpdateData();
            calculateShelves();
            calculateTotalOH();
            pgBar.Visible = false;
            updateSQL();
            bulkupdate = false;
            this.Cursor = Cursors.Default;
        }

        private void btnMin_Click(object sender, EventArgs e)
        {
            DataRow rank;
            int onHand = 0;
            int requestedQty = 0;
            int newQty = 0;
            int rowindex = 0;
            decimal rounded = 0;
            int packQty = 0;

            setupPGBar(ugrdItemDetail.Rows.GetFilteredInNonGroupByRows().Count());

            this.Cursor = Cursors.WaitCursor;
            bulkupdate = true;

            //we went through some steps to speed this shit up. The grid responds a ton faster if you turn painting off, updated the dataset, then refresh the grid
            ugrdItemDetail.BeginUpdate();
            ugrdItemDetail.SuspendRowSynchronization();
            ugrdItemDetail.SuspendSummaryUpdates();

            foreach (UltraGridRow r in ugrdItemDetail.Rows.GetFilteredInNonGroupByRows())
            {
                rank = rankDistinct.Select("Rank = '" + r.Cells["Rank"].Value + "'")[0];
                rowindex = r.ListIndex;

                onHand = Convert.ToInt32(gridTable.Rows[rowindex]["OnHand"]);
                requestedQty = Convert.ToInt32(rank["Quantity"].ToString());
                newQty = requestedQty - onHand > 0 ? requestedQty - onHand : 0;

                //Round to nearest pack qty
                rounded = newQty;
                packQty = Convert.ToInt32(gridTable.Rows[rowindex]["PackQty"]);
                rounded = Math.Round(rounded / packQty) * packQty;

                gridTable.Rows[rowindex]["qtyRequested"] = rounded;

                pgBar.PerformStep();
            }
            //end of speed enhancements
            ugrdItemDetail.ResumeRowSynchronization();
            ugrdItemDetail.ResumeSummaryUpdates(true);
            ugrdItemDetail.EndUpdate();

            ugrdItemDetail.UpdateData();
            calculateShelves();
            calculateTotalOH();
            pgBar.Visible = false;
            updateSQL();
            bulkupdate = false;
            this.Cursor = Cursors.Default;
        }
        #endregion

        private void eXPORTTOEXCELToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.SaveFileDialog saveFile = new System.Windows.Forms.SaveFileDialog();
            saveFile.FileName = gridTable.Rows[0]["SessionName"].ToString() == "" ? "untitled.xlsx" : gridTable.Rows[0]["SessionName"].ToString() + ".xlsx";
            saveFile.Filter = "Excel (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
            saveFile.RestoreDirectory = true;

            if (saveFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                pnlStatus.Visible = true;
                bgExport.RunWorkerAsync(saveFile.FileName);
            }

        }

        private void rEFRESHAVAILABILITYToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("This process will fetch new in stock quantities from inventory, and update the saved session. No work will be lost, " +
                "however available amounts may change based on units actually on hand or committed. Do you wish to continue?") == DialogResult.OK)
            {
                Global.GetData("usp_PB_UpdateAvailable @sessionId="+sessionID);
                LoadGrid(sessionID.ToString(), "LOAD", "", "");
                MessageBox.Show("Refresh Complete!");
            }
        }

        private void bgExport_DoWork(object sender, DoWorkEventArgs e)
        {
            xl4k.GenerateExport(sessionID, gridTable.Rows[0]["SessionName"].ToString(), (string)e.Argument);
        }

        private void bgExport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            pnlStatus.Visible = false;
        }

        private void aDDSTORETOSESSIONToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmAddStore frm = new frmAddStore();
            frm.ShowDialog();
        }
    }
}
