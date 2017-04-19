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
    public partial class frmSaveLoad : Form
    {
        bool initialLoadComplete = false;
        int selectedStores = 0;
        System.Data.DataView grdShipToView;
        System.Data.DataTable grdShipToTable;

        public frmSaveLoad()
        {
            InitializeComponent();
            fnSetupDB();
            initialLoadComplete = true;
            fnLoadGrid();

            
        }

        #region Major functions

        private void fnLoadGrid()
        {
            string skulevel = cmbSKULevel.SelectedValue.ToString();
            string billto = cmbAccount.SelectedValue.ToString();

            grdShipToView = new System.Data.DataView(Global.GetData("exec usp_PB_SelectStores @SKULevel='" + skulevel + "', @BillTo=" + billto).Tables[0]);
            grdShipToTable = grdShipToView.ToTable();
            grdShipToTable.Columns.Add("Selected", typeof(System.Boolean));
            foreach (DataRow dr in grdShipToTable.Rows)
                dr["Selected"] = false;

            grdStores.DataSource = grdShipToTable.DefaultView;
            btnNext.Enabled = false;
            GridLayout();
        }

        private void fnSetupDB()
        {
            string dbServer = Properties.Settings.Default.dbServer;
            string dbName = Properties.Settings.Default.dbName;

            if (dbServer == "" || dbName == "")
            {
                if (MessageBox.Show("Database is not currently configure. Do this now?", "Setup DB",MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    PBOptions form = new PBOptions();
                    form.ShowDialog();
                }
            }

            try
            {
                //Connect to server
                Global.SetConnectionString(dbServer,dbName);
                
                //Fill in and enable the accounts combo
                cmbAccount.DataSource = Global.GetData("exec usp_PB_SelectAccounts").Tables[0].DefaultView;
                cmbAccount.DisplayMember = "companyName";
                cmbAccount.ValueMember = "id";

                //Get active sites
                cmbSite.DataSource = Global.GetData("exec usp_PB_SelectSites").Tables[0].DefaultView;
                cmbSite.DisplayMember = "description";
                cmbSite.ValueMember = "id";

                //Enable combo, set it to the last selected customer
                cmbAccount.Enabled = true;
                cmbAccount.SelectedIndex = Properties.Settings.Default.defCustomer;

                cmbSKULevel.DataSource = Global.GetData("exec usp_PB_SelectSKULevels @BillToID=" + cmbAccount.SelectedValue.ToString()).Tables[0].DefaultView;
                cmbSKULevel.DisplayMember = "description";
                cmbSKULevel.ValueMember = "skuLevelID";

                cmbSKULevel.Enabled = true;
                cmbSKULevel.SelectedIndex = 0;

                //Update status bar
                sbConnected.Text = "Online";
                sbConnected.ForeColor = System.Drawing.Color.Green;

            }
            catch
            { MessageBox.Show("There was a problem connection. Check settings and double click the status bar to attempt again."); }
        }

        private void GridLayout()
        {
            grdStores.DisplayLayout.Bands[0].Columns["Selected"].Header.VisiblePosition = 0;
            grdStores.DisplayLayout.Bands[0].Columns["companyName"].Header.VisiblePosition = 1;
            grdStores.DisplayLayout.Bands[0].Columns["city"].Header.VisiblePosition = 2;
            grdStores.DisplayLayout.Bands[0].Columns["st"].Header.VisiblePosition = 3;
            grdStores.DisplayLayout.Bands[0].Columns["custom1"].Header.VisiblePosition = 4;
            grdStores.DisplayLayout.Bands[0].Columns["custom2"].Header.VisiblePosition = 5;
            grdStores.DisplayLayout.Bands[0].Columns["custom3"].Header.VisiblePosition = 6;
            grdStores.DisplayLayout.Bands[0].Columns["custom4"].Header.VisiblePosition = 7;

            grdStores.DisplayLayout.Bands[0].Columns["id"].Hidden = true;
            
            grdStores.DisplayLayout.Bands[0].Columns["companyName"].CellActivation = Activation.Disabled;
            grdStores.DisplayLayout.Bands[0].Columns["city"].CellActivation = Activation.Disabled;
            grdStores.DisplayLayout.Bands[0].Columns["st"].CellActivation = Activation.Disabled;
            grdStores.DisplayLayout.Bands[0].Columns["custom1"].CellActivation = Activation.Disabled;
            grdStores.DisplayLayout.Bands[0].Columns["custom2"].CellActivation = Activation.Disabled;
            grdStores.DisplayLayout.Bands[0].Columns["custom3"].CellActivation = Activation.Disabled;
            grdStores.DisplayLayout.Bands[0].Columns["custom4"].CellActivation = Activation.Disabled;

            grdStores.DisplayLayout.Bands[0].Columns["Selected"].Header.Caption = "Selected";
            grdStores.DisplayLayout.Bands[0].Columns["companyName"].Header.Caption = "Store Name";
            grdStores.DisplayLayout.Bands[0].Columns["city"].Header.Caption = "City";
            grdStores.DisplayLayout.Bands[0].Columns["st"].Header.Caption = "State";
            grdStores.DisplayLayout.Bands[0].Columns["custom1"].Header.Caption = "Custom 1";
            grdStores.DisplayLayout.Bands[0].Columns["custom2"].Header.Caption = "Custom 2";
            grdStores.DisplayLayout.Bands[0].Columns["custom3"].Header.Caption = "Custom 3";
            grdStores.DisplayLayout.Bands[0].Columns["custom4"].Header.Caption = "Custom 4";

            //Load grid layout
            battleEtc.GridLayout(grdStores,2,Properties.Settings.Default.grdStoreLayout);
        }
        #endregion
        #region BS Functions
        private void cLOSEToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void updateCount(int activeCell)
        {
            selectedStores = Convert.ToInt32(grdShipToTable.Compute("COUNT(Selected)", "Selected='true'")) + activeCell;
            sbStoreCount.Text = "Selected Stores: " + selectedStores.ToString();
        }
        private void oPTIONSToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PBOptions form = new PBOptions();
            form.ShowDialog();
        }

        private void sbConnected_Click(object sender, EventArgs e)
        {
            fnSetupDB();
        }

        private void grdStores_AfterRowFilterChanged(object sender, Infragistics.Win.UltraWinGrid.AfterRowFilterChangedEventArgs e)
        {
            chkSelectAll.Checked = false;
        }
        #endregion
        #region Comboboxes
        private void cmbAccount_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Do some misc cleanup
            Properties.Settings.Default.defCustomer = cmbAccount.SelectedIndex;
            Properties.Settings.Default.Save();
            chkSelectAll.Checked = false;

            if (initialLoadComplete)
            {
                try
                {
                    //Lets try and get some SKU levels now
                    cmbSKULevel.DataSource = Global.GetData("exec usp_PB_SelectSKULevels @BillToID=" + cmbAccount.SelectedValue.ToString()).Tables[0].DefaultView;
                    cmbSKULevel.DisplayMember = "description";
                    cmbSKULevel.ValueMember = "skuLevelID";

                    cmbSKULevel.Enabled = true;
                    cmbSKULevel.SelectedIndex = 0;
                }
                catch { MessageBox.Show("There was an error retrieving SKU Levels"); }
            }
        }

        private void cmbSKULevel_SelectedIndexChanged(object sender, EventArgs e)
        {
            chkSelectAll.Checked = false;//misc cleanup

            grdStores.DataSource = null;
            try
            {
                if (cmbSKULevel.SelectedIndex > -1) { fnLoadGrid(); }
                cmbSite.Enabled = true;
            }
            catch { MessageBox.Show("There was a problem loading the store data in to the grid."); }
        }

        private void cmbSite_SelectedIndexChanged(object sender, EventArgs e)
        {
            //btnNext.Enabled = true;
        }

        #endregion

        private void btnNext_Click(object sender, EventArgs e)
        {
            List<string> ShipTos = new List<string>();
            int SelectedRecords = 0;

            if (pnlLoadSession.Visible == true)
            {
                frmPBItemDetail form = new frmPBItemDetail("LOAD", ugrdSession.ActiveRow.Cells["id"].Text, "", "");
                this.Hide();
                form.ShowDialog();
                this.Show();
            }
            else
            {
                foreach (DataRow dr in grdShipToTable.Rows)
                {
                    if (Convert.ToBoolean(dr["Selected"]) == true)
                    {
                        ShipTos.Add(dr["id"].ToString());
                        SelectedRecords++;
                    }
                }

                if (SelectedRecords > 0)
                {
                    string STs = string.Join(",", ShipTos.ToArray());
                    frmPBItemDetail form = new frmPBItemDetail(cmbSKULevel.SelectedValue.ToString(), cmbSite.SelectedValue.ToString(), STs, txtSessionName.Text);
                    this.Hide();
                    form.ShowDialog();
                    this.Show();
                }
                else
                {
                    MessageBox.Show("No records selected. Please select some stores to continue.");
                }
            }
        }

        private void grdStores_CellChange(object sender, CellEventArgs e)
        {
            int activeCell = bool.Parse(this.grdStores.ActiveCell.Text) ? 1 : -1;
            updateCount(activeCell);
        }

        private void grdStores_InitializeLayout(object sender, InitializeLayoutEventArgs e)
        {
            updateCount(0);
        }

        private void txtSessionName_TextChanged(object sender, EventArgs e)
        {
            if (txtSessionName.Text.Length >= 3)
            { btnNext.Enabled = true; }
            else { btnNext.Enabled = false; }
        }

        private void lOADSAVEDSESSIONToolStripMenuItem_Click(object sender, EventArgs e)
        {
            pnlLoadSession.Visible = true;

            grdStores.Enabled = false;
            cmbAccount.Enabled = false;
            cmbSKULevel.Enabled = false;
            cmbSite.Enabled = false;
            txtSessionName.Enabled = false;
            chkSelectAll.Enabled = false;
            btnNext.Enabled = false;
            btnExport.Enabled = false;
            btnImport.Enabled = false;

            loadSessions();
        }

        private void loadSessions()
        {
            ugrdSession.DataSource = Global.GetData("usp_PB_SelectSessions").Tables[0];

            ugrdSession.DisplayLayout.Bands[0].Columns["id"].Hidden = true;

            ugrdSession.DisplayLayout.Bands[0].Columns["Session Name"].CellActivation = Activation.Disabled;
            ugrdSession.DisplayLayout.Bands[0].Columns["User"].CellActivation = Activation.Disabled;
            ugrdSession.DisplayLayout.Bands[0].Columns["Date Created"].CellActivation = Activation.Disabled;
        }

        private void lblClose_Click(object sender, EventArgs e)
        {
            pnlLoadSession.Visible = false;

            grdStores.Enabled = true;
            cmbAccount.Enabled = true;
            cmbSKULevel.Enabled = true;
            cmbSite.Enabled = true;
            txtSessionName.Enabled = true;
            chkSelectAll.Enabled = true;
            btnExport.Enabled = true;
            btnImport.Enabled = true;

            if (txtSessionName.Text.Length >= 3)
                btnNext.Enabled = true; 
        }

        private void ugrdSession_AfterRowActivate(object sender, EventArgs e)
        {
            btnNext.Enabled = true;
        }

        private void grdStores_KeyDown(object sender, KeyEventArgs e)
        {
            battleEtc.GridNavigation(grdStores, e);
        }

        private void chkSelectAll_Click(object sender, EventArgs e)
        {
            foreach (UltraGridRow r in grdStores.Rows.GetFilteredInNonGroupByRows())
            {
                r.Cells["Selected"].Value = chkSelectAll.Checked;
            }
            grdStores.UpdateData();

            updateCount(0);
        }

        private void sAVELAYOUTToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Save the grid layout
            Properties.Settings.Default.grdStoreLayout = battleEtc.GridLayout(grdStores, 1);
            Properties.Settings.Default.Save();
        }

        private void sELECTCOLUMNSToolStripMenuItem_Click(object sender, EventArgs e)
        {
            grdStores.ShowColumnChooser();
        }

        private void ugrdSession_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                //figure out where the cursor is, get the UI element it clicked on
                Point cursor = new Point(e.X, e.Y);
                Infragistics.Win.UIElement clicked = ((UltraGrid)sender).DisplayLayout.UIElement.ElementFromPoint(cursor);

                //try to get the cell if one was clicked on
                UltraGridCell cell = clicked.GetContext(typeof(UltraGridCell)) as UltraGridCell;

                if (cell != null) //Assuming we actually clicked a cell
                {
                    miInactive.Text = "Make " + cell.Row.Cells["Session Name"].Text + " Inactive";
                    miInactive.Tag = cell.Row.Cells["id"].Text;
                    mnSessionRightClick.Show(ugrdSession, cursor);
                    cell.Row.Selected = true;
                }
            }
        }

        private void miInactive_Click(object sender, EventArgs e)
        {
            Global.GetData("usp_PB_UpdateSessions @sessionId=" + miInactive.Tag.ToString());
            loadSessions();
        }

        private void ugrdSession_DoubleClickRow(object sender, DoubleClickRowEventArgs e)
        {
            btnNext.PerformClick();
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            string skulevel = cmbSKULevel.SelectedValue.ToString();
            string billto = cmbAccount.SelectedValue.ToString();
            string filename = xl4k.getFilePath(1, "export");
            xl4k.basicExcelExport(Global.GetData("exec usp_PB_SelectStores @SKULevel='" + skulevel + "', @BillTo=" + billto).Tables[0], filename);
            //xl4k.basicExcelExport(grdShipToTable, filename);
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            string filename = xl4k.getFilePath(2);
            grdShipToTable = xl4k.basicExcelImport(filename);
            grdShipToTable.Columns.Add("Selected", typeof(System.Boolean));
            foreach (DataRow dr in grdShipToTable.Rows)
                dr["Selected"] = true;

            grdStores.DataSource = grdShipToTable.DefaultView;
            btnNext.Enabled = false;
            GridLayout();
        }
    }
}
