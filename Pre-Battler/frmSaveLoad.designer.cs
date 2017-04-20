namespace Pre_Battler
{
    partial class frmSaveLoad
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
            Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmSaveLoad));
            this.cmbAccount = new System.Windows.Forms.ComboBox();
            this.cmbSKULevel = new System.Windows.Forms.ComboBox();
            this.cmbSite = new System.Windows.Forms.ComboBox();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fILEToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cLOSEToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tOOLSToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.lOADSAVEDSESSIONToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.sELECTCOLUMNSToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.sAVELAYOUTToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.sbConnected = new System.Windows.Forms.ToolStripStatusLabel();
            this.sbStoreCount = new System.Windows.Forms.ToolStripStatusLabel();
            this.btnNext = new System.Windows.Forms.Button();
            this.grdStores = new Infragistics.Win.UltraWinGrid.UltraGrid();
            this.chkSelectAll = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.txtSessionName = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.pnlLoadSession = new System.Windows.Forms.Panel();
            this.lblClose = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.ugrdSession = new Infragistics.Win.UltraWinGrid.UltraGrid();
            this.mnSessionRightClick = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.miInactive = new System.Windows.Forms.ToolStripMenuItem();
            this.btnExport = new System.Windows.Forms.Button();
            this.btnImport = new System.Windows.Forms.Button();
            this.menuStrip1.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.grdStores)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.pnlLoadSession.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ugrdSession)).BeginInit();
            this.mnSessionRightClick.SuspendLayout();
            this.SuspendLayout();
            // 
            // cmbAccount
            // 
            this.cmbAccount.FormattingEnabled = true;
            this.cmbAccount.Location = new System.Drawing.Point(64, 21);
            this.cmbAccount.Name = "cmbAccount";
            this.cmbAccount.Size = new System.Drawing.Size(178, 21);
            this.cmbAccount.TabIndex = 0;
            this.cmbAccount.SelectedIndexChanged += new System.EventHandler(this.cmbAccount_SelectedIndexChanged);
            // 
            // cmbSKULevel
            // 
            this.cmbSKULevel.Enabled = false;
            this.cmbSKULevel.FormattingEnabled = true;
            this.cmbSKULevel.Location = new System.Drawing.Point(327, 21);
            this.cmbSKULevel.Name = "cmbSKULevel";
            this.cmbSKULevel.Size = new System.Drawing.Size(121, 21);
            this.cmbSKULevel.TabIndex = 1;
            this.cmbSKULevel.SelectedIndexChanged += new System.EventHandler(this.cmbSKULevel_SelectedIndexChanged);
            // 
            // cmbSite
            // 
            this.cmbSite.Enabled = false;
            this.cmbSite.FormattingEnabled = true;
            this.cmbSite.Location = new System.Drawing.Point(506, 21);
            this.cmbSite.Name = "cmbSite";
            this.cmbSite.Size = new System.Drawing.Size(121, 21);
            this.cmbSite.TabIndex = 2;
            this.cmbSite.SelectedIndexChanged += new System.EventHandler(this.cmbSite_SelectedIndexChanged);
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fILEToolStripMenuItem,
            this.tOOLSToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1100, 24);
            this.menuStrip1.TabIndex = 3;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fILEToolStripMenuItem
            // 
            this.fILEToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.cLOSEToolStripMenuItem});
            this.fILEToolStripMenuItem.Name = "fILEToolStripMenuItem";
            this.fILEToolStripMenuItem.Size = new System.Drawing.Size(40, 20);
            this.fILEToolStripMenuItem.Text = "FILE";
            // 
            // cLOSEToolStripMenuItem
            // 
            this.cLOSEToolStripMenuItem.Name = "cLOSEToolStripMenuItem";
            this.cLOSEToolStripMenuItem.Size = new System.Drawing.Size(109, 22);
            this.cLOSEToolStripMenuItem.Text = "CLOSE";
            this.cLOSEToolStripMenuItem.Click += new System.EventHandler(this.cLOSEToolStripMenuItem_Click);
            // 
            // tOOLSToolStripMenuItem
            // 
            this.tOOLSToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.lOADSAVEDSESSIONToolStripMenuItem,
            this.sELECTCOLUMNSToolStripMenuItem,
            this.sAVELAYOUTToolStripMenuItem});
            this.tOOLSToolStripMenuItem.Name = "tOOLSToolStripMenuItem";
            this.tOOLSToolStripMenuItem.Size = new System.Drawing.Size(55, 20);
            this.tOOLSToolStripMenuItem.Text = "TOOLS";
            // 
            // lOADSAVEDSESSIONToolStripMenuItem
            // 
            this.lOADSAVEDSESSIONToolStripMenuItem.Name = "lOADSAVEDSESSIONToolStripMenuItem";
            this.lOADSAVEDSESSIONToolStripMenuItem.Size = new System.Drawing.Size(190, 22);
            this.lOADSAVEDSESSIONToolStripMenuItem.Text = "LOAD SAVED SESSION";
            this.lOADSAVEDSESSIONToolStripMenuItem.Click += new System.EventHandler(this.lOADSAVEDSESSIONToolStripMenuItem_Click);
            // 
            // sELECTCOLUMNSToolStripMenuItem
            // 
            this.sELECTCOLUMNSToolStripMenuItem.Name = "sELECTCOLUMNSToolStripMenuItem";
            this.sELECTCOLUMNSToolStripMenuItem.Size = new System.Drawing.Size(190, 22);
            this.sELECTCOLUMNSToolStripMenuItem.Text = "SELECT COLUMNS";
            this.sELECTCOLUMNSToolStripMenuItem.Click += new System.EventHandler(this.sELECTCOLUMNSToolStripMenuItem_Click);
            // 
            // sAVELAYOUTToolStripMenuItem
            // 
            this.sAVELAYOUTToolStripMenuItem.Name = "sAVELAYOUTToolStripMenuItem";
            this.sAVELAYOUTToolStripMenuItem.Size = new System.Drawing.Size(190, 22);
            this.sAVELAYOUTToolStripMenuItem.Text = "SAVE LAYOUT";
            this.sAVELAYOUTToolStripMenuItem.Click += new System.EventHandler(this.sAVELAYOUTToolStripMenuItem_Click);
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.sbConnected,
            this.sbStoreCount});
            this.statusStrip1.Location = new System.Drawing.Point(0, 711);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(1100, 22);
            this.statusStrip1.TabIndex = 4;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // sbConnected
            // 
            this.sbConnected.ForeColor = System.Drawing.Color.Red;
            this.sbConnected.Name = "sbConnected";
            this.sbConnected.Size = new System.Drawing.Size(43, 17);
            this.sbConnected.Text = "Offline";
            this.sbConnected.Click += new System.EventHandler(this.sbConnected_Click);
            // 
            // sbStoreCount
            // 
            this.sbStoreCount.Name = "sbStoreCount";
            this.sbStoreCount.Size = new System.Drawing.Size(1042, 17);
            this.sbStoreCount.Spring = true;
            this.sbStoreCount.Text = "Selected Stores: 0";
            this.sbStoreCount.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // btnNext
            // 
            this.btnNext.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnNext.Enabled = false;
            this.btnNext.Location = new System.Drawing.Point(993, 33);
            this.btnNext.Name = "btnNext";
            this.btnNext.Size = new System.Drawing.Size(95, 47);
            this.btnNext.TabIndex = 5;
            this.btnNext.Text = "NEXT ->";
            this.btnNext.UseVisualStyleBackColor = true;
            this.btnNext.Click += new System.EventHandler(this.btnNext_Click);
            // 
            // grdStores
            // 
            this.grdStores.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            appearance1.BackColor = System.Drawing.SystemColors.ControlLight;
            appearance1.BackColor2 = System.Drawing.SystemColors.ControlDark;
            appearance1.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
            appearance1.ForeColorDisabled = System.Drawing.Color.Black;
            this.grdStores.DisplayLayout.Appearance = appearance1;
            this.grdStores.DisplayLayout.AutoFitStyle = Infragistics.Win.UltraWinGrid.AutoFitStyle.ResizeAllColumns;
            this.grdStores.DisplayLayout.ColumnChooserEnabled = Infragistics.Win.DefaultableBoolean.True;
            this.grdStores.DisplayLayout.GroupByBox.Hidden = true;
            this.grdStores.DisplayLayout.InterBandSpacing = 10;
            this.grdStores.DisplayLayout.MaxColScrollRegions = 1;
            this.grdStores.DisplayLayout.MaxRowScrollRegions = 1;
            this.grdStores.DisplayLayout.Override.AllowRowFiltering = Infragistics.Win.DefaultableBoolean.True;
            appearance2.BackColor = System.Drawing.Color.Transparent;
            this.grdStores.DisplayLayout.Override.CardAreaAppearance = appearance2;
            appearance3.BackColor = System.Drawing.SystemColors.Control;
            appearance3.BackColor2 = System.Drawing.SystemColors.ControlLightLight;
            appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.VerticalBump;
            this.grdStores.DisplayLayout.Override.CellAppearance = appearance3;
            this.grdStores.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText;
            appearance4.BackColor = System.Drawing.SystemColors.Control;
            appearance4.BackColor2 = System.Drawing.SystemColors.ControlDark;
            appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.VerticalBump;
            appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
            this.grdStores.DisplayLayout.Override.HeaderAppearance = appearance4;
            this.grdStores.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
            appearance5.ForeColor = System.Drawing.SystemColors.ActiveCaption;
            this.grdStores.DisplayLayout.Override.RowSelectorAppearance = appearance5;
            appearance6.BackColor = System.Drawing.SystemColors.InactiveCaption;
            appearance6.BackColor2 = System.Drawing.SystemColors.ActiveCaption;
            appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.VerticalBump;
            this.grdStores.DisplayLayout.Override.SelectedRowAppearance = appearance6;
            this.grdStores.DisplayLayout.RowConnectorColor = System.Drawing.SystemColors.ControlDarkDark;
            this.grdStores.DisplayLayout.RowConnectorStyle = Infragistics.Win.UltraWinGrid.RowConnectorStyle.Dashed;
            this.grdStores.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
            this.grdStores.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
            this.grdStores.Location = new System.Drawing.Point(0, 111);
            this.grdStores.Name = "grdStores";
            this.grdStores.Size = new System.Drawing.Size(1100, 597);
            this.grdStores.TabIndex = 6;
            this.grdStores.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.grdStores_InitializeLayout);
            this.grdStores.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.grdStores_CellChange);
            this.grdStores.AfterRowFilterChanged += new Infragistics.Win.UltraWinGrid.AfterRowFilterChangedEventHandler(this.grdStores_AfterRowFilterChanged);
            this.grdStores.KeyDown += new System.Windows.Forms.KeyEventHandler(this.grdStores_KeyDown);
            // 
            // chkSelectAll
            // 
            this.chkSelectAll.AutoSize = true;
            this.chkSelectAll.Location = new System.Drawing.Point(21, 88);
            this.chkSelectAll.Name = "chkSelectAll";
            this.chkSelectAll.Size = new System.Drawing.Size(70, 17);
            this.chkSelectAll.TabIndex = 7;
            this.chkSelectAll.Text = "Select All";
            this.chkSelectAll.UseVisualStyleBackColor = true;
            this.chkSelectAll.Click += new System.EventHandler(this.chkSelectAll_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(3, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(58, 13);
            this.label1.TabIndex = 8;
            this.label1.Text = "Account:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(253, 24);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(71, 13);
            this.label2.TabIndex = 9;
            this.label2.Text = "SKU Level:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(470, 24);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(33, 13);
            this.label3.TabIndex = 10;
            this.label3.Text = "Site:";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.cmbSite);
            this.groupBox1.Controls.Add(this.cmbSKULevel);
            this.groupBox1.Controls.Add(this.cmbAccount);
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(14, 27);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(652, 52);
            this.groupBox1.TabIndex = 11;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Session Setup";
            // 
            // txtSessionName
            // 
            this.txtSessionName.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.txtSessionName.Location = new System.Drawing.Point(690, 49);
            this.txtSessionName.Name = "txtSessionName";
            this.txtSessionName.Size = new System.Drawing.Size(285, 20);
            this.txtSessionName.TabIndex = 12;
            this.txtSessionName.TextChanged += new System.EventHandler(this.txtSessionName_TextChanged);
            // 
            // label4
            // 
            this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(687, 33);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(87, 13);
            this.label4.TabIndex = 11;
            this.label4.Text = "Session Name";
            // 
            // pnlLoadSession
            // 
            this.pnlLoadSession.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pnlLoadSession.Controls.Add(this.lblClose);
            this.pnlLoadSession.Controls.Add(this.label5);
            this.pnlLoadSession.Controls.Add(this.ugrdSession);
            this.pnlLoadSession.Location = new System.Drawing.Point(190, 159);
            this.pnlLoadSession.Name = "pnlLoadSession";
            this.pnlLoadSession.Size = new System.Drawing.Size(720, 415);
            this.pnlLoadSession.TabIndex = 13;
            this.pnlLoadSession.Visible = false;
            // 
            // lblClose
            // 
            this.lblClose.AutoSize = true;
            this.lblClose.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblClose.Location = new System.Drawing.Point(641, 6);
            this.lblClose.Name = "lblClose";
            this.lblClose.Size = new System.Drawing.Size(67, 13);
            this.lblClose.TabIndex = 11;
            this.lblClose.Text = "CLOSE [X]";
            this.lblClose.Click += new System.EventHandler(this.lblClose_Click);
            // 
            // label5
            // 
            this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(4, 6);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(223, 13);
            this.label5.TabIndex = 14;
            this.label5.Text = "Select session to load and press next.";
            // 
            // ugrdSession
            // 
            appearance7.BackColor = System.Drawing.SystemColors.ControlLight;
            appearance7.BackColor2 = System.Drawing.SystemColors.ControlDark;
            appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
            appearance7.ForeColorDisabled = System.Drawing.Color.Black;
            this.ugrdSession.DisplayLayout.Appearance = appearance7;
            this.ugrdSession.DisplayLayout.AutoFitStyle = Infragistics.Win.UltraWinGrid.AutoFitStyle.ResizeAllColumns;
            this.ugrdSession.DisplayLayout.GroupByBox.Hidden = true;
            this.ugrdSession.DisplayLayout.InterBandSpacing = 10;
            this.ugrdSession.DisplayLayout.MaxColScrollRegions = 1;
            this.ugrdSession.DisplayLayout.MaxRowScrollRegions = 1;
            appearance8.BackColor = System.Drawing.Color.Transparent;
            this.ugrdSession.DisplayLayout.Override.CardAreaAppearance = appearance8;
            appearance9.BackColor = System.Drawing.SystemColors.Control;
            appearance9.BackColor2 = System.Drawing.SystemColors.ControlLightLight;
            appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.VerticalBump;
            this.ugrdSession.DisplayLayout.Override.CellAppearance = appearance9;
            this.ugrdSession.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
            appearance10.BackColor = System.Drawing.SystemColors.Control;
            appearance10.BackColor2 = System.Drawing.SystemColors.ControlDark;
            appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.VerticalBump;
            appearance10.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
            this.ugrdSession.DisplayLayout.Override.HeaderAppearance = appearance10;
            this.ugrdSession.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
            appearance11.ForeColor = System.Drawing.SystemColors.ActiveCaption;
            this.ugrdSession.DisplayLayout.Override.RowSelectorAppearance = appearance11;
            appearance12.BackColor = System.Drawing.SystemColors.InactiveCaption;
            appearance12.BackColor2 = System.Drawing.SystemColors.ActiveCaption;
            appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.VerticalBump;
            this.ugrdSession.DisplayLayout.Override.SelectedRowAppearance = appearance12;
            this.ugrdSession.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
            this.ugrdSession.DisplayLayout.RowConnectorColor = System.Drawing.SystemColors.ControlDarkDark;
            this.ugrdSession.DisplayLayout.RowConnectorStyle = Infragistics.Win.UltraWinGrid.RowConnectorStyle.Dashed;
            this.ugrdSession.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
            this.ugrdSession.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
            this.ugrdSession.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
            this.ugrdSession.Location = new System.Drawing.Point(7, 22);
            this.ugrdSession.Name = "ugrdSession";
            this.ugrdSession.Size = new System.Drawing.Size(701, 381);
            this.ugrdSession.TabIndex = 0;
            this.ugrdSession.AfterRowActivate += new System.EventHandler(this.ugrdSession_AfterRowActivate);
            this.ugrdSession.DoubleClickRow += new Infragistics.Win.UltraWinGrid.DoubleClickRowEventHandler(this.ugrdSession_DoubleClickRow);
            this.ugrdSession.MouseUp += new System.Windows.Forms.MouseEventHandler(this.ugrdSession_MouseUp);
            // 
            // mnSessionRightClick
            // 
            this.mnSessionRightClick.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.miInactive});
            this.mnSessionRightClick.Name = "mnSessionRightClick";
            this.mnSessionRightClick.Size = new System.Drawing.Size(148, 26);
            // 
            // miInactive
            // 
            this.miInactive.Name = "miInactive";
            this.miInactive.Size = new System.Drawing.Size(147, 22);
            this.miInactive.Text = "make inactive";
            this.miInactive.Click += new System.EventHandler(this.miInactive_Click);
            // 
            // btnExport
            // 
            this.btnExport.Location = new System.Drawing.Point(126, 85);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(140, 19);
            this.btnExport.TabIndex = 14;
            this.btnExport.Text = "Export store list to excel";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // btnImport
            // 
            this.btnImport.Location = new System.Drawing.Point(295, 85);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(140, 19);
            this.btnImport.TabIndex = 15;
            this.btnImport.Text = "Import store list from excel";
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
            // 
            // frmSaveLoad
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1100, 733);
            this.Controls.Add(this.btnImport);
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.pnlLoadSession);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.txtSessionName);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.chkSelectAll);
            this.Controls.Add(this.grdStores);
            this.Controls.Add(this.btnNext);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmSaveLoad";
            this.Text = "Pre-Battler";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.grdStores)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.pnlLoadSession.ResumeLayout(false);
            this.pnlLoadSession.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ugrdSession)).EndInit();
            this.mnSessionRightClick.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox cmbAccount;
        private System.Windows.Forms.ComboBox cmbSKULevel;
        private System.Windows.Forms.ComboBox cmbSite;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fILEToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem cLOSEToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem tOOLSToolStripMenuItem;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel sbConnected;
        private System.Windows.Forms.ToolStripStatusLabel sbStoreCount;
        private System.Windows.Forms.Button btnNext;
        private Infragistics.Win.UltraWinGrid.UltraGrid grdStores;
        private System.Windows.Forms.CheckBox chkSelectAll;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox txtSessionName;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ToolStripMenuItem lOADSAVEDSESSIONToolStripMenuItem;
        private System.Windows.Forms.Panel pnlLoadSession;
        private System.Windows.Forms.Label lblClose;
        private System.Windows.Forms.Label label5;
        private Infragistics.Win.UltraWinGrid.UltraGrid ugrdSession;
        private System.Windows.Forms.ToolStripMenuItem sAVELAYOUTToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem sELECTCOLUMNSToolStripMenuItem;
        private System.Windows.Forms.ContextMenuStrip mnSessionRightClick;
        private System.Windows.Forms.ToolStripMenuItem miInactive;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.Button btnImport;
    }
}