namespace XingManager
{
    partial class XingForm
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.DataGridView gridCrossings;
        private System.Windows.Forms.Button btnRescan;
        private System.Windows.Forms.Button btnApply;
        private System.Windows.Forms.Button btnDelete;
        private System.Windows.Forms.Button btnRenumber;
        private System.Windows.Forms.Button btnAddRncPolyline;
        private System.Windows.Forms.Button btnGeneratePage;
        private System.Windows.Forms.Button btnGenerateAllLatLongTables;
        private System.Windows.Forms.Button btnGenerateOtherLatLongTables;
        private System.Windows.Forms.Button btnLatLong;
        private System.Windows.Forms.Button btnAddLatLong;
        private System.Windows.Forms.Button btnMatchTable;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.Button btnImport;
        private System.Windows.Forms.FlowLayoutPanel buttonPanel;
        private System.Windows.Forms.Label lblUtmZone;
        private System.Windows.Forms.ComboBox cmbUtmZone;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.gridCrossings = new System.Windows.Forms.DataGridView();
            this.btnRescan = new System.Windows.Forms.Button();
            this.btnApply = new System.Windows.Forms.Button();
            this.btnDelete = new System.Windows.Forms.Button();
            this.btnRenumber = new System.Windows.Forms.Button();
            this.btnAddRncPolyline = new System.Windows.Forms.Button();
            this.btnGeneratePage = new System.Windows.Forms.Button();
            this.btnGenerateAllLatLongTables = new System.Windows.Forms.Button();
            this.btnGenerateOtherLatLongTables = new System.Windows.Forms.Button();
            this.btnLatLong = new System.Windows.Forms.Button();
            this.btnAddLatLong = new System.Windows.Forms.Button();
            this.btnMatchTable = new System.Windows.Forms.Button();
            this.btnExport = new System.Windows.Forms.Button();
            this.btnImport = new System.Windows.Forms.Button();
            this.buttonPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.lblUtmZone = new System.Windows.Forms.Label();
            this.cmbUtmZone = new System.Windows.Forms.ComboBox();
            ((System.ComponentModel.ISupportInitialize)(this.gridCrossings)).BeginInit();
            this.buttonPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // gridCrossings
            // 
            this.gridCrossings.AllowUserToAddRows = false;
            this.gridCrossings.AllowUserToDeleteRows = false;
            this.gridCrossings.AllowUserToResizeRows = false;
            this.gridCrossings.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.gridCrossings.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gridCrossings.Dock = System.Windows.Forms.DockStyle.Fill;
            this.gridCrossings.Location = new System.Drawing.Point(0, 47);
            this.gridCrossings.MultiSelect = false;
            this.gridCrossings.Name = "gridCrossings";
            this.gridCrossings.RowHeadersVisible = false;
            this.gridCrossings.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.gridCrossings.Size = new System.Drawing.Size(900, 503);
            this.gridCrossings.TabIndex = 0;
            // 
            // btnRescan
            // 
            this.btnRescan.Location = new System.Drawing.Point(3, 3);
            this.btnRescan.Name = "btnRescan";
            this.btnRescan.Size = new System.Drawing.Size(75, 32);
            this.btnRescan.TabIndex = 0;
            this.btnRescan.Text = "Scan";
            this.btnRescan.UseVisualStyleBackColor = true;
            this.btnRescan.Click += new System.EventHandler(this.btnRescan_Click);
            // 
            // btnApply
            // 
            this.btnApply.Location = new System.Drawing.Point(84, 3);
            this.btnApply.Name = "btnApply";
            this.btnApply.Size = new System.Drawing.Size(100, 32);
            this.btnApply.TabIndex = 1;
            this.btnApply.Text = "Apply to Drawing";
            this.btnApply.UseVisualStyleBackColor = true;
            this.btnApply.Click += new System.EventHandler(this.btnApply_Click);
            // 
            // btnDelete
            //
            this.btnDelete.Location = new System.Drawing.Point(190, 3);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(110, 32);
            this.btnDelete.TabIndex = 2;
            this.btnDelete.Text = "Delete Selected";
            this.btnDelete.UseVisualStyleBackColor = true;
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            //
            // btnRenumber
            //
            this.btnRenumber.Location = new System.Drawing.Point(306, 3);
            this.btnRenumber.Name = "btnRenumber";
            this.btnRenumber.Size = new System.Drawing.Size(85, 32);
            this.btnRenumber.TabIndex = 3;
            this.btnRenumber.Text = "Renumber";
            this.btnRenumber.UseVisualStyleBackColor = true;
            this.btnRenumber.Click += new System.EventHandler(this.btnRenumber_Click);
            //
            // btnAddRncPolyline
            //
            this.btnAddRncPolyline.Location = new System.Drawing.Point(397, 3);
            this.btnAddRncPolyline.Name = "btnAddRncPolyline";
            this.btnAddRncPolyline.Size = new System.Drawing.Size(90, 32);
            this.btnAddRncPolyline.TabIndex = 4;
            this.btnAddRncPolyline.Text = "Add RNC PL";
            this.btnAddRncPolyline.UseVisualStyleBackColor = true;
            this.btnAddRncPolyline.Click += new System.EventHandler(this.btnAddRncPolyline_Click);
            //
            // btnGeneratePage
            //
            this.btnGeneratePage.Location = new System.Drawing.Point(493, 3);
            this.btnGeneratePage.Name = "btnGeneratePage";
            this.btnGeneratePage.Size = new System.Drawing.Size(120, 32);
            this.btnGeneratePage.TabIndex = 5;
            this.btnGeneratePage.Text = "Generate XING PAGE";
            this.btnGeneratePage.UseVisualStyleBackColor = true;
            this.btnGeneratePage.Click += new System.EventHandler(this.btnGeneratePage_Click);
            //
            // btnGenerateAllLatLongTables
            //
            this.btnGenerateAllLatLongTables.Location = new System.Drawing.Point(619, 3);
            this.btnGenerateAllLatLongTables.Name = "btnGenerateAllLatLongTables";
            this.btnGenerateAllLatLongTables.Size = new System.Drawing.Size(160, 32);
            this.btnGenerateAllLatLongTables.TabIndex = 6;
            this.btnGenerateAllLatLongTables.Text = "Generate WATER LAT/LONGS";
            this.btnGenerateAllLatLongTables.UseVisualStyleBackColor = true;
            this.btnGenerateAllLatLongTables.Click += new System.EventHandler(this.btnGenerateAllLatLongTables_Click);
            //
            // btnGenerateOtherLatLongTables
            //
            this.btnGenerateOtherLatLongTables.Location = new System.Drawing.Point(785, 3);
            this.btnGenerateOtherLatLongTables.Name = "btnGenerateOtherLatLongTables";
            this.btnGenerateOtherLatLongTables.Size = new System.Drawing.Size(170, 32);
            this.btnGenerateOtherLatLongTables.TabIndex = 7;
            this.btnGenerateOtherLatLongTables.Text = "Generate OTHER LAT/LONGS";
            this.btnGenerateOtherLatLongTables.UseVisualStyleBackColor = true;
            this.btnGenerateOtherLatLongTables.Click += new System.EventHandler(this.btnGenerateOtherLatLongTables_Click);
            //
            // btnLatLong
            //
            this.btnLatLong.Location = new System.Drawing.Point(961, 3);
            this.btnLatLong.Name = "btnLatLong";
            this.btnLatLong.Size = new System.Drawing.Size(120, 32);
            this.btnLatLong.TabIndex = 8;
            this.btnLatLong.Text = "Create LAT/LONG";
            this.btnLatLong.UseVisualStyleBackColor = true;
            this.btnLatLong.Click += new System.EventHandler(this.btnLatLong_Click);
            //
            // btnAddLatLong
            //
            this.btnAddLatLong.Location = new System.Drawing.Point(1087, 3);
            this.btnAddLatLong.Name = "btnAddLatLong";
            this.btnAddLatLong.Size = new System.Drawing.Size(120, 32);
            this.btnAddLatLong.TabIndex = 9;
            this.btnAddLatLong.Text = "Add LAT/LONG";
            this.btnAddLatLong.UseVisualStyleBackColor = true;
            this.btnAddLatLong.Click += new System.EventHandler(this.btnAddLatLong_Click);
            //
            // btnMatchTable
            //
            this.btnMatchTable.Location = new System.Drawing.Point(1213, 3);
            this.btnMatchTable.Name = "btnMatchTable";
            this.btnMatchTable.Size = new System.Drawing.Size(120, 32);
            this.btnMatchTable.TabIndex = 10;
            this.btnMatchTable.Text = "Match Table";
            this.btnMatchTable.UseVisualStyleBackColor = true;
            this.btnMatchTable.Click += new System.EventHandler(this.btnMatchTable_Click);
            //
            // btnExport
            //
            this.btnExport.Location = new System.Drawing.Point(1339, 3);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(75, 32);
            this.btnExport.TabIndex = 11;
            this.btnExport.Text = "Export";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            //
            // btnImport
            //
            this.btnImport.Location = new System.Drawing.Point(1420, 3);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(75, 32);
            this.btnImport.TabIndex = 12;
            this.btnImport.Text = "Import";
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
            // 
            // buttonPanel
            // 
            this.buttonPanel.AutoSize = false;
            this.buttonPanel.AutoScroll = true;
            this.buttonPanel.Controls.Add(this.btnRescan);
            this.buttonPanel.Controls.Add(this.btnApply);
            this.buttonPanel.Controls.Add(this.btnDelete);
            this.buttonPanel.Controls.Add(this.btnRenumber);
            this.buttonPanel.Controls.Add(this.btnAddRncPolyline);
            this.buttonPanel.Controls.Add(this.btnGeneratePage);
            this.buttonPanel.Controls.Add(this.btnGenerateAllLatLongTables);
            this.buttonPanel.Controls.Add(this.btnGenerateOtherLatLongTables);
            this.buttonPanel.Controls.Add(this.btnLatLong);
            this.buttonPanel.Controls.Add(this.btnAddLatLong);
            this.buttonPanel.Controls.Add(this.btnMatchTable);
            this.buttonPanel.Controls.Add(this.btnExport);
            this.buttonPanel.Controls.Add(this.btnImport);
            this.buttonPanel.Controls.Add(this.lblUtmZone);
            this.buttonPanel.Controls.Add(this.cmbUtmZone);
            this.buttonPanel.Dock = System.Windows.Forms.DockStyle.Top;
            this.buttonPanel.Location = new System.Drawing.Point(0, 0);
            this.buttonPanel.Name = "buttonPanel";
            this.buttonPanel.Size = new System.Drawing.Size(900, 47);
            this.buttonPanel.TabIndex = 1;
            this.buttonPanel.WrapContents = false;
            //
            // lblUtmZone
            //
            this.lblUtmZone.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.lblUtmZone.AutoSize = true;
            this.lblUtmZone.Location = new System.Drawing.Point(1325, 13);
            this.lblUtmZone.Margin = new System.Windows.Forms.Padding(3, 10, 3, 0);
            this.lblUtmZone.Name = "lblUtmZone";
            this.lblUtmZone.Size = new System.Drawing.Size(64, 13);
            this.lblUtmZone.TabIndex = 13;
            this.lblUtmZone.Text = "UTM Zone:";
            //
            // cmbUtmZone
            //
            this.cmbUtmZone.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbUtmZone.FormattingEnabled = true;
            this.cmbUtmZone.Items.AddRange(new object[] {
            "11",
            "12"});
            this.cmbUtmZone.Location = new System.Drawing.Point(1395, 10);
            this.cmbUtmZone.Margin = new System.Windows.Forms.Padding(3, 7, 3, 3);
            this.cmbUtmZone.Name = "cmbUtmZone";
            this.cmbUtmZone.Size = new System.Drawing.Size(60, 21);
            this.cmbUtmZone.TabIndex = 14;
            this.cmbUtmZone.SelectedIndexChanged += new System.EventHandler(this.cmbUtmZone_SelectedIndexChanged);
            //
            // XingForm
            //
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.gridCrossings);
            this.Controls.Add(this.buttonPanel);
            this.Name = "XingForm";
            this.Size = new System.Drawing.Size(900, 550);
            ((System.ComponentModel.ISupportInitialize)(this.gridCrossings)).EndInit();
            this.buttonPanel.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
