namespace XingManager
{
    partial class XingForm
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.DataGridView gridCrossings;
        private System.Windows.Forms.Button btnRescan;
        private System.Windows.Forms.Button btnApply;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Button btnInsert;
        private System.Windows.Forms.Button btnDelete;
        private System.Windows.Forms.Button btnRenumber;
        private System.Windows.Forms.Button btnGeneratePage;
        private System.Windows.Forms.Button btnLatLong;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.Button btnImport;
        private System.Windows.Forms.FlowLayoutPanel buttonPanel;

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
            this.btnAdd = new System.Windows.Forms.Button();
            this.btnInsert = new System.Windows.Forms.Button();
            this.btnDelete = new System.Windows.Forms.Button();
            this.btnRenumber = new System.Windows.Forms.Button();
            this.btnGeneratePage = new System.Windows.Forms.Button();
            this.btnLatLong = new System.Windows.Forms.Button();
            this.btnExport = new System.Windows.Forms.Button();
            this.btnImport = new System.Windows.Forms.Button();
            this.buttonPanel = new System.Windows.Forms.FlowLayoutPanel();
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
            this.gridCrossings.Location = new System.Drawing.Point(0, 35);
            this.gridCrossings.MultiSelect = false;
            this.gridCrossings.Name = "gridCrossings";
            this.gridCrossings.RowHeadersVisible = false;
            this.gridCrossings.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.gridCrossings.Size = new System.Drawing.Size(900, 515);
            this.gridCrossings.TabIndex = 0;
            // 
            // btnRescan
            // 
            this.btnRescan.Location = new System.Drawing.Point(3, 3);
            this.btnRescan.Name = "btnRescan";
            this.btnRescan.Size = new System.Drawing.Size(75, 25);
            this.btnRescan.TabIndex = 0;
            this.btnRescan.Text = "Rescan";
            this.btnRescan.UseVisualStyleBackColor = true;
            this.btnRescan.Click += new System.EventHandler(this.btnRescan_Click);
            // 
            // btnApply
            // 
            this.btnApply.Location = new System.Drawing.Point(84, 3);
            this.btnApply.Name = "btnApply";
            this.btnApply.Size = new System.Drawing.Size(100, 25);
            this.btnApply.TabIndex = 1;
            this.btnApply.Text = "Apply to Drawing";
            this.btnApply.UseVisualStyleBackColor = true;
            this.btnApply.Click += new System.EventHandler(this.btnApply_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(190, 3);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(75, 25);
            this.btnAdd.TabIndex = 2;
            this.btnAdd.Text = "Add";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnInsert
            // 
            this.btnInsert.Location = new System.Drawing.Point(271, 3);
            this.btnInsert.Name = "btnInsert";
            this.btnInsert.Size = new System.Drawing.Size(90, 25);
            this.btnInsert.TabIndex = 3;
            this.btnInsert.Text = "Insert at...";
            this.btnInsert.UseVisualStyleBackColor = true;
            this.btnInsert.Click += new System.EventHandler(this.btnInsert_Click);
            // 
            // btnDelete
            // 
            this.btnDelete.Location = new System.Drawing.Point(367, 3);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(110, 25);
            this.btnDelete.TabIndex = 4;
            this.btnDelete.Text = "Delete Selected";
            this.btnDelete.UseVisualStyleBackColor = true;
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // btnRenumber
            // 
            this.btnRenumber.Location = new System.Drawing.Point(483, 3);
            this.btnRenumber.Name = "btnRenumber";
            this.btnRenumber.Size = new System.Drawing.Size(85, 25);
            this.btnRenumber.TabIndex = 5;
            this.btnRenumber.Text = "Renumber";
            this.btnRenumber.UseVisualStyleBackColor = true;
            this.btnRenumber.Click += new System.EventHandler(this.btnRenumber_Click);
            // 
            // btnGeneratePage
            // 
            this.btnGeneratePage.Location = new System.Drawing.Point(574, 3);
            this.btnGeneratePage.Name = "btnGeneratePage";
            this.btnGeneratePage.Size = new System.Drawing.Size(120, 25);
            this.btnGeneratePage.TabIndex = 6;
            this.btnGeneratePage.Text = "Generate XING PAGE";
            this.btnGeneratePage.UseVisualStyleBackColor = true;
            this.btnGeneratePage.Click += new System.EventHandler(this.btnGeneratePage_Click);
            // 
            // btnLatLong
            // 
            this.btnLatLong.Location = new System.Drawing.Point(700, 3);
            this.btnLatLong.Name = "btnLatLong";
            this.btnLatLong.Size = new System.Drawing.Size(120, 25);
            this.btnLatLong.TabIndex = 7;
            this.btnLatLong.Text = "Create LAT/LONG";
            this.btnLatLong.UseVisualStyleBackColor = true;
            this.btnLatLong.Click += new System.EventHandler(this.btnLatLong_Click);
            // 
            // btnExport
            // 
            this.btnExport.Location = new System.Drawing.Point(826, 3);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(75, 25);
            this.btnExport.TabIndex = 8;
            this.btnExport.Text = "Export";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // btnImport
            // 
            this.btnImport.Location = new System.Drawing.Point(907, 3);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(75, 25);
            this.btnImport.TabIndex = 9;
            this.btnImport.Text = "Import";
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
            // 
            // buttonPanel
            // 
            this.buttonPanel.AutoSize = true;
            this.buttonPanel.AutoScroll = true;
            this.buttonPanel.Controls.Add(this.btnRescan);
            this.buttonPanel.Controls.Add(this.btnApply);
            this.buttonPanel.Controls.Add(this.btnAdd);
            this.buttonPanel.Controls.Add(this.btnInsert);
            this.buttonPanel.Controls.Add(this.btnDelete);
            this.buttonPanel.Controls.Add(this.btnRenumber);
            this.buttonPanel.Controls.Add(this.btnGeneratePage);
            this.buttonPanel.Controls.Add(this.btnLatLong);
            this.buttonPanel.Controls.Add(this.btnExport);
            this.buttonPanel.Controls.Add(this.btnImport);
            this.buttonPanel.Dock = System.Windows.Forms.DockStyle.Top;
            this.buttonPanel.Location = new System.Drawing.Point(0, 0);
            this.buttonPanel.Name = "buttonPanel";
            this.buttonPanel.Size = new System.Drawing.Size(900, 35);
            this.buttonPanel.TabIndex = 1;
            this.buttonPanel.WrapContents = false;
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
