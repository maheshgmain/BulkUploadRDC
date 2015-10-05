namespace UploadApplication
{
    partial class frmUpload
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
            base.Dispose(false);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.tableLayoutPanel3 = new System.Windows.Forms.TableLayoutPanel();
            this.label3 = new System.Windows.Forms.Label();
            this.btnMoveLeft_XlsColumn = new System.Windows.Forms.Button();
            this.btnMoveRight_XlsColumn = new System.Windows.Forms.Button();
            this.btnMoveRight_Database = new System.Windows.Forms.Button();
            this.btnMoveLeft_Database = new System.Windows.Forms.Button();
            this.lstXLSColumn = new System.Windows.Forms.ListBox();
            this.listBox2 = new System.Windows.Forms.ListBox();
            this.listBox3 = new System.Windows.Forms.ListBox();
            this.lstDatabaseColumn = new System.Windows.Forms.ListBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.txtXLSFileName = new System.Windows.Forms.TextBox();
            this.btnGetXLSColumn = new System.Windows.Forms.Button();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.cmbDocumentType = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnClear = new System.Windows.Forms.Button();
            this.btnImportData = new System.Windows.Forms.Button();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.groupBox1.SuspendLayout();
            this.tableLayoutPanel3.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            this.tableLayoutPanel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.BackColor = System.Drawing.SystemColors.ActiveBorder;
            this.groupBox1.Controls.Add(this.progressBar1);
            this.groupBox1.Controls.Add(this.tableLayoutPanel3);
            this.groupBox1.Controls.Add(this.tableLayoutPanel1);
            this.groupBox1.Controls.Add(this.tableLayoutPanel2);
            this.groupBox1.Controls.Add(this.btnClear);
            this.groupBox1.Controls.Add(this.btnImportData);
            this.groupBox1.Location = new System.Drawing.Point(13, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(450, 372);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(25, 319);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(377, 23);
            this.progressBar1.Step = 1;
            this.progressBar1.TabIndex = 1;
            // 
            // tableLayoutPanel3
            // 
            this.tableLayoutPanel3.BackColor = System.Drawing.Color.Gainsboro;
            this.tableLayoutPanel3.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.Single;
            this.tableLayoutPanel3.ColumnCount = 6;
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.tableLayoutPanel3.Controls.Add(this.label3, 0, 0);
            this.tableLayoutPanel3.Controls.Add(this.btnMoveLeft_XlsColumn, 1, 2);
            this.tableLayoutPanel3.Controls.Add(this.btnMoveRight_XlsColumn, 1, 1);
            this.tableLayoutPanel3.Controls.Add(this.btnMoveRight_Database, 4, 2);
            this.tableLayoutPanel3.Controls.Add(this.btnMoveLeft_Database, 4, 1);
            this.tableLayoutPanel3.Controls.Add(this.lstXLSColumn, 0, 1);
            this.tableLayoutPanel3.Controls.Add(this.listBox2, 2, 1);
            this.tableLayoutPanel3.Controls.Add(this.listBox3, 3, 1);
            this.tableLayoutPanel3.Controls.Add(this.lstDatabaseColumn, 5, 1);
            this.tableLayoutPanel3.Controls.Add(this.label4, 2, 0);
            this.tableLayoutPanel3.Controls.Add(this.label6, 4, 0);
            this.tableLayoutPanel3.Location = new System.Drawing.Point(7, 115);
            this.tableLayoutPanel3.Name = "tableLayoutPanel3";
            this.tableLayoutPanel3.RowCount = 3;
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 25.5814F));
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 74.4186F));
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 76F));
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel3.Size = new System.Drawing.Size(438, 165);
            this.tableLayoutPanel3.TabIndex = 1;
            this.tableLayoutPanel3.Visible = false;
            // 
            // label3
            // 
            this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.tableLayoutPanel3.SetColumnSpan(this.label3, 2);
            this.label3.ForeColor = System.Drawing.Color.Maroon;
            this.label3.Location = new System.Drawing.Point(4, 5);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(112, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "XLS Column";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btnMoveLeft_XlsColumn
            // 
            this.btnMoveLeft_XlsColumn.Enabled = false;
            this.btnMoveLeft_XlsColumn.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnMoveLeft_XlsColumn.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.btnMoveLeft_XlsColumn.Location = new System.Drawing.Point(102, 90);
            this.btnMoveLeft_XlsColumn.Name = "btnMoveLeft_XlsColumn";
            this.btnMoveLeft_XlsColumn.Size = new System.Drawing.Size(14, 23);
            this.btnMoveLeft_XlsColumn.TabIndex = 10;
            this.btnMoveLeft_XlsColumn.Text = "<";
            this.btnMoveLeft_XlsColumn.UseVisualStyleBackColor = true;
            this.btnMoveLeft_XlsColumn.Click += new System.EventHandler(this.btnMoveLeft_XlsColumn_Click);
            // 
            // btnMoveRight_XlsColumn
            // 
            this.btnMoveRight_XlsColumn.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.btnMoveRight_XlsColumn.Enabled = false;
            this.btnMoveRight_XlsColumn.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnMoveRight_XlsColumn.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.btnMoveRight_XlsColumn.Location = new System.Drawing.Point(102, 60);
            this.btnMoveRight_XlsColumn.Name = "btnMoveRight_XlsColumn";
            this.btnMoveRight_XlsColumn.Size = new System.Drawing.Size(14, 23);
            this.btnMoveRight_XlsColumn.TabIndex = 8;
            this.btnMoveRight_XlsColumn.Text = ">";
            this.btnMoveRight_XlsColumn.UseVisualStyleBackColor = true;
            this.btnMoveRight_XlsColumn.Click += new System.EventHandler(this.btnMoveRight_XlsColumn_Click);
            // 
            // btnMoveRight_Database
            // 
            this.btnMoveRight_Database.Enabled = false;
            this.btnMoveRight_Database.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnMoveRight_Database.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.btnMoveRight_Database.Location = new System.Drawing.Point(319, 90);
            this.btnMoveRight_Database.Name = "btnMoveRight_Database";
            this.btnMoveRight_Database.Size = new System.Drawing.Size(14, 23);
            this.btnMoveRight_Database.TabIndex = 11;
            this.btnMoveRight_Database.Text = ">";
            this.btnMoveRight_Database.UseVisualStyleBackColor = true;
            this.btnMoveRight_Database.Click += new System.EventHandler(this.btnMoveRight_Database_Click);
            // 
            // btnMoveLeft_Database
            // 
            this.btnMoveLeft_Database.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.btnMoveLeft_Database.Enabled = false;
            this.btnMoveLeft_Database.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnMoveLeft_Database.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.btnMoveLeft_Database.Location = new System.Drawing.Point(319, 60);
            this.btnMoveLeft_Database.Name = "btnMoveLeft_Database";
            this.btnMoveLeft_Database.Size = new System.Drawing.Size(14, 23);
            this.btnMoveLeft_Database.TabIndex = 9;
            this.btnMoveLeft_Database.Text = "<";
            this.btnMoveLeft_Database.UseVisualStyleBackColor = true;
            this.btnMoveLeft_Database.Click += new System.EventHandler(this.btnMoveLeft_Database_Click);
            // 
            // lstXLSColumn
            // 
            this.lstXLSColumn.FormattingEnabled = true;
            this.lstXLSColumn.Location = new System.Drawing.Point(4, 26);
            this.lstXLSColumn.Name = "lstXLSColumn";
            this.tableLayoutPanel3.SetRowSpan(this.lstXLSColumn, 2);
            this.lstXLSColumn.Size = new System.Drawing.Size(91, 134);
            this.lstXLSColumn.TabIndex = 4;
            this.lstXLSColumn.SelectedIndexChanged += new System.EventHandler(this.lstXLSColumn_SelectedIndexChanged);
            // 
            // listBox2
            // 
            this.listBox2.FormattingEnabled = true;
            this.listBox2.Location = new System.Drawing.Point(123, 26);
            this.listBox2.Name = "listBox2";
            this.tableLayoutPanel3.SetRowSpan(this.listBox2, 2);
            this.listBox2.Size = new System.Drawing.Size(91, 134);
            this.listBox2.TabIndex = 5;
            this.listBox2.SelectedIndexChanged += new System.EventHandler(this.listBox2_SelectedIndexChanged);
            // 
            // listBox3
            // 
            this.listBox3.FormattingEnabled = true;
            this.listBox3.Location = new System.Drawing.Point(221, 26);
            this.listBox3.Name = "listBox3";
            this.tableLayoutPanel3.SetRowSpan(this.listBox3, 2);
            this.listBox3.Size = new System.Drawing.Size(91, 134);
            this.listBox3.TabIndex = 6;
            this.listBox3.SelectedIndexChanged += new System.EventHandler(this.listBox3_SelectedIndexChanged);
            // 
            // lstDatabaseColumn
            // 
            this.lstDatabaseColumn.FormattingEnabled = true;
            this.lstDatabaseColumn.Location = new System.Drawing.Point(340, 26);
            this.lstDatabaseColumn.Name = "lstDatabaseColumn";
            this.tableLayoutPanel3.SetRowSpan(this.lstDatabaseColumn, 2);
            this.lstDatabaseColumn.Size = new System.Drawing.Size(93, 134);
            this.lstDatabaseColumn.TabIndex = 7;
            this.lstDatabaseColumn.SelectedIndexChanged += new System.EventHandler(this.lstDatabaseColumn_SelectedIndexChanged);
            // 
            // label4
            // 
            this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.tableLayoutPanel3.SetColumnSpan(this.label4, 2);
            this.label4.ForeColor = System.Drawing.Color.Maroon;
            this.label4.Location = new System.Drawing.Point(123, 5);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(189, 13);
            this.label4.TabIndex = 12;
            this.label4.Text = "Map Columns";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label6
            // 
            this.label6.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.tableLayoutPanel3.SetColumnSpan(this.label6, 2);
            this.label6.ForeColor = System.Drawing.Color.Maroon;
            this.label6.Location = new System.Drawing.Point(319, 5);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(115, 13);
            this.label6.TabIndex = 14;
            this.label6.Text = "Index Column";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.Single;
            this.tableLayoutPanel1.ColumnCount = 3;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 28.13456F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 71.86544F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 139F));
            this.tableLayoutPanel1.Controls.Add(this.btnBrowse, 2, 0);
            this.tableLayoutPanel1.Controls.Add(this.label1, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.txtXLSFileName, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.btnGetXLSColumn, 0, 1);
            this.tableLayoutPanel1.Location = new System.Drawing.Point(7, 47);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 2;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.79365F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 49.20635F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(438, 64);
            this.tableLayoutPanel1.TabIndex = 4;
            // 
            // btnBrowse
            // 
            this.btnBrowse.Location = new System.Drawing.Point(300, 4);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(99, 23);
            this.btnBrowse.TabIndex = 2;
            this.btnBrowse.Text = "Browse";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.label1.Location = new System.Drawing.Point(4, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(76, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "XlS File Name    ";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtXLSFileName
            // 
            this.txtXLSFileName.Location = new System.Drawing.Point(87, 4);
            this.txtXLSFileName.Name = "txtXLSFileName";
            this.txtXLSFileName.Size = new System.Drawing.Size(206, 21);
            this.txtXLSFileName.TabIndex = 1;
            // 
            // btnGetXLSColumn
            // 
            this.tableLayoutPanel1.SetColumnSpan(this.btnGetXLSColumn, 3);
            this.btnGetXLSColumn.Location = new System.Drawing.Point(4, 35);
            this.btnGetXLSColumn.Name = "btnGetXLSColumn";
            this.btnGetXLSColumn.Size = new System.Drawing.Size(429, 25);
            this.btnGetXLSColumn.TabIndex = 3;
            this.btnGetXLSColumn.Text = "Get XlS Column";
            this.btnGetXLSColumn.UseVisualStyleBackColor = true;
            this.btnGetXLSColumn.Click += new System.EventHandler(this.btnGetXLSColumn_Click);
            // 
            // tableLayoutPanel2
            // 
            this.tableLayoutPanel2.BackColor = System.Drawing.Color.Silver;
            this.tableLayoutPanel2.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.Single;
            this.tableLayoutPanel2.ColumnCount = 2;
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 289F));
            this.tableLayoutPanel2.Controls.Add(this.cmbDocumentType, 1, 0);
            this.tableLayoutPanel2.Controls.Add(this.label2, 0, 0);
            this.tableLayoutPanel2.Location = new System.Drawing.Point(7, 12);
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            this.tableLayoutPanel2.RowCount = 1;
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel2.Size = new System.Drawing.Size(437, 28);
            this.tableLayoutPanel2.TabIndex = 2;
            // 
            // cmbDocumentType
            // 
            this.cmbDocumentType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbDocumentType.FormattingEnabled = true;
            this.cmbDocumentType.Location = new System.Drawing.Point(150, 4);
            this.cmbDocumentType.Name = "cmbDocumentType";
            this.cmbDocumentType.Size = new System.Drawing.Size(174, 21);
            this.cmbDocumentType.TabIndex = 0;
            this.cmbDocumentType.SelectedIndexChanged += new System.EventHandler(this.cmbDocumentType_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(4, 7);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(139, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Document Type ";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // btnClear
            // 
            this.btnClear.Location = new System.Drawing.Point(243, 284);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(133, 29);
            this.btnClear.TabIndex = 1;
            this.btnClear.Text = "Clear And Browse";
            this.btnClear.UseVisualStyleBackColor = true;
            this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
            // 
            // btnImportData
            // 
            this.btnImportData.Location = new System.Drawing.Point(104, 284);
            this.btnImportData.Name = "btnImportData";
            this.btnImportData.Size = new System.Drawing.Size(133, 29);
            this.btnImportData.TabIndex = 1;
            this.btnImportData.Text = "Import Data";
            this.btnImportData.UseVisualStyleBackColor = true;
            this.btnImportData.Click += new System.EventHandler(this.btnImportData_Click);
            // 
            // frmUpload
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.InactiveBorder;
            this.ClientSize = new System.Drawing.Size(475, 396);
            this.Controls.Add(this.groupBox1);
            this.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ForeColor = System.Drawing.Color.Black;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmUpload";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Upload Application ";
            this.Load += new System.EventHandler(this.frmUpload_Load);
            this.groupBox1.ResumeLayout(false);
            this.tableLayoutPanel3.ResumeLayout(false);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.tableLayoutPanel2.ResumeLayout(false);
            this.tableLayoutPanel2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtXLSFileName;
        private System.Windows.Forms.Button btnGetXLSColumn;
        private System.Windows.Forms.ListBox lstXLSColumn;
        private System.Windows.Forms.ListBox listBox2;
        private System.Windows.Forms.ListBox listBox3;
        private System.Windows.Forms.ListBox lstDatabaseColumn;
        private System.Windows.Forms.Button btnMoveRight_XlsColumn;
        private System.Windows.Forms.Button btnMoveLeft_Database;
        private System.Windows.Forms.Button btnImportData;
        private System.Windows.Forms.Button btnMoveLeft_XlsColumn;
        private System.Windows.Forms.Button btnMoveRight_Database;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel2;
        private System.Windows.Forms.ComboBox cmbDocumentType;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel3;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button btnClear;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
    }
}

