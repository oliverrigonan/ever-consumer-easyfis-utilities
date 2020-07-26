namespace EverConsumerUtilities.Forms
{
    partial class SalesInvoiceForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SalesInvoiceForm));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.imageListMenu = new System.Windows.Forms.ImageList(this.components);
            this.panel1 = new System.Windows.Forms.Panel();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.label1 = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.panel3 = new System.Windows.Forms.Panel();
            this.label9 = new System.Windows.Forms.Label();
            this.textBoxFilter = new System.Windows.Forms.TextBox();
            this.buttonGet = new System.Windows.Forms.Button();
            this.dataGridViewSalesInvoice = new System.Windows.Forms.DataGridView();
            this.label8 = new System.Windows.Forms.Label();
            this.comboBoxBranch = new System.Windows.Forms.ComboBox();
            this.label7 = new System.Windows.Forms.Label();
            this.dateTimePickerEndDate = new System.Windows.Forms.DateTimePicker();
            this.label6 = new System.Windows.Forms.Label();
            this.dateTimePickerStartDate = new System.Windows.Forms.DateTimePicker();
            this.ColumnButtonPrint = new System.Windows.Forms.DataGridViewButtonColumn();
            this.ColumnId = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnSINumber = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnSIDate = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnSIReferenceNumber = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnCustomer = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnRemarks = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnDocumentReference = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnAmount = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnIsLocked = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.panel2.SuspendLayout();
            this.panel3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewSalesInvoice)).BeginInit();
            this.SuspendLayout();
            // 
            // imageListMenu
            // 
            this.imageListMenu.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageListMenu.ImageStream")));
            this.imageListMenu.TransparentColor = System.Drawing.Color.Transparent;
            this.imageListMenu.Images.SetKeyName(0, "menu.png");
            this.imageListMenu.Images.SetKeyName(1, "sales_invoice.png");
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.White;
            this.panel1.Controls.Add(this.pictureBox1);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1173, 54);
            this.panel1.TabIndex = 3;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::EverConsumerUtilities.Properties.Resources.sales_invoice;
            this.pictureBox1.Location = new System.Drawing.Point(9, 9);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(33, 34);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 15F, System.Drawing.FontStyle.Bold);
            this.label1.Location = new System.Drawing.Point(41, 11);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(134, 28);
            this.label1.TabIndex = 0;
            this.label1.Text = "Sales Invoice";
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(17)))), ((int)(((byte)(37)))), ((int)(((byte)(64)))));
            this.panel2.Controls.Add(this.label5);
            this.panel2.Controls.Add(this.label4);
            this.panel2.Controls.Add(this.label3);
            this.panel2.Controls.Add(this.label2);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel2.Location = new System.Drawing.Point(0, 403);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1173, 87);
            this.panel2.TabIndex = 5;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.label5.ForeColor = System.Drawing.Color.White;
            this.label5.Location = new System.Drawing.Point(6, 32);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(160, 15);
            this.label5.TabIndex = 3;
            this.label5.Text = "Version: 1.202007262136.NOR";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.label4.ForeColor = System.Drawing.Color.White;
            this.label4.Location = new System.Drawing.Point(6, 62);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(184, 15);
            this.label4.TabIndex = 2;
            this.label4.Text = "© 2018 Easyfis | www.easyfis.com";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.label3.ForeColor = System.Drawing.Color.White;
            this.label3.Location = new System.Drawing.Point(6, 47);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(168, 15);
            this.label3.TabIndex = 1;
            this.label3.Text = "Developer: Easyfis Corporation";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(5, 7);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(224, 20);
            this.label2.TabIndex = 0;
            this.label2.Text = "Ever Consumer Easyfis Utilities";
            // 
            // panel3
            // 
            this.panel3.BackColor = System.Drawing.SystemColors.Control;
            this.panel3.Controls.Add(this.label9);
            this.panel3.Controls.Add(this.textBoxFilter);
            this.panel3.Controls.Add(this.buttonGet);
            this.panel3.Controls.Add(this.dataGridViewSalesInvoice);
            this.panel3.Controls.Add(this.label8);
            this.panel3.Controls.Add(this.comboBoxBranch);
            this.panel3.Controls.Add(this.label7);
            this.panel3.Controls.Add(this.dateTimePickerEndDate);
            this.panel3.Controls.Add(this.label6);
            this.panel3.Controls.Add(this.dateTimePickerStartDate);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel3.Location = new System.Drawing.Point(0, 54);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(1173, 349);
            this.panel3.TabIndex = 6;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(525, 3);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(42, 20);
            this.label9.TabIndex = 9;
            this.label9.Text = "Filter";
            // 
            // textBoxFilter
            // 
            this.textBoxFilter.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxFilter.Location = new System.Drawing.Point(529, 26);
            this.textBoxFilter.Name = "textBoxFilter";
            this.textBoxFilter.Size = new System.Drawing.Size(546, 27);
            this.textBoxFilter.TabIndex = 8;
            // 
            // buttonGet
            // 
            this.buttonGet.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonGet.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(92)))), ((int)(((byte)(184)))), ((int)(((byte)(92)))));
            this.buttonGet.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonGet.ForeColor = System.Drawing.Color.White;
            this.buttonGet.Location = new System.Drawing.Point(1081, 4);
            this.buttonGet.Name = "buttonGet";
            this.buttonGet.Size = new System.Drawing.Size(84, 50);
            this.buttonGet.TabIndex = 7;
            this.buttonGet.Text = "Get";
            this.buttonGet.UseVisualStyleBackColor = false;
            this.buttonGet.Click += new System.EventHandler(this.buttonGet_Click);
            // 
            // dataGridViewSalesInvoice
            // 
            this.dataGridViewSalesInvoice.AllowUserToAddRows = false;
            this.dataGridViewSalesInvoice.AllowUserToDeleteRows = false;
            this.dataGridViewSalesInvoice.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridViewSalesInvoice.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewSalesInvoice.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ColumnButtonPrint,
            this.ColumnId,
            this.ColumnSINumber,
            this.ColumnSIDate,
            this.ColumnSIReferenceNumber,
            this.ColumnCustomer,
            this.ColumnRemarks,
            this.ColumnDocumentReference,
            this.ColumnAmount,
            this.ColumnIsLocked});
            this.dataGridViewSalesInvoice.Location = new System.Drawing.Point(9, 59);
            this.dataGridViewSalesInvoice.Name = "dataGridViewSalesInvoice";
            this.dataGridViewSalesInvoice.ReadOnly = true;
            this.dataGridViewSalesInvoice.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridViewSalesInvoice.Size = new System.Drawing.Size(1156, 284);
            this.dataGridViewSalesInvoice.TabIndex = 6;
            this.dataGridViewSalesInvoice.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewSalesInvoice_CellClick);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(285, 3);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(54, 20);
            this.label8.TabIndex = 5;
            this.label8.Text = "Branch";
            // 
            // comboBoxBranch
            // 
            this.comboBoxBranch.FormattingEnabled = true;
            this.comboBoxBranch.Location = new System.Drawing.Point(289, 26);
            this.comboBoxBranch.Name = "comboBoxBranch";
            this.comboBoxBranch.Size = new System.Drawing.Size(234, 28);
            this.comboBoxBranch.TabIndex = 4;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(147, 3);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(70, 20);
            this.label7.TabIndex = 3;
            this.label7.Text = "End Date";
            // 
            // dateTimePickerEndDate
            // 
            this.dateTimePickerEndDate.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dateTimePickerEndDate.Location = new System.Drawing.Point(151, 26);
            this.dateTimePickerEndDate.Name = "dateTimePickerEndDate";
            this.dateTimePickerEndDate.Size = new System.Drawing.Size(132, 27);
            this.dateTimePickerEndDate.TabIndex = 2;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(5, 3);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(76, 20);
            this.label6.TabIndex = 1;
            this.label6.Text = "Start Date";
            // 
            // dateTimePickerStartDate
            // 
            this.dateTimePickerStartDate.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dateTimePickerStartDate.Location = new System.Drawing.Point(9, 26);
            this.dateTimePickerStartDate.Name = "dateTimePickerStartDate";
            this.dateTimePickerStartDate.Size = new System.Drawing.Size(132, 27);
            this.dateTimePickerStartDate.TabIndex = 0;
            // 
            // ColumnButtonPrint
            // 
            this.ColumnButtonPrint.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.ColumnButtonPrint.HeaderText = "Print";
            this.ColumnButtonPrint.Name = "ColumnButtonPrint";
            this.ColumnButtonPrint.ReadOnly = true;
            this.ColumnButtonPrint.Width = 80;
            // 
            // ColumnId
            // 
            this.ColumnId.HeaderText = "Id";
            this.ColumnId.Name = "ColumnId";
            this.ColumnId.ReadOnly = true;
            this.ColumnId.Visible = false;
            // 
            // ColumnSINumber
            // 
            this.ColumnSINumber.HeaderText = "SI No.";
            this.ColumnSINumber.Name = "ColumnSINumber";
            this.ColumnSINumber.ReadOnly = true;
            // 
            // ColumnSIDate
            // 
            this.ColumnSIDate.HeaderText = "SI Date";
            this.ColumnSIDate.Name = "ColumnSIDate";
            this.ColumnSIDate.ReadOnly = true;
            // 
            // ColumnSIReferenceNumber
            // 
            this.ColumnSIReferenceNumber.HeaderText = "SI Ref. No.";
            this.ColumnSIReferenceNumber.Name = "ColumnSIReferenceNumber";
            this.ColumnSIReferenceNumber.ReadOnly = true;
            this.ColumnSIReferenceNumber.Width = 150;
            // 
            // ColumnCustomer
            // 
            this.ColumnCustomer.HeaderText = "Customer";
            this.ColumnCustomer.Name = "ColumnCustomer";
            this.ColumnCustomer.ReadOnly = true;
            this.ColumnCustomer.Width = 200;
            // 
            // ColumnRemarks
            // 
            this.ColumnRemarks.HeaderText = "Remarks";
            this.ColumnRemarks.Name = "ColumnRemarks";
            this.ColumnRemarks.ReadOnly = true;
            this.ColumnRemarks.Width = 150;
            // 
            // ColumnDocumentReference
            // 
            this.ColumnDocumentReference.HeaderText = "Doc. Ref.";
            this.ColumnDocumentReference.Name = "ColumnDocumentReference";
            this.ColumnDocumentReference.ReadOnly = true;
            this.ColumnDocumentReference.Width = 150;
            // 
            // ColumnAmount
            // 
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            this.ColumnAmount.DefaultCellStyle = dataGridViewCellStyle1;
            this.ColumnAmount.HeaderText = "Amount";
            this.ColumnAmount.Name = "ColumnAmount";
            this.ColumnAmount.ReadOnly = true;
            this.ColumnAmount.Width = 120;
            // 
            // ColumnIsLocked
            // 
            this.ColumnIsLocked.HeaderText = "L";
            this.ColumnIsLocked.Name = "ColumnIsLocked";
            this.ColumnIsLocked.ReadOnly = true;
            this.ColumnIsLocked.Width = 50;
            // 
            // SalesInvoiceForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(1173, 490);
            this.Controls.Add(this.panel3);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.panel2);
            this.Font = new System.Drawing.Font("Segoe UI", 11.25F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "SalesInvoiceForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Sales Invoice";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.SalesInvoiceForm_FormClosing);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewSalesInvoice)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.ImageList imageListMenu;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox textBoxFilter;
        private System.Windows.Forms.Button buttonGet;
        private System.Windows.Forms.DataGridView dataGridViewSalesInvoice;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.ComboBox comboBoxBranch;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.DateTimePicker dateTimePickerEndDate;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.DateTimePicker dateTimePickerStartDate;
        private System.Windows.Forms.DataGridViewButtonColumn ColumnButtonPrint;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnId;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnSINumber;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnSIDate;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnSIReferenceNumber;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnCustomer;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnRemarks;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnDocumentReference;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnAmount;
        private System.Windows.Forms.DataGridViewCheckBoxColumn ColumnIsLocked;
    }
}