namespace WinFormCCCDataSet
{
    partial class frmMainForm
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
            this.btnDoWork = new System.Windows.Forms.Button();
            this.dgvMasterDataSet = new System.Windows.Forms.DataGridView();
            this.lblROCounter = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.lblCurrentROProcessing = new System.Windows.Forms.Label();
            this.btnExportToExcel = new System.Windows.Forms.Button();
            this.lblExportRow = new System.Windows.Forms.Label();
            this.dateTimePickerStart = new System.Windows.Forms.DateTimePicker();
            this.dateTimePickerEnd = new System.Windows.Forms.DateTimePicker();
            this.btnGetLastMonthOnly = new System.Windows.Forms.Button();
            this.lblTo = new System.Windows.Forms.Label();
            this.lblFrom = new System.Windows.Forms.Label();
            this.lblRebuildIndexes = new System.Windows.Forms.Label();
            this.tmrStartUpTasks = new System.Windows.Forms.Timer(this.components);
            this.chkGroupBySalesItem = new System.Windows.Forms.CheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.dgvMasterDataSet)).BeginInit();
            this.SuspendLayout();
            // 
            // btnDoWork
            // 
            this.btnDoWork.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnDoWork.Location = new System.Drawing.Point(12, 486);
            this.btnDoWork.Name = "btnDoWork";
            this.btnDoWork.Size = new System.Drawing.Size(75, 23);
            this.btnDoWork.TabIndex = 0;
            this.btnDoWork.Text = "GET DATA";
            this.btnDoWork.UseVisualStyleBackColor = true;
            this.btnDoWork.Click += new System.EventHandler(this.btnDoWork_Click);
            // 
            // dgvMasterDataSet
            // 
            this.dgvMasterDataSet.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvMasterDataSet.Location = new System.Drawing.Point(12, 41);
            this.dgvMasterDataSet.Name = "dgvMasterDataSet";
            this.dgvMasterDataSet.Size = new System.Drawing.Size(962, 439);
            this.dgvMasterDataSet.TabIndex = 1;
            // 
            // lblROCounter
            // 
            this.lblROCounter.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.lblROCounter.AutoSize = true;
            this.lblROCounter.Location = new System.Drawing.Point(513, 491);
            this.lblROCounter.Name = "lblROCounter";
            this.lblROCounter.Size = new System.Drawing.Size(27, 13);
            this.lblROCounter.TabIndex = 3;
            this.lblROCounter.Text = "[ % ]";
            // 
            // progressBar1
            // 
            this.progressBar1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBar1.Location = new System.Drawing.Point(263, 486);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(244, 23);
            this.progressBar1.TabIndex = 4;
            // 
            // lblCurrentROProcessing
            // 
            this.lblCurrentROProcessing.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblCurrentROProcessing.AutoSize = true;
            this.lblCurrentROProcessing.Location = new System.Drawing.Point(93, 491);
            this.lblCurrentROProcessing.Name = "lblCurrentROProcessing";
            this.lblCurrentROProcessing.Size = new System.Drawing.Size(53, 13);
            this.lblCurrentROProcessing.TabIndex = 5;
            this.lblCurrentROProcessing.Text = "[ Current ]";
            // 
            // btnExportToExcel
            // 
            this.btnExportToExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnExportToExcel.Location = new System.Drawing.Point(736, 486);
            this.btnExportToExcel.Name = "btnExportToExcel";
            this.btnExportToExcel.Size = new System.Drawing.Size(120, 23);
            this.btnExportToExcel.TabIndex = 6;
            this.btnExportToExcel.Text = "Export to Excel";
            this.btnExportToExcel.UseVisualStyleBackColor = true;
            this.btnExportToExcel.Click += new System.EventHandler(this.btnExportToExcel_Click);
            // 
            // lblExportRow
            // 
            this.lblExportRow.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.lblExportRow.AutoSize = true;
            this.lblExportRow.Location = new System.Drawing.Point(862, 491);
            this.lblExportRow.Name = "lblExportRow";
            this.lblExportRow.Size = new System.Drawing.Size(81, 13);
            this.lblExportRow.TabIndex = 7;
            this.lblExportRow.Text = "[ lblExportRow ]";
            // 
            // dateTimePickerStart
            // 
            this.dateTimePickerStart.Location = new System.Drawing.Point(76, 8);
            this.dateTimePickerStart.Name = "dateTimePickerStart";
            this.dateTimePickerStart.Size = new System.Drawing.Size(200, 20);
            this.dateTimePickerStart.TabIndex = 8;
            // 
            // dateTimePickerEnd
            // 
            this.dateTimePickerEnd.Location = new System.Drawing.Point(362, 8);
            this.dateTimePickerEnd.Name = "dateTimePickerEnd";
            this.dateTimePickerEnd.Size = new System.Drawing.Size(200, 20);
            this.dateTimePickerEnd.TabIndex = 9;
            // 
            // btnGetLastMonthOnly
            // 
            this.btnGetLastMonthOnly.Location = new System.Drawing.Point(602, 7);
            this.btnGetLastMonthOnly.Name = "btnGetLastMonthOnly";
            this.btnGetLastMonthOnly.Size = new System.Drawing.Size(75, 23);
            this.btnGetLastMonthOnly.TabIndex = 10;
            this.btnGetLastMonthOnly.Text = "Last Month";
            this.btnGetLastMonthOnly.UseVisualStyleBackColor = true;
            this.btnGetLastMonthOnly.Click += new System.EventHandler(this.btnGetLastMonthOnly_Click);
            // 
            // lblTo
            // 
            this.lblTo.AutoSize = true;
            this.lblTo.Location = new System.Drawing.Point(301, 12);
            this.lblTo.Name = "lblTo";
            this.lblTo.Size = new System.Drawing.Size(55, 13);
            this.lblTo.TabIndex = 11;
            this.lblTo.Text = "End Date:";
            // 
            // lblFrom
            // 
            this.lblFrom.AutoSize = true;
            this.lblFrom.Location = new System.Drawing.Point(12, 12);
            this.lblFrom.Name = "lblFrom";
            this.lblFrom.Size = new System.Drawing.Size(58, 13);
            this.lblFrom.TabIndex = 12;
            this.lblFrom.Text = "Start Date:";
            // 
            // lblRebuildIndexes
            // 
            this.lblRebuildIndexes.AutoSize = true;
            this.lblRebuildIndexes.Location = new System.Drawing.Point(735, 12);
            this.lblRebuildIndexes.Name = "lblRebuildIndexes";
            this.lblRebuildIndexes.Size = new System.Drawing.Size(102, 13);
            this.lblRebuildIndexes.TabIndex = 13;
            this.lblRebuildIndexes.Text = "[ lblRebuildIndexes ]";
            // 
            // tmrStartUpTasks
            // 
            this.tmrStartUpTasks.Tick += new System.EventHandler(this.tmrStartUpTasks_Tick);
            // 
            // chkGroupBySalesItem
            // 
            this.chkGroupBySalesItem.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.chkGroupBySalesItem.AutoSize = true;
            this.chkGroupBySalesItem.Location = new System.Drawing.Point(12, 524);
            this.chkGroupBySalesItem.Name = "chkGroupBySalesItem";
            this.chkGroupBySalesItem.Size = new System.Drawing.Size(426, 17);
            this.chkGroupBySalesItem.TabIndex = 14;
            this.chkGroupBySalesItem.Text = "Group and Report by Sales Item (select this BEFORE clicking the GET DATA button)";
            this.chkGroupBySalesItem.UseVisualStyleBackColor = true;
            this.chkGroupBySalesItem.Visible = false;
            // 
            // frmMainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(986, 553);
            this.Controls.Add(this.chkGroupBySalesItem);
            this.Controls.Add(this.lblRebuildIndexes);
            this.Controls.Add(this.lblFrom);
            this.Controls.Add(this.lblTo);
            this.Controls.Add(this.btnGetLastMonthOnly);
            this.Controls.Add(this.dateTimePickerEnd);
            this.Controls.Add(this.dateTimePickerStart);
            this.Controls.Add(this.lblExportRow);
            this.Controls.Add(this.btnExportToExcel);
            this.Controls.Add(this.lblCurrentROProcessing);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.lblROCounter);
            this.Controls.Add(this.dgvMasterDataSet);
            this.Controls.Add(this.btnDoWork);
            this.Name = "frmMainForm";
            this.Text = "Consolidated Data v6";
            this.Load += new System.EventHandler(this.frmMainForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvMasterDataSet)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnDoWork;
        private System.Windows.Forms.DataGridView dgvMasterDataSet;
        private System.Windows.Forms.Label lblROCounter;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label lblCurrentROProcessing;
        private System.Windows.Forms.Button btnExportToExcel;
        private System.Windows.Forms.Label lblExportRow;
        private System.Windows.Forms.DateTimePicker dateTimePickerStart;
        private System.Windows.Forms.DateTimePicker dateTimePickerEnd;
        private System.Windows.Forms.Button btnGetLastMonthOnly;
        private System.Windows.Forms.Label lblTo;
        private System.Windows.Forms.Label lblFrom;
        private System.Windows.Forms.Label lblRebuildIndexes;
        private System.Windows.Forms.Timer tmrStartUpTasks;
        private System.Windows.Forms.CheckBox chkGroupBySalesItem;
    }
}

