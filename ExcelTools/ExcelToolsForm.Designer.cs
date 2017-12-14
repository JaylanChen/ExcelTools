namespace ExcelTools
{
    partial class ExcelToolsForm
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ExcelToolsForm));
            this.ExcelToolsStatus = new System.Windows.Forms.StatusStrip();
            this.ExcelStatusProgressBar = new System.Windows.Forms.ToolStripProgressBar();
            this.copyrightLab = new System.Windows.Forms.ToolStripStatusLabel();
            this.ExcelTabCtrol = new System.Windows.Forms.TabControl();
            this.SplitSheetTabPage = new System.Windows.Forms.TabPage();
            this.BtnSplitExcel = new System.Windows.Forms.Button();
            this.BtnSelectFile = new System.Windows.Forms.Button();
            this.BtnSelectFolder = new System.Windows.Forms.Button();
            this.MergerSheetTabPage = new System.Windows.Forms.TabPage();
            this.CheckBox_OneSheet = new System.Windows.Forms.CheckBox();
            this.MergerSheetBtn = new System.Windows.Forms.Button();
            this.CheckBox_OneFile = new System.Windows.Forms.CheckBox();
            this.MergerSheetSaveFolderBtn = new System.Windows.Forms.Button();
            this.MergerSheetFolderBtn = new System.Windows.Forms.Button();
            this.ExcelToolTip = new System.Windows.Forms.ToolTip(this.components);
            this.ExcelToolsStatus.SuspendLayout();
            this.ExcelTabCtrol.SuspendLayout();
            this.SplitSheetTabPage.SuspendLayout();
            this.MergerSheetTabPage.SuspendLayout();
            this.SuspendLayout();
            // 
            // ExcelToolsStatus
            // 
            this.ExcelToolsStatus.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.ExcelToolsStatus.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.copyrightLab,
            this.ExcelStatusProgressBar});
            this.ExcelToolsStatus.Location = new System.Drawing.Point(0, 161);
            this.ExcelToolsStatus.Name = "ExcelToolsStatus";
            this.ExcelToolsStatus.Padding = new System.Windows.Forms.Padding(16, 0, 1, 0);
            this.ExcelToolsStatus.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.ExcelToolsStatus.Size = new System.Drawing.Size(684, 28);
            this.ExcelToolsStatus.TabIndex = 2;
            this.ExcelToolsStatus.Text = "ExcelToolsStatus";
            // 
            // ExcelStatusProgressBar
            // 
            this.ExcelStatusProgressBar.ForeColor = System.Drawing.Color.LawnGreen;
            this.ExcelStatusProgressBar.Margin = new System.Windows.Forms.Padding(0, 3, 1, 3);
            this.ExcelStatusProgressBar.Name = "ExcelStatusProgressBar";
            this.ExcelStatusProgressBar.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.ExcelStatusProgressBar.Size = new System.Drawing.Size(570, 22);
            this.ExcelStatusProgressBar.Step = 1;
            this.ExcelStatusProgressBar.ToolTipText = "进度";
            // 
            // copyrightLab
            // 
            this.copyrightLab.Margin = new System.Windows.Forms.Padding(5, 2, 5, 2);
            this.copyrightLab.Name = "copyrightLab";
            this.copyrightLab.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.copyrightLab.Size = new System.Drawing.Size(87, 24);
            this.copyrightLab.Text = "©2017 Jaylan";
            this.copyrightLab.ToolTipText = "Jaylan";
            // 
            // ExcelTabCtrol
            // 
            this.ExcelTabCtrol.Controls.Add(this.SplitSheetTabPage);
            this.ExcelTabCtrol.Controls.Add(this.MergerSheetTabPage);
            this.ExcelTabCtrol.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.ExcelTabCtrol.Location = new System.Drawing.Point(1, 0);
            this.ExcelTabCtrol.Margin = new System.Windows.Forms.Padding(0);
            this.ExcelTabCtrol.Name = "ExcelTabCtrol";
            this.ExcelTabCtrol.Padding = new System.Drawing.Point(3, 3);
            this.ExcelTabCtrol.SelectedIndex = 0;
            this.ExcelTabCtrol.Size = new System.Drawing.Size(685, 163);
            this.ExcelTabCtrol.TabIndex = 1;
            this.ExcelTabCtrol.SelectedIndexChanged += new System.EventHandler(this.ExcelTabCtrol_SelectedIndexChanged);
            // 
            // SplitSheetTabPage
            // 
            this.SplitSheetTabPage.Controls.Add(this.BtnSplitExcel);
            this.SplitSheetTabPage.Controls.Add(this.BtnSelectFile);
            this.SplitSheetTabPage.Controls.Add(this.BtnSelectFolder);
            this.SplitSheetTabPage.Location = new System.Drawing.Point(4, 26);
            this.SplitSheetTabPage.Margin = new System.Windows.Forms.Padding(0);
            this.SplitSheetTabPage.Name = "SplitSheetTabPage";
            this.SplitSheetTabPage.Size = new System.Drawing.Size(677, 133);
            this.SplitSheetTabPage.TabIndex = 0;
            this.SplitSheetTabPage.Text = "Sheet拆分";
            this.SplitSheetTabPage.ToolTipText = "把每个sheet拆分成独立的文件";
            this.SplitSheetTabPage.UseVisualStyleBackColor = true;
            // 
            // BtnSplitExcel
            // 
            this.BtnSplitExcel.Location = new System.Drawing.Point(467, 41);
            this.BtnSplitExcel.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.BtnSplitExcel.Name = "BtnSplitExcel";
            this.BtnSplitExcel.Size = new System.Drawing.Size(131, 45);
            this.BtnSplitExcel.TabIndex = 10;
            this.BtnSplitExcel.Text = "开始拆分";
            this.BtnSplitExcel.UseVisualStyleBackColor = true;
            this.BtnSplitExcel.Click += new System.EventHandler(this.BtnSplitExcel_Click);
            // 
            // BtnSelectFile
            // 
            this.BtnSelectFile.Location = new System.Drawing.Point(58, 41);
            this.BtnSelectFile.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.BtnSelectFile.Name = "BtnSelectFile";
            this.BtnSelectFile.Size = new System.Drawing.Size(131, 45);
            this.BtnSelectFile.TabIndex = 9;
            this.BtnSelectFile.Text = "选择Excel文件";
            this.BtnSelectFile.UseVisualStyleBackColor = true;
            this.BtnSelectFile.Click += new System.EventHandler(this.BtnSelectFile_Click);
            // 
            // BtnSelectFolder
            // 
            this.BtnSelectFolder.Location = new System.Drawing.Point(260, 41);
            this.BtnSelectFolder.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.BtnSelectFolder.Name = "BtnSelectFolder";
            this.BtnSelectFolder.Size = new System.Drawing.Size(131, 45);
            this.BtnSelectFolder.TabIndex = 8;
            this.BtnSelectFolder.Text = "选择文件夹";
            this.BtnSelectFolder.UseVisualStyleBackColor = true;
            this.BtnSelectFolder.Click += new System.EventHandler(this.BtnSelectFolder_Click);
            // 
            // MergerSheetTabPage
            // 
            this.MergerSheetTabPage.Controls.Add(this.CheckBox_OneSheet);
            this.MergerSheetTabPage.Controls.Add(this.MergerSheetBtn);
            this.MergerSheetTabPage.Controls.Add(this.CheckBox_OneFile);
            this.MergerSheetTabPage.Controls.Add(this.MergerSheetSaveFolderBtn);
            this.MergerSheetTabPage.Controls.Add(this.MergerSheetFolderBtn);
            this.MergerSheetTabPage.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.MergerSheetTabPage.Location = new System.Drawing.Point(4, 26);
            this.MergerSheetTabPage.Margin = new System.Windows.Forms.Padding(0);
            this.MergerSheetTabPage.Name = "MergerSheetTabPage";
            this.MergerSheetTabPage.Size = new System.Drawing.Size(677, 133);
            this.MergerSheetTabPage.TabIndex = 1;
            this.MergerSheetTabPage.Text = "合并Sheet";
            this.MergerSheetTabPage.ToolTipText = "合并不同文件的相同顺序sheet为单独文件";
            this.MergerSheetTabPage.UseVisualStyleBackColor = true;
            // 
            // CheckBox_OneSheet
            // 
            this.CheckBox_OneSheet.AutoSize = true;
            this.CheckBox_OneSheet.Location = new System.Drawing.Point(347, 63);
            this.CheckBox_OneSheet.Name = "CheckBox_OneSheet";
            this.CheckBox_OneSheet.Size = new System.Drawing.Size(119, 21);
            this.CheckBox_OneSheet.TabIndex = 14;
            this.CheckBox_OneSheet.Text = "合并成一个Sheet";
            this.ExcelToolTip.SetToolTip(this.CheckBox_OneSheet, "勾选后，所有sheet合并成一个Excel，否则，相同sheet的合并成一个Sheet");
            this.CheckBox_OneSheet.UseVisualStyleBackColor = true;
            // 
            // MergerSheetBtn
            // 
            this.MergerSheetBtn.Location = new System.Drawing.Point(484, 36);
            this.MergerSheetBtn.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.MergerSheetBtn.Name = "MergerSheetBtn";
            this.MergerSheetBtn.Size = new System.Drawing.Size(131, 45);
            this.MergerSheetBtn.TabIndex = 13;
            this.MergerSheetBtn.Text = "开始合并";
            this.MergerSheetBtn.UseVisualStyleBackColor = true;
            this.MergerSheetBtn.Click += new System.EventHandler(this.MergerSheetBtn_Click);
            // 
            // CheckBox_OneFile
            // 
            this.CheckBox_OneFile.AutoSize = true;
            this.CheckBox_OneFile.Location = new System.Drawing.Point(347, 36);
            this.CheckBox_OneFile.Name = "CheckBox_OneFile";
            this.CheckBox_OneFile.Size = new System.Drawing.Size(116, 21);
            this.CheckBox_OneFile.TabIndex = 12;
            this.CheckBox_OneFile.Text = "合并成一个Excel";
            this.ExcelToolTip.SetToolTip(this.CheckBox_OneFile, "勾选后，所有sheet合并成一个Excel，否则，相同sheet的合并成一个Excel");
            this.CheckBox_OneFile.UseVisualStyleBackColor = true;
            // 
            // MergerSheetSaveFolderBtn
            // 
            this.MergerSheetSaveFolderBtn.Location = new System.Drawing.Point(189, 36);
            this.MergerSheetSaveFolderBtn.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.MergerSheetSaveFolderBtn.Name = "MergerSheetSaveFolderBtn";
            this.MergerSheetSaveFolderBtn.Size = new System.Drawing.Size(131, 45);
            this.MergerSheetSaveFolderBtn.TabIndex = 10;
            this.MergerSheetSaveFolderBtn.Text = "选择保存文件夹";
            this.MergerSheetSaveFolderBtn.UseVisualStyleBackColor = true;
            this.MergerSheetSaveFolderBtn.Click += new System.EventHandler(this.MergerSheetSaveFolderBtn_Click);
            // 
            // MergerSheetFolderBtn
            // 
            this.MergerSheetFolderBtn.Location = new System.Drawing.Point(37, 36);
            this.MergerSheetFolderBtn.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.MergerSheetFolderBtn.Name = "MergerSheetFolderBtn";
            this.MergerSheetFolderBtn.Size = new System.Drawing.Size(131, 45);
            this.MergerSheetFolderBtn.TabIndex = 9;
            this.MergerSheetFolderBtn.Text = "选择源文件夹";
            this.MergerSheetFolderBtn.UseVisualStyleBackColor = true;
            this.MergerSheetFolderBtn.Click += new System.EventHandler(this.MergerSheetFolderBtn_Click);
            // 
            // ExcelToolsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(684, 189);
            this.Controls.Add(this.ExcelTabCtrol);
            this.Controls.Add(this.ExcelToolsStatus);
            this.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ExcelToolsForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Excel Tools";
            this.ExcelToolsStatus.ResumeLayout(false);
            this.ExcelToolsStatus.PerformLayout();
            this.ExcelTabCtrol.ResumeLayout(false);
            this.SplitSheetTabPage.ResumeLayout(false);
            this.MergerSheetTabPage.ResumeLayout(false);
            this.MergerSheetTabPage.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.StatusStrip ExcelToolsStatus;
        private System.Windows.Forms.ToolStripStatusLabel copyrightLab;
        private System.Windows.Forms.ToolStripProgressBar ExcelStatusProgressBar;
        private System.Windows.Forms.TabControl ExcelTabCtrol;
        private System.Windows.Forms.TabPage SplitSheetTabPage;
        private System.Windows.Forms.Button BtnSplitExcel;
        private System.Windows.Forms.Button BtnSelectFile;
        private System.Windows.Forms.Button BtnSelectFolder;
        private System.Windows.Forms.TabPage MergerSheetTabPage;
        private System.Windows.Forms.ToolTip ExcelToolTip;
        private System.Windows.Forms.Button MergerSheetBtn;
        private System.Windows.Forms.CheckBox CheckBox_OneFile;
        private System.Windows.Forms.Button MergerSheetSaveFolderBtn;
        private System.Windows.Forms.Button MergerSheetFolderBtn;
        private System.Windows.Forms.CheckBox CheckBox_OneSheet;
    }
}

