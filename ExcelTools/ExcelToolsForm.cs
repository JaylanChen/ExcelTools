using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;

namespace ExcelTools
{
    public partial class ExcelToolsForm : Form
    {
        public ExcelToolsForm()
        {
            InitializeComponent();
            BtnSplitExcel.Enabled = false;
            MergerSheetBtn.Enabled = false;
            ShowToolTip();
        }

        /// <summary>
        /// 切换Tab时展示不同的提示
        /// </summary>
        private void ExcelTabCtrol_SelectedIndexChanged(object sender, EventArgs e)
        {
            ShowToolTip();
        }

        private void ShowToolTip()
        {
            var toolTipText = ExcelTabCtrol.TabPages[ExcelTabCtrol.SelectedIndex].ToolTipText;
            ExcelToolTip.SetToolTip(ExcelTabCtrol, toolTipText);
        }


        #region 合并Sheet

        // 源文件夹路径
        private string _targetMergerSourceFolderPath;
        // 保存文件夹路径
        private string _targetMergerSaveFolderPath;

        /// <summary>
        /// 选择源文件夹路径
        /// </summary>
        private void MergerSheetFolderBtn_Click(object sender, EventArgs e)
        {
            MergerSheetFolderBtn.Enabled = false;
            var folderDialog = new FolderBrowserDialog
            {
                Description = @"请选择需要合并Sheet的Excel文件夹",
                SelectedPath = @"c:\"
            };
            if (folderDialog.ShowDialog() == DialogResult.OK)
            {
                _targetMergerSourceFolderPath = folderDialog.SelectedPath;
                lbSourceFolder.Text = _targetMergerSourceFolderPath;
                ExcelToolTip.SetToolTip(lbSourceFolder, _targetMergerSourceFolderPath);
            }
            MergerSheetFolderBtn.Enabled = true;
            CheckBtnMerger();
        }

        /// <summary>
        /// 选择保存文件夹路径
        /// </summary>
        private void MergerSheetSaveFolderBtn_Click(object sender, EventArgs e)
        {
            MergerSheetSaveFolderBtn.Enabled = false;
            var folderDialog = new FolderBrowserDialog
            {
                Description = @"请选择合并Sheet后的Excel保存文件夹",
                SelectedPath = @"c:\"
            };
            if (folderDialog.ShowDialog() == DialogResult.OK)
            {
                _targetMergerSaveFolderPath = folderDialog.SelectedPath;
                lbTargetFolder.Text = _targetMergerSaveFolderPath;
                ExcelToolTip.SetToolTip(lbTargetFolder, _targetMergerSaveFolderPath);
            }
            MergerSheetSaveFolderBtn.Enabled = true;
            CheckBtnMerger();
        }

        /// <summary>
        /// 检测路径是否已选择
        /// </summary>
        private void CheckBtnMerger()
        {
            if (!string.IsNullOrEmpty(_targetMergerSourceFolderPath) && !string.IsNullOrEmpty(_targetMergerSaveFolderPath))
            {
                MergerSheetBtn.Enabled = true;
            }
        }

        /// <summary>
        /// 是否合并为一个Sheet
        /// </summary>
        private void CheckBox_OneSheet_CheckedChanged(object sender, EventArgs e)
        {
            numUdSkipRows.Visible = CheckBox_OneSheet.Checked;
            lbSkipRows.Visible = CheckBox_OneSheet.Checked;
            numUdSkipRows.Value = CheckBox_OneSheet.Checked ? 1 : 0;
        }

        /// <summary>
        /// Excel 合并
        /// </summary>
        private void MergerSheetBtn_Click(object sender, EventArgs e)
        {
            MergerSheetFolderBtn.Enabled = false;
            MergerSheetSaveFolderBtn.Enabled = false;
            MergerSheetBtn.Enabled = false;
            CheckBox_OneFile.Enabled = false;
            CheckBox_OneSheet.Enabled = false;
            numUdSkipRows.Enabled = false;
            ExcelStatusProgressBar.Value = 0;
            var filePaths = Directory.GetFiles(_targetMergerSourceFolderPath, "*.xlsx");
            if (filePaths.Length == 0)
            {
                MessageBox.Show(@"该文件夹没有Excel文件");
                MergerSheetFolderBtn.Enabled = true;
                MergerSheetSaveFolderBtn.Enabled = true;
                MergerSheetBtn.Enabled = false;
                return;
            }
            try
            {
                var task = Task.Run(() => MergerExcels(filePaths));
            }
            catch (Exception exception)
            {
                MessageBox.Show($@"执行错误: {exception.Message}");
                MergerSheetFolderBtn.Enabled = true;
                MergerSheetSaveFolderBtn.Enabled = true;
                MergerSheetBtn.Enabled = true;
                CheckBox_OneFile.Enabled = true;
                CheckBox_OneSheet.Enabled = true;
                numUdSkipRows.Enabled = true;
                return;
            }
        }

        /// <summary>
        /// 合并Excel
        /// </summary>
        /// <param name="filePaths">Excel文件路径列表</param>
        private void MergerExcels(IReadOnlyList<string> filePaths)
        {
            const string excelExt = ".xlsx";
            var excelPackageDicts = new Dictionary<string, ExcelPackage>();
            for (var i = 0; i < filePaths.Count; i++)
            {
                var fileInfo = new FileInfo(filePaths[i]);
                using (var excelPackage = new ExcelPackage(fileInfo))
                {
                    for (var sheetIndex = 1; sheetIndex <= excelPackage.Workbook.Worksheets.Count; sheetIndex++)
                    {
                        var workSheet = excelPackage.Workbook.Worksheets[sheetIndex];
                        var excelName = workSheet.Name;
                        var excelIndex = sheetIndex - 1;
                        if (CheckBox_OneFile.Checked)
                        {
                            excelName = $@"{DateTime.Now:yyyy-MM-dd HH-mm} 合并所有文件";
                            excelIndex = 0;
                        }
                        var currentExcelPackage = excelPackageDicts.Values.ElementAtOrDefault(excelIndex);
                        if (currentExcelPackage == null)
                        {
                            currentExcelPackage = new ExcelPackage();
                            excelPackageDicts.Add(excelName, currentExcelPackage);
                        }

                        if (CheckBox_OneSheet.Checked)
                        {
                            if (currentExcelPackage.Workbook.Worksheets.Count > 0)
                            {
                                currentExcelPackage.Workbook.Worksheets.First().Combine(workSheet, (int)numUdSkipRows.Value);
                            }
                            else
                            {
                                currentExcelPackage.Workbook.Worksheets.Add("Sheet", workSheet);
                            }
                        }
                        else
                        {
                            currentExcelPackage.Workbook.Worksheets.Add(fileInfo.Name.Replace(fileInfo.Extension, "") + " " + workSheet.Name, workSheet);
                        }
                    }
                }

                var progressValue = Convert.ToInt32(Math.Floor(i * 90.0 / filePaths.Count));
                this.Invoke((MethodInvoker)(() => ExcelStatusProgressBar.Value = progressValue));
            }
            var fileCount = 0;
            foreach (var excelPackageDict in excelPackageDicts)
            {
                var filePath = Path.Combine(_targetMergerSaveFolderPath, excelPackageDict.Key + excelExt);
                while (File.Exists(filePath))
                {
                    filePath = Path.Combine(_targetMergerSaveFolderPath, $"{excelPackageDict.Key}-{DateTime.Now.Millisecond}{excelExt}");
                }
                excelPackageDict.Value.SaveAs(new FileInfo(filePath));
                excelPackageDict.Value.Dispose();

                fileCount++;
                var progressValue = 90 + Convert.ToInt32(Math.Floor(fileCount * 10.0 / excelPackageDicts.Count));
                this.Invoke((MethodInvoker)(() => ExcelStatusProgressBar.Value = progressValue));
            }

            this.Invoke((MethodInvoker)(() =>
            {
                ExcelStatusProgressBar.Value = 100;
                MessageBox.Show(@"Excel合并完成");
                MergerSheetFolderBtn.Enabled = true;
                MergerSheetSaveFolderBtn.Enabled = true;
                MergerSheetBtn.Enabled = true;
                CheckBox_OneFile.Enabled = true;
                CheckBox_OneSheet.Enabled = true;
                numUdSkipRows.Enabled = true;
            }));
        }

        #endregion


        #region Excel拆分

        /// <summary>
        /// 源文件路径
        /// </summary>
        private string _targetFilePath;
        /// <summary>
        /// 保存文件夹路径
        /// </summary>
        private string _targetFolderPath;


        /// <summary>
        /// 选择拆分源文件路径
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnSelectFile_Click(object sender, EventArgs e)
        {
            BtnSelectFile.Enabled = false;
            var fileDialog = new OpenFileDialog
            {
                Title = @"请选择要拆分Sheet的Excel文件",
                InitialDirectory = @"c:\",
                Filter = @"Excel(*.xlsx)|*.xlsx",
                AddExtension = true,
                DefaultExt = "(.xlsx)"
            };
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                _targetFilePath = fileDialog.FileName;
                lbSourceFilePath.Text = _targetFilePath;
                ExcelToolTip.SetToolTip(lbSourceFilePath, _targetFilePath);
            }
            CheckBtnSplit();
            BtnSelectFile.Enabled = true;
        }

        /// <summary>
        /// 选择保存文件夹路径
        /// </summary>
        private void BtnSelectFolder_Click(object sender, EventArgs e)
        {
            BtnSelectFolder.Enabled = false;
            var folderDialog = new FolderBrowserDialog
            {
                Description = @"请选择拆分后Excel存储文件夹",
                SelectedPath = @"c:\"
            };
            if (folderDialog.ShowDialog() == DialogResult.OK)
            {
                _targetFolderPath = folderDialog.SelectedPath;
                lbSaveFolder.Text = _targetFolderPath;
                ExcelToolTip.SetToolTip(lbSaveFolder, _targetFolderPath);
            }
            CheckBtnSplit();
            BtnSelectFolder.Enabled = true;
        }

        /// <summary>
        /// 检测拆分Excel 源文件路径 和 保存文件夹路径是否已选择
        /// </summary>
        private void CheckBtnSplit()
        {
            if (!string.IsNullOrEmpty(_targetFilePath) && !string.IsNullOrEmpty(_targetFolderPath))
            {
                BtnSplitExcel.Enabled = true;
            }
        }

        /// <summary>
        /// Excel 拆分
        /// </summary>
        private void BtnSplitExcel_Click(object sender, EventArgs e)
        {
            BtnSplitExcel.Enabled = false;
            BtnSelectFile.Enabled = false;
            BtnSelectFolder.Enabled = false;
            ExcelStatusProgressBar.Value = 0;
            try
            {
                Task.Run(() => SplitExcel(_targetFilePath));
            }
            catch (Exception exception)
            {
                MessageBox.Show($@"执行错误: {exception.Message}");
                BtnSplitExcel.Enabled = true;
                BtnSelectFile.Enabled = true;
                BtnSelectFolder.Enabled = true;
                return;
            }
        }

        /// <summary>
        /// 拆分Excel
        /// </summary>
        /// <param name="sourceFilePath"></param>
        private void SplitExcel(string sourceFilePath)
        {
            var fileInfo = new FileInfo(sourceFilePath);
            var excelExt = Path.GetExtension(_targetFilePath);
            using (var excelPackage = new ExcelPackage(fileInfo))
            {
                var workSheetCount = excelPackage.Workbook.Worksheets.Count;
                for (var i = 1; i < workSheetCount; i++)
                {
                    var worksheet = excelPackage.Workbook.Worksheets[i];
                    var newExcelName = worksheet.Name + excelExt;

                    using (var newexcelPackage = new ExcelPackage())
                    {
                        var filePath = Path.Combine(_targetFolderPath, newExcelName);
                        while (File.Exists(filePath))
                        {
                            filePath = Path.Combine(_targetFolderPath, i + "-" + newExcelName);
                        }
                        newexcelPackage.Workbook.Worksheets.Add(newExcelName, worksheet);
                        newexcelPackage.SaveAs(new FileInfo(filePath));
                    }
                    var progressValue = Convert.ToInt32(Math.Floor(i * 100.0 / workSheetCount));
                    this.Invoke((MethodInvoker)(() => ExcelStatusProgressBar.Value = progressValue));
                }
            }
            this.Invoke((MethodInvoker)(() =>
            {
                ExcelStatusProgressBar.Value = 100;
                MessageBox.Show(@"Excel拆分完成");
                BtnSplitExcel.Enabled = true;
                BtnSelectFile.Enabled = true;
                BtnSelectFolder.Enabled = true;
            }));
        }

        #endregion
    }
}
