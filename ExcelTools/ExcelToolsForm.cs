using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using OfficeOpenXml;

namespace ExcelTools
{
    public partial class ExcelToolsForm : Form
    {
        private string _targetFilePath;
        private string _targetFolderPath;

        public ExcelToolsForm()
        {
            InitializeComponent();
            BtnSplitExcel.Enabled = false;
            MergerSheetBtn.Enabled = false;
            ShowToolTip();
        }

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
            }
            CheckBtnSplit();
            BtnSelectFile.Enabled = true;
        }

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
            }
            CheckBtnSplit();
            BtnSelectFolder.Enabled = true;
        }

        private void CheckBtnSplit()
        {
            if (!string.IsNullOrEmpty(_targetFilePath) && !string.IsNullOrEmpty(_targetFolderPath))
            {
                BtnSplitExcel.Enabled = true;
            }
        }

        private void BtnSplitExcel_Click(object sender, EventArgs e)
        {
            BtnSplitExcel.Enabled = false;
            BtnSelectFile.Enabled = false;
            BtnSelectFolder.Enabled = false;
            ExcelStatusProgressBar.Value = 0;
            var fileInfo = new FileInfo(_targetFilePath);
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
                    ExcelStatusProgressBar.Value = Convert.ToInt32(Math.Floor(i * 100.0 / workSheetCount));
                }
            }
            ExcelStatusProgressBar.Value = 100;
            MessageBox.Show(@"Excel拆分完成");
            BtnSplitExcel.Enabled = true;
            BtnSelectFile.Enabled = true;
            BtnSelectFolder.Enabled = true;
        }


        private void ExcelTabCtrol_SelectedIndexChanged(object sender, EventArgs e)
        {
            ShowToolTip();
        }

        private void ShowToolTip()
        {
            var toolTipText = ExcelTabCtrol.TabPages[ExcelTabCtrol.SelectedIndex].ToolTipText;
            ExcelToolTip.SetToolTip(ExcelTabCtrol, toolTipText);
        }


        // 合并Sheet

        private string _targetMergerSourceFolderPath;
        private string _targetMergerSaveFolderPath;

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
            }
            MergerSheetFolderBtn.Enabled = true;
            CheckBtnMerger();
        }

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
            }
            MergerSheetSaveFolderBtn.Enabled = true;
            CheckBtnMerger();
        }

        private void CheckBtnMerger()
        {
            if (!string.IsNullOrEmpty(_targetMergerSourceFolderPath) && !string.IsNullOrEmpty(_targetMergerSaveFolderPath))
            {
                MergerSheetBtn.Enabled = true;
            }
        }

        private void MergerSheetBtn_Click(object sender, EventArgs e)
        {
            MergerSheetFolderBtn.Enabled = false;
            MergerSheetSaveFolderBtn.Enabled = false;
            MergerSheetBtn.Enabled = false;
            ExcelStatusProgressBar.Value = 0;
            const string excelExt = ".xlsx";
            var filePaths = Directory.GetFiles(_targetMergerSourceFolderPath, "*.xlsx");
            if (filePaths.Length == 0)
            {
                MessageBox.Show(@"该文件夹没有Excel文件");
                MergerSheetFolderBtn.Enabled = true;
                MergerSheetSaveFolderBtn.Enabled = true;
                MergerSheetBtn.Enabled = false;
                return;
            }
            var maxSheetCount = 0;
            var fileInfos = new List<FileInfo>();
            foreach (var filePath in filePaths)
            {
                var fileInfo = new FileInfo(filePath);
                fileInfos.Add(fileInfo);
                using (var excelPackage = new ExcelPackage(fileInfo))
                {
                    if (maxSheetCount < excelPackage.Workbook.Worksheets.Count)
                    {
                        maxSheetCount = excelPackage.Workbook.Worksheets.Count;
                    }
                }
            }

            if (CheckBox_OneFile.Checked)
            {
                using (var newexcelPackage = new ExcelPackage())
                {
                    for (var i = 1; i <= maxSheetCount; i++)
                    {
                        foreach (var fileInfo in fileInfos)
                        {
                            using (var excelPackage = new ExcelPackage(fileInfo))
                            {
                                if (excelPackage.Workbook.Worksheets.Count < i)
                                {
                                    continue;
                                }
                                var worksheet = excelPackage.Workbook.Worksheets[i];
                                if (CheckBox_OneSheet.Checked)
                                {
                                    if (newexcelPackage.Workbook.Worksheets.Count > 0)
                                    {
                                        newexcelPackage.Workbook.Worksheets.First().Combine(worksheet);
                                    }
                                    else
                                    {
                                        newexcelPackage.Workbook.Worksheets.Add("Sheet", worksheet);
                                    }
                                }
                                else
                                {
                                    newexcelPackage.Workbook.Worksheets.Add(fileInfo.Name.Replace(fileInfo.Extension, "") + " " + worksheet.Name, worksheet);
                                }
                            }
                        }
                        ExcelStatusProgressBar.Value = Convert.ToInt32(Math.Floor(i * 100.0 / maxSheetCount));
                    }
                    var newExcelName = $"{DateTime.Now:yyyy-MM-dd} 合并所有Sheet{excelExt}";
                    var filePath = Path.Combine(_targetMergerSaveFolderPath, newExcelName);
                    while (File.Exists(filePath))
                    {
                        filePath = Path.Combine(_targetMergerSaveFolderPath, $"{DateTime.Now.Millisecond}-{newExcelName}");
                    }
                    newexcelPackage.SaveAs(new FileInfo(filePath));
                }
            }
            else
            {
                if (CheckBox_OneSheet.Checked)
                {
                    for (var i = 1; i <= maxSheetCount; i++)
                    {
                        using (var newexcelPackage = new ExcelPackage())
                        {
                            var newExcelName = string.Empty;
                            foreach (var fileInfo in fileInfos)
                            {
                                using (var excelPackage = new ExcelPackage(fileInfo))
                                {
                                    if (excelPackage.Workbook.Worksheets.Count < i)
                                    {
                                        continue;
                                    }
                                    var worksheet = excelPackage.Workbook.Worksheets[i];
                                    if (CheckBox_OneSheet.Checked)
                                    {
                                        if (newexcelPackage.Workbook.Worksheets.Count > 0)
                                        {
                                            newexcelPackage.Workbook.Worksheets.First().Combine(worksheet);
                                        }
                                        else
                                        {
                                            newexcelPackage.Workbook.Worksheets.Add("Sheet", worksheet);
                                            newExcelName = worksheet.Name;
                                        }
                                    }
                                    else
                                    {
                                        newexcelPackage.Workbook.Worksheets.Add(fileInfo.Name.Replace(fileInfo.Extension, ""), worksheet);
                                        if (string.IsNullOrEmpty(newExcelName))
                                        {
                                            newExcelName = worksheet.Name;
                                        }
                                    }
                                }
                            }
                            newExcelName += excelExt;
                            var filePath = Path.Combine(_targetMergerSaveFolderPath, newExcelName);
                            while (File.Exists(filePath))
                            {
                                filePath = Path.Combine(_targetMergerSaveFolderPath, i + "-" + newExcelName);
                            }
                            newexcelPackage.SaveAs(new FileInfo(filePath));
                        }
                        ExcelStatusProgressBar.Value = Convert.ToInt32(Math.Floor(i * 100.0 / maxSheetCount));
                    }
                }
            }

            ExcelStatusProgressBar.Value = 100;
            MessageBox.Show(@"Excel合并完成");
            MergerSheetFolderBtn.Enabled = true;
            MergerSheetSaveFolderBtn.Enabled = true;
            MergerSheetBtn.Enabled = true;
        }
    }
}
