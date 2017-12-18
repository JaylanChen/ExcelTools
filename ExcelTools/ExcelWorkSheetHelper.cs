using OfficeOpenXml;

namespace ExcelTools
{
    public static class ExcelWorkSheetHelper
    {
        public static ExcelWorksheet Combine(this ExcelWorksheet firstSheet, ExcelWorksheet secondSheet, int skipSecondHeaderRows = 1)
        {
            var beginRowIndex = skipSecondHeaderRows + 1;
            var appendRowIndex = firstSheet.Dimension.End.Row;
            for (var i = beginRowIndex; i <= secondSheet.Dimension.End.Row; i++)
            {
                for (var j = secondSheet.Dimension.Start.Column; j <= secondSheet.Dimension.End.Column; j++)
                {
                    secondSheet.Cells[i, j].Copy(firstSheet.Cells[appendRowIndex + i, j], ExcelRangeCopyOptionFlags.ExcludeFormulas);
                }
            }
            return firstSheet;
        }
    }
}
