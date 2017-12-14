using OfficeOpenXml;

namespace ExcelTools
{
    public static class ExcelWorkSheetHelper
    {
        public static ExcelWorksheet Combine(this ExcelWorksheet firstSheet, ExcelWorksheet secondSheet, bool skipSecondHeader = true)
        {
            var beginRowIndex = 1;
            if (skipSecondHeader)
            {
                beginRowIndex = 2;
            }
            var appendRowIndex = firstSheet.Dimension.End.Row;
            for (var i = beginRowIndex; i < secondSheet.Dimension.End.Row; i++)
            {
                for (var j = secondSheet.Dimension.Start.Column; j < secondSheet.Dimension.End.Column; j++)
                {
                    firstSheet.Cells[appendRowIndex + i, j].Value = secondSheet.Cells[i, j].Value;
                }
            }
            return firstSheet;
        }
    }
}
