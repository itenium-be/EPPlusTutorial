using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace EPPlusTutorial.Util
{
    public static class ExcelHelpers
    {
        public static void SetHeaders(this ExcelRangeBase cell, params string[] headers)
        {
            foreach (string text in headers)
            {
                cell.Value = text;
                cell.Style.Font.Bold = true;
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(Color.DarkBlue);
                cell.Style.Font.Color.SetColor(Color.White);
                cell.AutoFilter = true;

                cell = cell.Offset(0, 1);
            }
        }
    }
}