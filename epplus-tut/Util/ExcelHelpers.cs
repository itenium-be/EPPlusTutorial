using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace EPPlusTutorial.Util
{
    public static class ExcelHelpers
    {
        public static ExcelRangeBase SetHeaders(this ExcelRangeBase cells, params string[] headers)
        {
            foreach (string text in headers)
            {
                cells.Value = text;
                cells.Style.Font.Bold = true;
                cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cells.Style.Fill.BackgroundColor.SetColor(Color.DarkBlue);
                cells.Style.Font.Color.SetColor(Color.White);

                cells = cells.Offset(0, 1);
            }
            return cells;
        }
    }
}