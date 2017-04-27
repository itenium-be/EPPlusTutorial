using System.IO;
using NUnit.Framework;
using OfficeOpenXml;

namespace EPPlusTutorial
{
    //[TestFixture]
    public class Formulas
    {
        [Test]
        public void WorkingWithFormulas()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Formula");
                var sheetData = package.Workbook.Worksheets.Add("Data");

                //Assert.That(cellA2.FullAddress, Is.EqualTo("A2"));

                //sheet.Dimension.Table
                //ConditionalFormatting

                //worksheet.Cells["A1"].Formula="CONCATENATE(\"string1_\",\"test\")";
                //.Formula = "=B3*10"

                // Absolute reference: $B$5

                package.SaveAs(new FileInfo(BinDir.GetPath()));
            }
        }
    }
}