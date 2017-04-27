using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using NUnit.Framework;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using EPPlusTutorial.Util;

namespace EPPlusTutorial
{
    [TestFixture]
    public class QuickTutorial
    {
        [Test]
        public void BasicUsage()
        {
            using (var package = new ExcelPackage())
            {
                ExcelWorksheet sheet1 = package.Workbook.Worksheets.Add("MySheet");
                ExcelRange firstCell = sheet1.Cells[1, 1]; // or use "A1"
                firstCell.Value = "will it work...";
                sheet1.Cells.AutoFitColumns();
                package.SaveAs(new FileInfo(BinDir.GetPath()));
            }
        }

        [Test]
        public void LoadingAndSaving()
        {
            // Open an existing Excel
            // Or if the file does not exist, create a new one
            using (var package = new ExcelPackage(new FileInfo(BinDir.GetPath()), "optionalPassword"))
            using (var basicUsageExcel = File.Open(BinDir.GetPath(nameof(BasicUsage)), FileMode.Open))
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");
                sheet.Cells["D1"].Value = "Everything in the package will be overwritten";
                sheet.Cells["D2"].Value = "by the package.Load() below!!!";

                // Loads the worksheets from BasicUsage
                // (MySheet with A1 = will it work...)
                package.Load(basicUsageExcel);

                // See 3-Import for more loading techniques

                package.Save("optionalPassword");
                //package.SaveAs(FileInfo / Stream)
                //Byte[] p = package.GetAsByteArray();
            }
        }

        [Test]
        public void SelectingCells()
        {
            using (var package = new ExcelPackage())
            {
                var sheet1 = package.Workbook.Worksheets.Add("MySheet");

                // One cell
                ExcelRange cellA2 = sheet1.Cells["A2"];
                var alsoCellA2 = sheet1.Cells[2, 1];
                Assert.That(cellA2.Address, Is.EqualTo("A2"));
                Assert.That(cellA2.Address, Is.EqualTo(alsoCellA2.Address));

                // Column from a cell
                // ExcelRange.Start is the top and left most cell
                Assert.That(cellA2.Start.Column, Is.EqualTo(1));
                // To really get the column: sheet1.Column(1)

                // A range
                ExcelRange ranger = sheet1.Cells["A2:C5"];
                var sameRanger = sheet1.Cells[2, 1, 5, 3];
                Assert.That(ranger.Address, Is.EqualTo(sameRanger.Address));

                // Dimensions used
                Assert.That(sheet1.Dimension, Is.Null);

                ranger.Value = "pushing";
                var usedDimensions = sheet1.Dimension;
                Assert.That(usedDimensions.Address, Is.EqualTo(ranger.Address));

                // Offset: down 5 rows, right 10 columns
                var movedRanger = ranger.Offset(5, 10);
                Assert.That(movedRanger.Address, Is.EqualTo("K7:M10"));
                movedRanger.Value = "Moved";

                package.SaveAs(new FileInfo(BinDir.GetPath()));
            }
        }

        [Test]
        public void WritingValues()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
            using (var package = new ExcelPackage())
            {
                var sheet1 = package.Workbook.Worksheets.Add("MySheet");

                // Numbers
                sheet1.SetValue("A1", "Numbers");
                Assert.That(sheet1.GetValue<string>(1, 1), Is.EqualTo("Numbers"));
                sheet1.Cells["B1"].Value = 15.32;
                sheet1.Cells["B1"].Style.Numberformat.Format = "#,##0.00";
                Assert.That(sheet1.Cells["B1"].Text, Is.EqualTo("15.32"));

                // Money
                sheet1.Cells["A2"].Value = "Moneyz";
                sheet1.Cells["B2"].Value = 15000.23D;
                sheet1.Cells["C2"].Value = -2000.50D;
                sheet1.Cells["B2:C2"].Style.Numberformat.Format = "#,##0.00 [$€-813];[RED]-#,##0.00 [$€-813]";

                // DateTime
                sheet1.Cells["A3"].Value = "Timey Wimey";
                sheet1.Cells["B3"].Style.Numberformat.Format = "yyyy-mm-dd";
                sheet1.Cells["B3"].Formula = $"=DATE({DateTime.Now:yyyy,MM,dd})";
                sheet1.Cells["C3"].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.FullDateTimePattern;
                sheet1.Cells["C3"].Value = DateTime.Now;
                sheet1.Cells["D3"].Style.Numberformat.Format = "dd/MM/yyyy HH:mm";
                sheet1.Cells["D3"].Value = DateTime.Now;
                 

                // An external hyperlink
                sheet1.Cells["C25"].Formula = "HYPERLINK(\"mailto:support@pongit.be\",\"Contact support\")";
                sheet1.Cells["C25"].Style.Font.Color.SetColor(Color.Blue);
                sheet1.Cells["C25"].Style.Font.UnderLine = true;
                
                // An internal hyperlink
                package.Workbook.Worksheets.Add("Data");
                sheet1.Cells["C26"].Hyperlink = new ExcelHyperLink("Data!A1", "Goto data sheet");


                sheet1.Cells.AutoFitColumns();
                package.SaveAs(new FileInfo(BinDir.GetPath()));
            }
        }

        [Test]
        public void FormattingCells()
        {
            using (var package = new ExcelPackage())
            {
                var sheet1 = package.Workbook.Worksheets.Add("Styling");

                // Cells with style
                ExcelFont font = sheet1.Cells["A1"].Style.Font;
                sheet1.Cells["A1"].Value = "Bold and proud";
                sheet1.Cells["A1"].Style.Font.Name = "Arial";
                font.Bold = true;
                font.Color.SetColor(Color.Green);
                // ExcelFont also has: Size, Italic, Underline, Strike, ...

                sheet1.Cells["A3"].Style.Font.SetFromFont(new Font(new FontFamily("Arial"), 15, FontStyle.Strikeout));
                sheet1.Cells["A3"].Value = "SetFromFont(Font)";

                // Borders need to be made
                sheet1.Cells["A1:A2"].Style.Border.BorderAround(ExcelBorderStyle.Dotted);
                sheet1.Cells[5, 5, 9, 8].Style.Border.BorderAround(ExcelBorderStyle.Dotted);

                // Merge cells
                sheet1.Cells[5, 5, 9, 8].Merge = true;

                // More style
                sheet1.Cells["D14"].Style.ShrinkToFit = true;
                sheet1.Cells["D14"].Style.Font.Size = 24;
                sheet1.Cells["D14"].Value = "Shrinking for fit";

                sheet1.Cells["D15"].Style.WrapText = true;
                sheet1.Cells["D15"].Value = "A wrap, yummy!";
                sheet1.Cells["D16"].Value = "No wrap, ouch!";

                // Setting a background color requires setting the PatternType first
                sheet1.Cells["F6:G8"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet1.Cells["F6:G8"].Style.Fill.BackgroundColor.SetColor(Color.Red);

                // Horizontal Alignment needs a little workaround
                // http://stackoverflow.com/questions/34660560/epplus-isnt-honoring-excelhorizontalalignment-center-or-right
                var centerStyle = package.Workbook.Styles.CreateNamedStyle("Center");
                centerStyle.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet1.Cells["B5"].StyleName = "Center";
                sheet1.Cells["B5"].Value = "I'm centered";

                // MIGHT NOT WORK:
                sheet1.Cells["B6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet1.Cells["B6"].Value = "I'm not centered? :(";

                package.SaveAs(new FileInfo(BinDir.GetPath()));
            }
        }

        [Test]
        public void FormattingSheetsAndColumns()
        {
            using (var package = new ExcelPackage())
            {
                var sheet1 = package.Workbook.Worksheets.Add("Styling");
                sheet1.TabColor = Color.Red;
                //sheet1.DeleteColumn();
                //sheet1.InsertColumn();

                // Default selected cells when opening the xslx
                sheet1.Select("B6");

                var colE = sheet1.Column(5);
                //ExcelStyle colStyle = colE.Style; // See FormattingCells
                colE.AutoFit(); // or colE.Width

                // Who likes A's
                sheet1.Column(1).Hidden = true;

                package.SaveAs(new FileInfo(BinDir.GetPath()));
            }
        }

        [Test]
        public void ConvertingIndexesAndAddresses()
        {
            Assert.That(ExcelCellBase.GetAddress(1, 1), Is.EqualTo("A1"));
            Assert.That(ExcelCellBase.IsValidCellAddress("A5"), Is.True);

            Assert.That(ExcelCellBase.GetFullAddress("MySheet", "A1:A3"), Is.EqualTo("'MySheet'!A1:A3"));

            Assert.That(ExcelCellBase.TranslateToR1C1("AB23", 0, 0), Is.EqualTo("R[23]C[28]"));
            Assert.That(ExcelCellBase.TranslateFromR1C1("R23C28", 0, 0), Is.EqualTo("$AB$23"));
        }
    }
}
