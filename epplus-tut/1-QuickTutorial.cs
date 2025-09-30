using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using NUnit.Framework;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using EPPlusTutorial.Util;
using OfficeOpenXml.ConditionalFormatting;

namespace EPPlusTutorial
{
    [TestFixture]
    public class QuickTutorial
    {
        [Test]
        public void BasicUsage()
        {
            // Put key in web.config or appconfig:
            // EPPlus:ExcelPackage.LicenseContext
            // See: https://github.com/EPPlusSoftware/EPPlus#licensecontext-parameter-must-be-set
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage())
            {
                ExcelWorksheet sheet = package.Workbook.Worksheets.Add("MySheet");

                // Setting & getting values
                ExcelRange firstCell = sheet.Cells[1, 1];
                firstCell.Value = "will it work?";
                sheet.Cells["A2"].Formula = "CONCATENATE(A1,\" ... Of course it will!\")";
                Assert.That(firstCell.Text, Is.EqualTo("will it work?"));

                // Numbers
                var moneyCell = sheet.Cells["A3"];
                moneyCell.Style.Numberformat.Format = "$#,##0.00";
                moneyCell.Value = 1500.25M;

                // Easily write any Enumerable to a sheet
                // In this case: All Excel functions implemented by EPPlus
                var funcs = package.Workbook.FormulaParserManager.GetImplementedFunctions()
                    .Select(x => new { FunctionName = x.Key, TypeName = x.Value.GetType().FullName });
                sheet.Cells["A4"].LoadFromCollection(funcs, true);

                // Styling cells
                var someCells = sheet.Cells["A1,A4:B4"];
                someCells.Style.Font.Bold = true;
                someCells.Style.Font.Color.SetColor(Color.Ivory);
                someCells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                someCells.Style.Fill.BackgroundColor.SetColor(Color.Navy);

                sheet.Cells.AutoFitColumns();
                //package.SaveAs(new FileInfo(@"basicUsage.xslx"));
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

                // Load the worksheets from BasicUsage
                // (MySheet with A1 = will it work?)
                package.Load(basicUsageExcel);

                // See 3-Import for more loading techniques

                package.Compression = CompressionLevel.BestSpeed;
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
                var sheet = package.Workbook.Worksheets.Add("MySheet");

                // One cell
                ExcelRange cellA2 = sheet.Cells["A2"];
                var alsoCellA2 = sheet.Cells[2, 1];
                Assert.That(cellA2.Address, Is.EqualTo("A2"));
                Assert.That(cellA2.Address, Is.EqualTo(alsoCellA2.Address));

                // Get the column from a cell
                // ExcelRange.Start is the top and left most cell
                Assert.That(cellA2.Start.Column, Is.EqualTo(1));
                // To really get the column: sheet.Column(1)

                // A range
                ExcelRange ranger = sheet.Cells["A2:C5"];
                var sameRanger = sheet.Cells[2, 1, 5, 3];
                Assert.That(ranger.Address, Is.EqualTo(sameRanger.Address));

                //sheet.Cells["A1,A4"] // Just A1 and A4
                //sheet.Cells["1:1"] // A row
                //sheet.Cells["A:B"] // Two columns

                // Linq
                var l = sheet.Cells["A1:A5"].Where(range => range.Comment != null);

                // Dimensions used
                Assert.That(sheet.Dimension, Is.Null);

                ranger.Value = "pushing";
                var usedDimensions = sheet.Dimension;
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
                var sheet = package.Workbook.Worksheets.Add("MySheet");

                // Format as text
                sheet.Cells["A1"].Style.Numberformat.Format = "@";

                // Numbers
                sheet.SetValue("A1", "Numbers");
                Assert.That(sheet.GetValue<string>(1, 1), Is.EqualTo("Numbers"));
                sheet.Cells["B1"].Value = 15.32;
                sheet.Cells["B1"].Style.Numberformat.Format = "#,##0.00";
                // Alternatively: sheet.Cells["B1"].Formula = "FIXED(B1; 2)";
                Assert.That(sheet.Cells["B1"].Text, Is.EqualTo("15.32"));

                // Percentage
                sheet.Cells["C1"].Value = 0.5;
                sheet.Cells["C1"].Style.Numberformat.Format = "0%";
                Assert.That(sheet.Cells["C1"].Text, Is.EqualTo("50%"));

                // Money
                sheet.Cells["A2"].Value = "Moneyz";
                sheet.Cells["B2,D2"].Value = 15000.23D;
                sheet.Cells["C2,E2"].Value = -2000.50D;
                sheet.Cells["B2:C2"].Style.Numberformat.Format = "#,##0.00 [$€-813];[RED]-#,##0.00 [$€-813]";
                sheet.Cells["D2:E2"].Style.Numberformat.Format = "[$$-409]#,##0";

                // DateTime
                sheet.Cells["A3"].Value = "Timey Wimey";
                sheet.Cells["B3"].Style.Numberformat.Format = "yyyy-mm-dd";
                sheet.Cells["B3"].Formula = $"=DATE({DateTime.Now:yyyy,MM,dd})";
                sheet.Cells["C3"].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.FullDateTimePattern;
                sheet.Cells["C3"].Value = DateTime.Now;
                sheet.Cells["D3"].Style.Numberformat.Format = "dd/MM/yyyy HH:mm";
                sheet.Cells["D3"].Value = DateTime.Now;

                // Write a sortable date
                //TimeSpan diff = new DateTime(1900, 1, 1) - DateTime.Now;
                //sheet.Cells["D3"].Value = diff;
                // TODO: Still need to check which would be the best/most convenient way to handle dates


                // An external hyperlink
                sheet.Cells["C24"].Hyperlink = new Uri("https://itenium.be", UriKind.Absolute);
                sheet.Cells["C24"].Value = "Visit us";
                sheet.Cells["C24"].Style.Font.Color.SetColor(Color.Blue);
                sheet.Cells["C24"].Style.Font.UnderLine = true;

                //sheet.Cells["C25"].Formula = "HYPERLINK(\"mailto:info@itenium.be\",\"Contact support\")";
                //package.Workbook.Properties.HyperlinkBase = new Uri("");

                // An internal hyperlink
                package.Workbook.Worksheets.Add("Data");
                sheet.Cells["C26"].Hyperlink = new ExcelHyperLink("Data!A1", "Goto data sheet");

                sheet.Cells["Z1"].Clear();

                sheet.Cells.AutoFitColumns();
                package.SaveAs(new FileInfo(BinDir.GetPath()));
            }
        }

        [Test]
        public void FormattingCells()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Styling");

                // Cells with style
                ExcelFont font = sheet.Cells["A1"].Style.Font;
                sheet.Cells["A1"].Value = "Bold and proud";
                sheet.Cells["A1"].Style.Font.Name = "Arial";
                font.Bold = true;
                font.Color.SetColor(Color.Green);
                // ExcelFont also has: Size, Italic, Underline, Strike, ...

                // TODO: breaking change here in paid version
                //sheet.Cells["A3"].Style.Font.SetFromFont(new Font(new FontFamily("Arial"), 15, FontStyle.Strikeout));
                sheet.Cells["A3"].Value = "SetFromFont(Font)";

                // Borders need to be made
                sheet.Cells["A1:A2"].Style.Border.BorderAround(ExcelBorderStyle.Dotted);
                sheet.Cells[5, 5, 9, 8].Style.Border.BorderAround(ExcelBorderStyle.Dotted);

                // Merge cells
                sheet.Cells[5, 5, 9, 8].Merge = true;

                // More style
                sheet.Cells["D14"].Style.ShrinkToFit = true;
                sheet.Cells["D14"].Style.Font.Size = 24;
                sheet.Cells["D14"].Value = "Shrinking for fit";

                sheet.Cells["D15"].Style.WrapText = true;
                sheet.Cells["D15"].Value = "A wrap, yummy!";
                sheet.Cells["D16"].Value = "No wrap, ouch!";

                // Setting a background color requires setting the PatternType first
                sheet.Cells["F6:G8"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["F6:G8"].Style.Fill.BackgroundColor.SetColor(Color.Red);

                // Horizontal Alignment needs a little workaround
                // http://stackoverflow.com/questions/34660560/epplus-isnt-honoring-excelhorizontalalignment-center-or-right
                var centerStyle = package.Workbook.Styles.CreateNamedStyle("Center");
                centerStyle.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["B5"].StyleName = "Center";
                sheet.Cells["B5"].Value = "I'm centered";

                // MIGHT NOT WORK:
                sheet.Cells["B6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["B6"].Value = "I'm not centered? :(";

                // Check for an example of Conditional formatting:
                // https://github.com/JanKallman/EPPlus/wiki/Conditional-formatting

                package.SaveAs(new FileInfo(BinDir.GetPath()));
            }
        }

        [Test]
        public void ConditionalFormatting()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("ConditionalFormatting");

                #region Prepare Data
                sheet.Cells["B4"].Value = 5;
                sheet.Cells["B5"].Value = 10;
                sheet.Cells["B6"].Value = 20;
                sheet.Cells["B7"].Value = 40;
                sheet.Cells["B8"].Value = 50;
                sheet.Cells["B9"].Value = 30;
                #endregion

                ExcelAddress cfAddress1 = new ExcelAddress("B4:B9");
                var cfRule1 = sheet.ConditionalFormatting.AddTwoColorScale(cfAddress1);

                cfRule1.LowValue.Type = eExcelConditionalFormattingValueObjectType.Num;
                cfRule1.LowValue.Value = 30;
                cfRule1.LowValue.Color = Color.Blue;

                cfRule1.HighValue.Type = eExcelConditionalFormattingValueObjectType.Formula;
                cfRule1.HighValue.Formula = "MAX(B4:B9)";
                cfRule1.HighValue.Color = Color.Red;

                cfRule1.StopIfTrue = true;
                cfRule1.Style.Font.Bold = true;

                // package.SaveAs(new FileInfo(BinDir.GetPath()));
                BinDir.Save(package, false);
            }
        }


        [Test]
        public void FormattingSheetsAndColumns()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Victor");
                sheet.TabColor = Color.Ivory;

                // Freeze the top row and left 4 columns when scrolling
                sheet.View.FreezePanes(2, 5);

                sheet.View.ShowGridLines = false;
                sheet.View.ShowHeaders = false;

                //sheet.DeleteColumn();
                //sheet.InsertColumn();

                // Default selected cells when opening the xslx
                sheet.Select("B6");

                var colE = sheet.Column(5);
                //ExcelStyle colStyle = colE.Style; // See FormattingCells
                colE.AutoFit(); // or colE.Width

                // Who likes A's
                sheet.Column(1).Hidden = true;

                package.SaveAs(new FileInfo(BinDir.GetPath()));
            }
        }
    }
}
