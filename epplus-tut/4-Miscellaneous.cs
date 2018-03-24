using System.IO;
using EPPlusTutorial.Util;
using NUnit.Framework;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace EPPlusTutorial
{
    [TestFixture]
    public class Miscellaneous
    {
        [Test]
        public void ExcelPrinting()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Printing");
                sheet.Cells["A1"].Value = "Check the print preview (Ctrl+P)";

                var header = sheet.HeaderFooter.OddHeader;
                // &24: Font size
                // &U: Underlined
                // &"": Font name
                header.CenteredText = "&24&U&\"Arial,Regular Bold\" YourTitle";
                header.RightAlignedText = ExcelHeaderFooter.CurrentDate;
                header.LeftAlignedText = ExcelHeaderFooter.SheetName;

                ExcelHeaderFooterText footer = sheet.HeaderFooter.OddFooter;
                footer.RightAlignedText = $"Page {ExcelHeaderFooter.PageNumber} of {ExcelHeaderFooter.NumberOfPages}";
                footer.CenteredText = ExcelHeaderFooter.SheetName;
                footer.LeftAlignedText = ExcelHeaderFooter.FilePath + ExcelHeaderFooter.FileName;

                sheet.PrinterSettings.RepeatRows = sheet.Cells["1:2"];
                sheet.PrinterSettings.RepeatColumns = sheet.Cells["A:G"];

                // Change the sheet view
                // (Did not work in LibreOffice5)
                sheet.View.PageLayoutView = true;

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

        [Test]
        public void SettingWorkbookProperties()
        {
            using (var package = new ExcelPackage())
            {
                package.Workbook.Properties.Title = "EPPlus Tutorial Series";
                package.Workbook.Properties.Author = "Wouter Van Schandevijl";
                package.Workbook.Properties.Comments = "";
                package.Workbook.Properties.Keywords = "";
                package.Workbook.Properties.Category = "";

                package.Workbook.Properties.Company = "itenium";
                package.Workbook.Properties.SetCustomPropertyValue("Checked by", "Jan Källman");

                package.Workbook.Worksheets.Add("Sheet1");
                package.SaveAs(new FileInfo(BinDir.GetPath()));
            }
        }

        [Test]
        public void AddingCommentsWithRichText()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Rich Comments");

                ExcelComment comment = sheet.Cells["A1"].AddComment("Bold title:\r\n", "evil corp");
                comment.Font.Bold = true;
                comment.AutoFit = true;

                ExcelRichText rt = comment.RichText.Add("Unbolded subtext");
                rt.Bold = false;

                // A more extensive example can be found in Sample6.cs::AddComments of the official examples project
                // https://github.com/JanKallman/EPPlus/blob/master/SampleApp/Sample6.cs

                package.SaveAs(new FileInfo(BinDir.GetPath()));
            }
        }

        [Test]
        public void PasswordProtectionFromManualEditing()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Secret");

                // Block entire sheet except D5
                sheet.Cells["D5"].Value = "Can't touch this";
                sheet.Cells["D5"].Style.Locked = false;

                sheet.Protection.AllowDeleteRows = false;
                sheet.Protection.SetPassword("Secret");

                //sheet.Protection.IsProtected = true;

                // Or if you need more serious locking
                //var book = package.Workbook;
                //book.Protection.LockWindows = true;
                //book.Protection.LockStructure = true;
                //book.View.ShowHorizontalScrollBar = false;
                //book.View.ShowVerticalScrollBar = false;
                //book.View.ShowSheetTabs = false;

                package.SaveAs(new FileInfo(BinDir.GetPath()));
            }
        }
    }
}