using System.IO;
using EPPlusTutorial.Util;
using NUnit.Framework;
using OfficeOpenXml;

namespace EPPlusTutorial
{
    /// <summary>
    /// See <see cref="OfficeOpenXml.FormulaParsing.Excel.Functions"/> for all supported formulas
    /// 
    /// Assign a formula with either
    /// - .Formula = "A$5"
    /// - .FormulaR1C1 = "RC[-2]*RC[-1]"
    /// 
    /// Note:
    /// - Do not start formula with =
    /// - Use English function names
    /// - Use , as function argument separator
    /// 
    /// Troubles? See <see cref="FormulasAndDataValidation.TroubleshootingFormulas"/>
    /// </summary>
    [TestFixture]
    public class Formulas
    {
        private const string Fox = "The quick brown fox jumps over the lazy dog";

        [Test]
        public void StringManipulation()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("StringManipulation");
                sheet.Cells["A1"].Value = Fox;

                sheet.Cells["A2"].Value = "=LEN(A1)";
                sheet.Cells["B2"].Formula = "LEN(A1)"; // Formulas do not start with =

                sheet.Cells["A3"].Value = "=UPPER(A1)";
                sheet.Cells["B3"].Formula = "UPPER(A1)"; // also LOWER()
                sheet.Cells["C3"].Formula = "PROPER(A1)"; // = ToTitleCase()

                // LibreOffice uses ; as function argument separation
                sheet.Cells["A4"].Value = "=LEFT(A1; 3)";
                // But EPPlus works with ,
                sheet.Cells["B4"].Formula = "LEFT(A1, 3)"; // also RIGHT()

                sheet.Cells["A5"].Value = "=MID(A1; 5; 5)";
                sheet.Cells["B5"].Formula = "MID(A1, 5, 5)"; // !! String indexes are 1 based !!

                sheet.Cells["A6"].Value = "=REPLACE(A1; 1; 3; \"A\")"; // Replace text with indexes
                sheet.Cells["B6"].Formula = "REPLACE(A1, 1, 3, \"A\")";

                sheet.Cells["A7"].Value = "=SUBSTITUTE(LOWER(A1); \"the\"; \"a\")"; // Replace text (case sensitive)
                sheet.Cells["B7"].Formula = "SUBSTITUTE(LOWER(A1), \"the\", \"a\")";

                sheet.Cells["A8"].Value = "=REPT(A1; 1; 3; \"A\")"; // Repeat
                sheet.Cells["B8"].Formula = "REPT(\"A\", 3)";

                sheet.Cells["A9"].Value = "=CONCATENATE(A1; \" over and\"; \" over again\")"; // accepts x arguments
                sheet.Cells["B9"].Formula = "CONCATENATE(A1, \" over and over again\")";

                // return index of needle in text
                sheet.Cells["A10"].Value = "=FIND(\"fox\"; A1)"; // find_text, text, startingPosition (case sensitive)
                sheet.Cells["B10"].Formula = "FIND(\"fox\", A1)";
                sheet.Cells["C10"].Formula = "FIND(\"FOX\", A1)"; // Not found: #VALUE!
                sheet.Cells["D10"].Formula = "SEARCH(\"FOX\", A1)"; // not case sensitive

                // returns the text itself, if it is a string
                sheet.Cells["A11"].Value = "=T(A1)"; // typeof A1 === "string" ? A1 : ""
                sheet.Cells["B11"].Formula = "T(A1)";

                sheet.Calculate();

                Assert.That(sheet.Cells["B2"].GetValue<int>(), Is.EqualTo(Fox.Length));
                Assert.That(sheet.Cells["B3"].Value, Is.EqualTo(Fox.ToUpper()));
                Assert.That(sheet.Cells["B4"].Value, Is.EqualTo(Fox.Substring(0, 3)));
                Assert.That(sheet.Cells["B5"].Value, Is.EqualTo("quick"));
                Assert.That(sheet.Cells["B6"].Value, Is.EqualTo("A quick brown fox jumps over the lazy dog"));
                Assert.That(sheet.Cells["B7"].Value, Is.EqualTo("a quick brown fox jumps over a lazy dog"));
                Assert.That(sheet.Cells["B8"].Value, Is.EqualTo("AAA"));
                Assert.That(sheet.Cells["B9"].Value, Is.EqualTo(Fox + " over and over again"));
                Assert.That(sheet.Cells["B10"].Value, Is.EqualTo(17));

                // HYPERLINK: Also see <see cref="QuickTutorial.WritingValues"/> for a hyperlink?

                sheet.Column(1).Width = 50;
                package.SaveAs(new FileInfo(BinDir.GetPath()));
            }
        }

        [Test]
        public void Math()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Math");
                sheet.Cells["A1"].Value = Fox;

                // Convertion, Formatting?
                // =VALUE("15.32") Converts to number
                // =INT(15.32) Converts to int (Math.Floor)
                // =TEXT(1554543,5154; "#.###,00")
                // =FIXED(1554543,5154; 2) - format number (number, decimals) <see cref="QuickTutorial.WritingValues"/>

                // TODO: stuff with boolean logic?

                // returns true if the strings are the same (case sensitive)
                sheet.Cells["A16"].Value = "=EXACT(A1; \"The quick brown fox jumps over the lazy dog\")";
                sheet.Cells["B16"].Formula = "EXACT(A1, \"The quick brown fox jumps over the lazy dog\")";


                sheet.Calculate();

                Assert.That(sheet.Cells["B16"].Value, Is.EqualTo(true));



                sheet.Column(1).Width = 50;
                package.SaveAs(new FileInfo(BinDir.GetPath()));
            }
        }
    }
}