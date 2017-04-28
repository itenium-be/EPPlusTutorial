using System;
using System.Collections;
using System.IO;
using System.Linq;
using System.Reflection;
using NUnit.Framework;
using OfficeOpenXml;
using EPPlusTutorial.Util;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System.Collections.Generic;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.DataValidation.Contracts;

namespace EPPlusTutorial
{
    [TestFixture]
    public class FormulasAndDataValidation
    {
        [Test]
        public void BasicFormulas()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Formula");

                // SetHeaders is an extension method
                sheet.Cells["A1"].SetHeaders("Id", "Product", "Quantity", "Price", "Base total", "Discount", "Total", "Special discount", "Payup");
                sheet.Cells["H5"].AddComment("Special discount for our most valued customers", "evil corp");

                // Turn filtering on for the headers
                sheet.Cells[1, 1, 1, sheet.Dimension.End.Column].AutoFilter = true;

                var data = AddThreeRowsDataAndFormat(sheet);

                // Starting = is optional
                sheet.Cells["A5"].Formula = "=COUNT(A2:A4)";
                // Hide the formula (when the sheet.IsProtected)
                sheet.Cells["A5"].Style.Hidden = true;

                // Total column
                sheet.Cells["E2:E4"].Formula = "C2*D2"; // quantity * price
                Assert.That(sheet.Cells["E2"].FormulaR1C1, Is.EqualTo("RC[-2]*RC[-1]"));
                Assert.That(sheet.Cells["E4"].FormulaR1C1, Is.EqualTo("RC[-2]*RC[-1]"));

                // Total - discount column
                // Calculate formulas before they are available in the sheet
                // (Opening an Excel with Office will do this automatically)
                sheet.Cells["G2:G4"].Formula = "IF(ISBLANK(F2),E2,E2*(1-F2))";
                Assert.That(sheet.Cells["G2"].Text, Is.Empty);
                sheet.Calculate();
                Assert.That(sheet.Cells["G2"].Text, Is.Not.Empty);

                // Total row
                // R1C1 reference style
                sheet.Cells["E5"].FormulaR1C1 = "SUBTOTAL(9,R[-3]C:R[-1]C)"; // total
                Assert.That(sheet.Cells["E5"].Formula, Is.EqualTo("SUBTOTAL(9,E2:E4)"));
                sheet.Cells["G5"].FormulaR1C1 = "SUBTOTAL(9,R[-3]C:R[-1]C)"; // total - discount
                Assert.That(sheet.Cells["G5"].Formula, Is.EqualTo("SUBTOTAL(9,G2:G4)"));

                sheet.Calculate();
                sheet.Cells["I2:I5"].Formula = "G2*(1-$H$5)"; // Pin H5

                // SUBTOTAL(9 = SUM) // 109 = Sum excluding manually hidden rows
                // AVERAGE (1), COUNT (2), COUNTA (3), MAX (4), MIN (5)
                // PRODUCT (6), STDEV (7), STDEVP (8), SUM (9), VAR (10)

                sheet.Cells.AutoFitColumns();
                package.SaveAs(new FileInfo(BinDir.GetPath()));
            }
        }

        private ICollection<Sell> AddThreeRowsDataAndFormat(ExcelWorksheet sheet)
        {
            var data = new SalesGenerator().Generate(3).ToArray();
            // See 3-Import for more about LoadFromXXX
            sheet.Cells["A2"].LoadFromCollection(data, false);

            // Special discount
            sheet.Cells["H5"].Value = 0.2;
            sheet.Cells["H5"].Style.Numberformat.Format = "0%";

            // Formatting is covered in 1-QuickTutorial
            sheet.Cells["C2:C5"].Style.Numberformat.Format = "#,##0"; // number
            sheet.Cells["D2:E5,G2:G5,I2:I5"].Style.Numberformat.Format = "[$$-409]#,##0.00"; // money
            sheet.Cells["F2:F5"].Style.Numberformat.Format = "0%"; // percentage

            // Border above the totals row
            var lastCell = sheet.Dimension.End;
            sheet.Cells[lastCell.Row, 1, lastCell.Row, lastCell.Column].Style.Border.Top.Style = ExcelBorderStyle.Double;

            return data;
        }

        [Test]
        public void ImplementedFormulaFunctions()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Funcs");

                var funcs = package.Workbook.FormulaParserManager.GetImplementedFunctions()
                    .OrderBy(x => x.Value.GetType().FullName)
                    .ThenBy(x => x.Key)
                    .Select(x => new { FunctionName = x.Key, TypeName = x.Value.GetType().FullName, x.Value.IsErrorHandlingFunction, x.Value.IsLookupFuction });

                sheet.Cells.LoadFromCollection(funcs, true);

                sheet.Cells.AutoFitColumns();
                package.SaveAs(new FileInfo(BinDir.GetPath()));
            }
        }

        [Test]
        public void DataValidation_DropDownComboCell()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Validation");

                var list1 = sheet.Cells["C7"].DataValidation.AddListDataValidation();
                list1.Formula.Values.Add("Apples");
                list1.Formula.Values.Add("Oranges");
                list1.Formula.Values.Add("Lemons");

                list1.ShowErrorMessage = true;
                list1.Error = "We only have those available :(";

                list1.ShowInputMessage = true;
                list1.PromptTitle = "Choose your juice";
                list1.Prompt = "Apples, oranges or lemons?";

                list1.AllowBlank = true;

                sheet.Cells["C7"].Value = "Pick";
                package.SaveAs(new FileInfo(BinDir.GetPath()));
            }
        }

        [Test]
        public void DataValidation_FromOtherSheet()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Validation");

                var otherSheet = package.Workbook.Worksheets.Add("OtherSheet");
                otherSheet.Cells["A1"].Value = "Kwan";
                otherSheet.Cells["A2"].Value = "Nancy";
                otherSheet.Cells["A3"].Value = "Tonya";

                var list1 = sheet.Cells["C7"].DataValidation.AddListDataValidation();
                list1.Formula.ExcelFormula = "OtherSheet!A1:A4";

                package.SaveAs(new FileInfo(BinDir.GetPath()));
            }
        }

        [Test]
        public void DataValidation_IntAndDateTime()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("intsAndSuch");

                // Integer validation
                IExcelDataValidationInt intValidation = sheet.DataValidations.AddIntegerValidation("A1");
                intValidation.Prompt = "Value between 1 and 5";
                intValidation.Operator = ExcelDataValidationOperator.between;
                intValidation.Formula.Value = 1;
                intValidation.Formula2.Value = 5;

                // DateTime validation
                IExcelDataValidationDateTime dateTimeValidation = sheet.DataValidations.AddDateTimeValidation("A2");
                dateTimeValidation.Prompt = "A date greater than today";
                dateTimeValidation.Operator = ExcelDataValidationOperator.greaterThan;
                dateTimeValidation.Formula.Value = DateTime.Now.Date;

                // Time validation
                IExcelDataValidationTime timeValidation = sheet.DataValidations.AddTimeValidation("A3");
                timeValidation.Operator = ExcelDataValidationOperator.greaterThan;
                var time = timeValidation.Formula.Value;
                time.Hour = 13;
                time.Minute = 30;
                time.Second = 10;

                // Existing validations
                var validations = package.Workbook.Worksheets.SelectMany(sheet1 => sheet1.DataValidations);

                package.SaveAs(new FileInfo(BinDir.GetPath()));
            }
        }

        [Test]
        public void TroubleshootingFormulas()
        {
            using (var package = new ExcelPackage())
            {
                var logfile = new FileInfo(BinDir.GetPath("TroubleshootingFormulas.txt"));
                package.Workbook.FormulaParserManager.AttachLogger(logfile);
                package.Workbook.Calculate();
                package.Workbook.FormulaParserManager.DetachLogger();
            }
        }
    }
}