using System;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
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
    public class FormulasReference
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
                sheet.Where("B2", Is.EqualTo(Fox.Length));

                sheet.Cells["A3"].Value = "=UPPER(A1)";
                sheet.Cells["B3"].Formula = "UPPER(A1)"; // also LOWER()
                sheet.Where("B3", Is.EqualTo(Fox.ToUpper()));
                sheet.Cells["C3"].Formula = "PROPER(A1)"; // = ToTitleCase()
                sheet.Where("C3", Is.EqualTo("The Quick Brown Fox Jumps Over The Lazy Dog"));

                // LibreOffice uses ; as function argument separation
                sheet.Cells["A4"].Value = "=LEFT(A1; 3)";
                // But EPPlus works with ,
                sheet.Cells["B4"].Formula = "LEFT(A1, 3)"; // also RIGHT()
                sheet.Where("B4", Is.EqualTo(Fox.Substring(0, 3)));

                sheet.Cells["A5"].Value = "=MID(A1; 5; 5)";
                sheet.Cells["B5"].Formula = "MID(A1, 5, 5)"; // !! String indexes are 1 based !!
                sheet.Where("B5", Is.EqualTo(Fox.Substring(4, 5)));

                sheet.Cells["A6"].Value = "=REPLACE(A1; 1; 3; \"A\")"; // Replace text with indexes
                sheet.Cells["B6"].Formula = "REPLACE(A1, 1, 3, \"A\")";
                sheet.Where("B6", Is.EqualTo("A quick brown fox jumps over the lazy dog"));

                sheet.Cells["A7"].Value = "=SUBSTITUTE(LOWER(A1); \"the\"; \"a\")"; // Replace text (case sensitive - but LOWER(A1))
                sheet.Cells["B7"].Formula = "SUBSTITUTE(LOWER(A1), \"the\", \"a\")";
                sheet.Where("B7", Is.EqualTo(Regex.Replace(Fox, "the", "a", RegexOptions.IgnoreCase)));

                sheet.Cells["A8"].Value = "=REPT(A1; 1; 3; \"A\")"; // Repeat
                sheet.Cells["B8"].Formula = "REPT(\"A\", 3)";
                sheet.Where("B8", Is.EqualTo("AAA"));

                sheet.Cells["A9"].Value = "=CONCATENATE(A1; \" over and\"; \" over again\")"; // accepts x parameters
                sheet.Cells["B9"].Formula = "CONCATENATE(A1, \" over and over again\")"; // or use & to concatenate
                sheet.Where("B9", Is.EqualTo(Fox + " over and over again"));

                // FIND/SEARCH: return index of needle in text
                sheet.Cells["A10"].Value = "=FIND(\"fox\"; A1)"; // find_text, text, startingPosition (case sensitive)
                sheet.Cells["B10"].Formula = "FIND(\"fox\", A1)";
                sheet.Where("B10", Is.EqualTo(Fox.IndexOf("fox") + 1));
                sheet.Cells["C10"].Formula = "FIND(\"FOX\", A1)"; // Not found: #VALUE!
                sheet.Cells["D10"].Formula = "SEARCH(\"FOX\", A1)"; // not case sensitive
                sheet.Where("D10", Is.EqualTo(Fox.IndexOf("fox") + 1));

                // returns the text itself, if it is a string
                sheet.Cells["A11"].Value = "=T(A1)"; // typeof A1 === "string" ? A1 : ""
                sheet.Cells["B11"].Formula = "T(A1)";
                sheet.Where("B11", Is.EqualTo(Fox));

                // returns true if the strings are the same (case sensitive)
                sheet.Cells["A12"].Value = $"=EXACT(A1, \"{Fox}\")";
                sheet.Cells["B12"].Formula = $"EXACT(A1; \"{Fox}\")"; ;
                sheet.Where("B12", Is.EqualTo(true));

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
                sheet.Cells["B1"].Value = "15,32";
                sheet.Cells["C1"].Value = 15.32;

                // Formatting:
                // =TEXT(1554543,5154; "#.###,00")
                // =FIXED(1554543,5154; 2) - format number (number, decimals) <see cref="QuickTutorial.WritingValues"/>

                // Convert
                sheet.Assert("VALUE(B1)", Is.EqualTo(15.32));
                sheet.Assert("INT(\"15.62\")", Is.EqualTo(15)); // or FLOOR(number, significance) or ROUNDDOWN or TRUNC
                // CEILING(number, significance) or ROUNDUP
                // ROUND(number, significance)

                // ISNUMBER, ISEVEN, ISODD
                // MAX, MIN
                // COUNT(range) - Counts all numeric cell values
                // COUNTA(range) - Counts all non empty cell values
                // COUNTBLANK(range)
                // COUNTIF(range, citeria)
                // COUNTIFS(range1, criteria1, range2, criteria2, ...)

                // Criteria possibilities for IF:
                // A literal, another cell, ">=10", "<>0"
                // "<>"&A1 - not equal to A1
                // "gr?y" - single letter wildcard
                // "cat*" - 0..x wildcard

                // SUM, SUMIF, SUMIFS
                // AVERAGE, AVERAGEIF, AVERAGEIFS

                // ABS(number) - abolute value
                // SIGN(number) - returns -1 or 1
                // PRODUCT(range...) - returns arg1 * arg2 * ...
                // POWER(base, exponent) - Or base^exp. Also: SQRT
                // MOD(divident, divisor) - modulo. Also: QUOTIENT
                // RAND() - between 0 and 1
                // RANDBETWEEN(lowest, highest) - both params inclusive
                // LARGE(range, xth) - returns xth largest number; also SMALL()

                // PI, SIN, COS, ASIN, ASINH, TAN, ATAN, ...
                // EXP, LOG, LOG10, LN
                // MEDIAN, STDEV, RANK, VAR

                sheet.Column(1).Width = 50;
                package.SaveAs(new FileInfo(BinDir.GetPath()));
            }
        }

        [Test]
        public void Logical()
        {
            // Information
            // ISBLANK, ISTEXT, ISNONTEXT

            // Booleans
            // ISLOGICAL - checks if is boolean
            // TRUE() and FALSE()

            // IF(condition, ifTrue, ifFalse) also: SWITCH, IFS (MS Excel 2016)
            // IF(A1="value", "value", "not value")

            // OR, AND, NOT
            // IF(OR(A1="value", A1="value2"), "value1-2", "other")

            // Check if is blank
            // IF(OR(ISBLANK(A1), TRIM(A1)=""), 1, 0)
        }

        [Test]
        public void DateAndTime()
        {
            // DATE(year, month, day)
            // TODAY()-2 - DateTime.Now.Date.Subtract(TimeSpan.FromDays(2))
            // NOW()+"2:00" - DateTime.Now.Add(TimeSpan.FromHours(2))
            // DATE(date) - DateTime.Now.Day
            // MONTH, YEAR, TIME, HOUR, MINUTE, SECOND
            // WEEKNUM - week of year. Also: ISOWEEKNUM
            // WEEKDAY - day of week index (sunday=0)

            // DAYS360(date1, date2) - Difference in days
            // YEARFRAC(date1, date2) - Difference in years (including fractional part)

            // EDATE(date, nrOfMonths) - add nrOfMonths to date
            // EOMONTHS(date, 0) - returns last day of the date month
            // EOMONTHS(date, -2) - returns the last day of (date - 2 months) month
            // WORKDAY(date, workDaysToAdd, holidaysRange) - holidaysRange is optional

            //using (var package = new ExcelPackage())
            //{
            //    var sheet = package.Workbook.Worksheets.Add("DateAndTime");
            //    sheet.Cells["A1"].Value = Fox;
            //    sheet.Assert("TODAY()", Is.EqualTo(DateTime.Now.Date));
            //    package.SaveAs(new FileInfo(BinDir.GetPath()));
            //}
        }
    }
}