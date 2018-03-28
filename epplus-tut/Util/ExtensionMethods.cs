using System.Collections;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using NUnit.Framework;
using NUnit.Framework.Constraints;
using OfficeOpenXml;

namespace EPPlusTutorial.Util
{
    public static class ExtensionMethods
    {
        // This method is called 'Where' because it is equally long as 'Cells'
        // which makes the code line out nicely.
        public static void Where(this ExcelWorksheet sheet, string cellAddress, Constraint constraint)
        {
            sheet.Calculate();
            NUnit.Framework.Assert.That(sheet.Cells[cellAddress].Value, constraint);
        }

        private static readonly IDictionary<ExcelWorksheet, int> _rowIndexes = new Dictionary<ExcelWorksheet, int>();

        /// <summary>
        /// Helper for displaying and inserting the formula on the sheet + assertion
        /// </summary>
        public static void Assert(this ExcelWorksheet sheet, string formula, Constraint constraint)
        {
            if (!_rowIndexes.TryGetValue(sheet, out int row))
            {
                _rowIndexes.Add(sheet, 2);
                row = 2;
            }

            sheet.Cells["A" + row].Value = "=" + formula.Replace(", ", "; ");
            sheet.Cells["B" + row].Formula = formula;
            sheet.Calculate();
            NUnit.Framework.Assert.That(sheet.Cells["B" + row].Value, constraint);

            _rowIndexes[sheet]++;
        }
    }
}