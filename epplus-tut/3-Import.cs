using System;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using EPPlusTutorial.Util;
using NUnit.Framework;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace EPPlusTutorial
{
    [TestFixture]
    public class Import
    {
        [Test]
        public void LoadFromCollection()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Collection");
                var data = new[]
                {
                    new {Name = "A", Value = 1},
                    new {Name = "B", Value = 2},
                    new {Name = "C", Value = 3},
                };
                sheet.Cells["A2"].LoadFromCollection(data);
                sheet.Cells["A1"].SetHeaders("Name", "Value");
                package.SaveAs(new FileInfo(BinDir.GetPath()));
            }
        }

        [Test]
        public void LoadFromArrays()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Arrays");
                var data = new[]
                {
                    new[] {"A1", "B1", "C1"},
                    new[] {"A2", "B2", "C3"},
                };
                sheet.Cells["A1"].LoadFromArrays(data);
                package.SaveAs(new FileInfo(BinDir.GetPath()));
            }
        }

        [Test]
        public void LoadFromDataTable()
        {
            // Also: LoadFromDataReader()

            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("DataTable");
                var data = new DataTable();

                

                sheet.Cells["A1"].LoadFromDataTable(data, true);
                package.SaveAs(new FileInfo(BinDir.GetPath()));
            }
        }

        [Test]
        public void LoadFromText()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("CSV");
                var file = new FileInfo(BinDir.GetPath("LoadFromText.csv"));

                var format = new ExcelTextFormat()
                {
                    Delimiter = ',',
                    Culture = CultureInfo.InvariantCulture,
                    TextQualifier = '"'
                    // EOL, DataTypes, Encoding, SkipLinesBeginning/End
                };
                sheet.Cells["A1"].LoadFromText(file, format);
                package.SaveAs(new FileInfo(BinDir.GetPath()));
            }
        }

        // Check Sample9.cs for example of import + Excel Table (Column TotalRowsFunction etc?)

        [Test]
        public void TableStyles()
        {
            var data = new[]
            {
                new {Name = "A", Value = 1},
                new {Name = "B", Value = 2},
                new {Name = "C", Value = 3},
            };

            var tableStyles = Enum.GetValues(typeof(TableStyles)).OfType<TableStyles>();
            using (var package = new ExcelPackage())
            foreach (var tableStyle in tableStyles)
            {
                    var sheet = package.Workbook.Worksheets.Add(tableStyle.ToString());
                    sheet.Cells["A1"].LoadFromCollection(data, true, tableStyle);
                    package.SaveAs(new FileInfo(BinDir.GetPath()));
            }
        }
    }
}