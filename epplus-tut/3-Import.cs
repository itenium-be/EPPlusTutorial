using System.IO;
using EPPlusTutorial.Util;
using NUnit.Framework;
using OfficeOpenXml;

namespace EPPlusTutorial
{
    //[TestFixture]
    public class Import
    {
        [Test]
        public void LoadFromCollection()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("");

                // Check Sample9.cs for example of import + Excel Table (Column TotalRowsFunction etc?)


                package.SaveAs(new FileInfo(BinDir.GetPath()));
            }
        }

        [Test]
        public void NewOne()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("");
                package.SaveAs(new FileInfo(BinDir.GetPath()));
            }
        }

        [Test]
        public void WritingCsv()
        {
            // TODO: create a sheet for each TableStyles
            //sheet.Cells["A2"].LoadFromCollection(data, true, TableStyles.Dark6);

            // Find a good CSV reader
            // download a csv
            // pretty print to Excel :)

            //ExcelRangeBase.LoadFromCollection
            // LoadFromTExt = CSV... :p


            //Bottom of: http://epplus.codeplex.com/wikipage?title=faq
            // --> create a chart

            // More advanced stuff (separate post:)
            // - Pictures
            // - Header & Footer?
            // - Printer settings
            // - Grouping
        }

        // WebApi example?
        //Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        //Response.AddHeader("content-disposition", "attachment;  filename=ExcelDemo.xlsx");
        //Response.BinaryWrite(pck.GetAsByteArray());
    }
}