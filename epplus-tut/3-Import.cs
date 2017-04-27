using NUnit.Framework;

namespace EPPlusTutorial
{
    //[TestFixture]
    public class Import
    {
        [Test]
        public void WritingCsv()
        {
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