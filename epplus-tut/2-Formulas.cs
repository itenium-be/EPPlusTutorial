using System;
using System.IO;
using NUnit.Framework;
using OfficeOpenXml;
using EPPlusTutorial.Util;

namespace EPPlusTutorial
{
    [TestFixture]
    public class Formulas
    {
        // also selecting like: sheet.Cells["d:d"] here
        // ...

        //Select all cells in column d between 9990 and 10000
        //        var query1 = (from cell in sheet.Cells["d:d"] where cell.Value is double && (double)cell.Value >= 9990 && (double)cell.Value <= 10000 select cell);
        //        In combination with the Range.Offset method you can also check values of other columns...


        //        //Here we use more than one column in the where clause. 
        //        //We start by searching column D, then use the Offset method to check the value of column C.
        //        var query3 = (from cell in sheet.Cells["d:d"]
        //                      where cell.Value is double &&
        //                       (double)cell.Value >= 9500 && (double)cell.Value <= 10000 &&
        //                       cell.Offset(0, -1).Value is double &&      //Column C is a double since its not a default date format.
        //                       DateTime.FromOADate((double)cell.Offset(0, -1).Value).Year == DateTime.Today.Year + 1
        //                      select cell);

        //        Console.WriteLine();
        //Console.WriteLine("Print all cells with a value between 9500 and 10000 in column D and the year of Column C is {0} ...", DateTime.Today.Year + 1);
        //Console.WriteLine();    

        // count = 0;
        ////The cells returned here will all be in column D, since that is the address in the indexer. 
        ////Use the Offset method to print any other cells from the same row.
        //foreach (var cell in query3)    
        //{
        //   Console.WriteLine("Cell {0} has value {1:N0} Date is {2:d}", cell.Address, cell.Value, DateTime.FromOADate((double)cell.Offset(0, -1).Value));
        //   count++;
        //}

        [Test]
        public void Subtotals()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Formula");
                var sheetData = package.Workbook.Worksheets.Add("Data");

                AddSomeData(sheetData);

                // 
                sheetData.Cells["A1:E4"].AutoFilter = true;

                //

                

                //Assert.That(cellA2.FullAddress, Is.EqualTo("A2"));

                //sheet.Dimension.Table
                //ConditionalFormatting

                //worksheet.Cells["A1"].Formula="CONCATENATE(\"string1_\",\"test\")";
                //.Formula = "=B3*10"

                // Absolute reference: $B$5

                // http://epplus.codeplex.com/wikipage?title=ContentSheetExample



                package.SaveAs(new FileInfo(BinDir.GetPath()));
            }
        }

        private void AddSomeData(ExcelWorksheet sheet)
        {
            //Add the headers
            sheet.Cells[1, 1].Value = "ID";
            sheet.Cells[1, 2].Value = "Product";
            sheet.Cells[1, 3].Value = "Quantity";
            sheet.Cells[1, 4].Value = "Price";
            sheet.Cells[1, 5].Value = "Value";

            //Add some items...
            sheet.Cells["A2"].Value = 12001;
            sheet.Cells["B2"].Value = "Nails";
            sheet.Cells["C2"].Value = 37;
            sheet.Cells["D2"].Value = 3.99;

            sheet.Cells["A3"].Value = 12002;
            sheet.Cells["B3"].Value = "Hammer";
            sheet.Cells["C3"].Value = 5;
            sheet.Cells["D3"].Value = 12.10;

            sheet.Cells["A4"].Value = 12003;
            sheet.Cells["B4"].Value = "Saw";
            sheet.Cells["C4"].Value = 12;
            sheet.Cells["D4"].Value = 15.37;
        }

        // TODO: http://epplus.codeplex.com/wikipage?title=Supported%20Functions&referringTitle=Documentation
        // Put all functions with link to docs
    }
}