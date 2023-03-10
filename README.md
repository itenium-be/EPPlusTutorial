EPPlusTutorial
==============

### Attention: EPPlus is no longer free for commercial use


The UnitTests can be executed to generate Excels in `bin/Debug/excels`.
(Some tests even contain assertions!)

```powershell
# Install last free version:
Install-Package EPPlus -Version 4.5.3

# Install commercial version:
Install-Package EPPlus
```

1-QuickTutorial
---------------
- Basic usage
- Loading & saving
- Selecting cells
- Writing values
- Formatting sheets, columns and cells
- Conditional formatting

2-Formulas-DataValidation
----------
- BasicFormulas
- DataValidation
- Attaching a logger to the FormulaParser

3-Import
--------
Loading data from
- LoadFromCollection & LoadFromArrays (IEnumerable)
- LoadFromDataTable & DataReader
- LoadFromText (CSV)

4-Miscellaneous
---------------
- Printing
- Workbook properties
- Converting Excel Addresses
- Adding comments & rich text
- Protection against edit

5-Formulas-Reference
----------
- String manipulation
- Numbers & Math
- Date & Time
- Boolean logic

Charts
------
Don't seem to work for LibreOffice. Example code can be found in the official EPPlus examples.

- Sample4.cs: Basic example
- Sample5.cs: A pie
- Sample6.cs: Pretty nifty, worth checking out!

[They also have a chart example on their main documentation][chart-github].

Wish List
---------
Will we cover these also, sometime?

- ConditionalFormatting in more detail: See [Sample14.cs][github-sample-cf]
- Filtering
- Grouping and ungrouping
- Tables
- Inserting VBA: See Sample15.cs
- Numberformat.Format = [$$-409] --> Get info on those numbers
- 1-QuickTutorial ExcelPrinting: insert a company picture in the print footer: sheet.HeaderFooter.EvenHeader.InsertPicture()
- Password protection: Add Encryption
- WebApi integration and calling code for Superagent, Angular Http and Fetch? :)


**WebApi**:  
```c#
public class ExcelResult : ActionResult
{
    public string FileName { get; set; }
    public ExcelPackage Package { get; set; }

    public override void ExecuteResult(ControllerContext context)
    {
        context.HttpContext.Response.Buffer = true;
        context.HttpContext.Response.Clear();
        context.HttpContext.Response.AddHeader("content-disposition", "attachment; filename=" + FileName);
        context.HttpContext.Response.ContentType = "application/vnd.ms-excel";
        context.HttpContext.Response.BinaryWrite(Package.GetAsByteArray());
    }
}
```

**Adding a picture**:  
```
Bitmap icon = GetIcon(dir.FullName);
ws.Row(row).Height = height;
if (icon != null)
{
    ExcelPicture pic = ws.Drawings.AddPicture("pic" + (row).ToString(), icon);
    pic.SetPosition((int)20 * (row - 1) + 2, 0);
}
```

[chart-github]: https://github.com/JanKallman/EPPlus/wiki/Shapes,-Pictures-and-Charts
[github-sample-cf]: https://github.com/JanKallman/EPPlus/blob/master/SampleApp/Sample14.cs
