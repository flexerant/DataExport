# DataExport

DataExport is a utility for serializing POCO collections to Excel, motivated by a need to create simple Excel spreadsheets from .NET objects. Using [NPOI](https://github.com/tonyqus/npoi) as the underlying Excel engine, it uses property attributes to specify the serializtion behaviour.

## Example

Use the ExcelSpreadsheet* attributes to describe how the object is to be serialized.

```csharp
[ExcelSpreadsheet("People")]
public class Person
{
  [ExcelSpreadsheetColumn("First name", Order = 0)]
  public string FirstName { get; set; }

  [ExcelSpreadsheetColumn("Last name", Order = 1)]
  public string LastName { get; set; }

  [ExcelSpreadsheetColumn("Date of birth", Order = 2)]
  [ExcelCellFormat(ExcelCellFormatAttribute.ShortDate)]
  public DateTime BirthDate { get; set; }
  
  [ExcelSpreadsheetColumn("Female", Order = 3)]
  public bool IsFemale { get; set; }

  [ExcelSpreadsheetIgnoreColumn()]
  public Guid UUID { get; set; } = new Guid();

  [ExcelCellFormat(ExcelCellFormatAttribute.Accounting)]
  public double Worth { get; set; }

  [ExcelCellFormat(ExcelCellFormatAttribute.Percentage)]
  public double Percent { get; set; }

  public string Text { get; set; }

  [ExcelCellFormat(ExcelCellFormatAttribute.Text)]
  public int Integer { get; set; }
}
```

The resulting Excel spreadsheet looks like this...

![excel screen shot](./assets/excel_screen_shot.PNG "Excel screen shot")


