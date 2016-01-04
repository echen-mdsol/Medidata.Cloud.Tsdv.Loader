# Medidata.Cloud.ExcelLoader
`Medidata.Cloud.ExcelLoader` is a NuGet package for object/Excel spreadsheet mapping. The mapping is bidirectional - objects to Excel and Excel to objects.

# NuGet package
http://nuget.imedidata.net/feed/mdsol/package/nuget/Medidata.Cloud.ExcelLoader
```powershell
PM> Install-Package Medidata.Cloud.ExcelLoader
```

# Usage examples of `ExcelLoader`
## Defines sheet model object
One **sheet model class** is mapping to one worksheet in Excel. Public properties within a sheet model are mapping to columns on the sheet in defined order.

A sheet model class must inherit from `SheetModel` class, and marked with a class level attribute `SheetNameAttribute` which indicates the sheet name. Property level attribute `ColumnHeaderNameAttribute` is optional whose value indicates the column header text. If this attribute isn't defined, the property name will be used as header text. Below is an example.
```cs
[SheetName("Performance")]
public class MySheet: SheetModel {
  
  public string Name {get;set;}

  public DateTime? DateOfBirth {get;set;}

  [ColumnHeaderName("Years of work experience")]
  public int WorkYears {get;set;}
}
```
## Objects to Excel
```cs
var loader = DiContainer.Resolve<ExcelLoader>();
loader.Sheet<MySheet>()
      .Data
      .Add(new MySheet{Name="Tom", DateOfBirth=new DateTime(1980,1,1)})
      .Add(new MySheet{Name="Jerry", WorkYears=11});

using (var fs = new FileStream("output.xlsx", FileMode.Create)) {
  loader.Save(fs);
}
```

## Excel to Objects
```cs
var loader = DiContainer.Resolve<ExcelLoader>();
using (var fs = new FileStream("output.xlsx", FileMode.Open)) {
  loader.Load();
}

var tom = loader.Sheet<MySheet>().Data[0];
Assert.AreEqual("Tom", tom.Name);
Assert.AreEqual(new DateTime(1980,1,1), tom.DateOfBirth);

var jerry = loader.Sheet<MySheet>().Data[1];
Assert.AreEqual("Jerry", jerry.Name);
Assert.AreEqual(11, jerry.WorkYears);
```
The sheet "Performance" in the Excel workbook will look like

| Name | DateOfBirth | Years of work experience |
| --- | --- | --- |
| Tom | 1980/1/1  |  |
| Jerry | | 11 |

## Dynamic columns
Sometimes we want to add dynamic columns. In this case we cannot predefine the columns in a sheet model class. We have to add additional column definitions on the fly.
```cs
var loader = DiContainer.Resolve<ExcelLoader>();
loader.Sheet<MySheet>()
      .Definition
      .AddColumn(new ColumnDefinition {PropertyName="2015", Header="Year of 2015"})
      .AddColumn("2016")
      .AddColumn("2017", "Year of 2017");

var kim = new MySheet{Name="Kim", WorkYears=5}
          .AddProperty("2015", "good")
          .AddProperty("2017", "excellent");
loader.Sheet<MySheet>()
      .Data
      .Add(kim);

using (var fs = new FileStream("output.xlsx", FileMode.Create)) {
  loader.Save(fs);
}

using (var fs = new FileStream("output.xlsx", FileMode.Open)) {
  loader.Load();
}

var kim2 = loader.Sheet<MySheet>().Data[0];
Assert.AreEqual("Kim", kim2.Name);
Assert.IsNull(kim2.DateOfBirth);
Assert.AreEqual(5, kim2.WorkYears);
Assert.AreEqual("good", kim2.GetExtraProperties()["2015"]);
Assert.IsNull(kim2.GetExtraProperties()["2016"]);
Assert.AreEqual("excellent", kim2.GetExtraProperties()["2015"]);
```

The sheet "Performance" in the Excel workbook will look like

| Name | DateOfBirth | Years of work experience | Year of 2015 | 2016 | Year of 2017 |
| --- | --- | --- | --- | --- | --- |
| Kim |  | 5 | good |  | excellent |

# Sheet builder decorators
Several implementations of `ISheetBuilderDecorator` are provided in this package. You can also implement your own `ISheetBuilderDecorator` and inject it into an `IExcelLoader` via `ISheetBuilder` to decorate output worksheets.

| decorator class | feature |
| --- | --- |
| `HeaderSheetDecorator` | Create header row with column names. |
| `TextStyleSheetDecorator` | Specify text style on the sheet. |
| `HeaderStyleSheetDecorator` | Specify header row text style on the sheet. Used together with `HeaderSheetDecorator`. |
| `AutoFilterSheetDecorator` | Auto filter header row. Used together with `HeaderSheetDecorator`.  |
| `AutoFitColumnSheetDecorator` | Auto fit column width |

# `unity.config` example
```xml
<?xml version="1.0" encoding="utf-8"?>
<unity xmlns="http://schemas.microsoft.com/practices/2010/unity">
  <container>
    <register type="Medidata.Cloud.ExcelLoader.ISheetBuilderDecorator, Medidata.Cloud.ExcelLoader"
              mapTo="Medidata.Cloud.ExcelLoader.SheetDecorators.TextStyleSheetDecorator, Medidata.Cloud.ExcelLoader"
              name="textStyleDecorator">
      <constructor>
        <param name="styleName" value="Normal" />
      </constructor>
    </register>
    <register type="Medidata.Cloud.ExcelLoader.ISheetBuilderDecorator, Medidata.Cloud.ExcelLoader"
              mapTo="Medidata.Cloud.ExcelLoader.SheetDecorators.HeaderStyleSheetDecorator, Medidata.Cloud.ExcelLoader"
              name="headerStyleDecorator">
      <constructor>
        <param name="styleName" value="Output" />
      </constructor>
    </register>
    <register type="Medidata.Cloud.ExcelLoader.ICellTypeValueConverterManager, Medidata.Cloud.ExcelLoader"
              mapTo="Medidata.Cloud.ExcelLoader.CellTypeValueConverterManager, Medidata.Cloud.ExcelLoader">
      <lifetime type="singleton" />
    </register>
    <register type="Medidata.Cloud.ExcelLoader.ISheetBuilder, Medidata.Cloud.ExcelLoader"
              mapTo="Medidata.Cloud.ExcelLoader.SheetBuilder, Medidata.Cloud.ExcelLoader">
      <constructor>
        <param name="converterManager" dependencyType="Medidata.Cloud.ExcelLoader.ICellTypeValueConverterManager, Medidata.Cloud.ExcelLoader" />
        <param name="decorators">
          <array>
            <dependency type="Medidata.Cloud.ExcelLoader.SheetDecorators.HeaderSheetDecorator, Medidata.Cloud.ExcelLoader" />
            <dependency name="textStyleDecorator" />
            <dependency name="headerStyleDecorator" />
            <dependency type="Medidata.Cloud.ExcelLoader.SheetDecorators.AutoFilterSheetDecorator, Medidata.Cloud.ExcelLoader" />
            <dependency type="Medidata.Cloud.ExcelLoader.SheetDecorators.AutoFitColumnSheetDecorator, Medidata.Cloud.ExcelLoader" />
          </array>
        </param>
      </constructor>
    </register>
    <register type="Medidata.Cloud.ExcelLoader.ISheetParser, Medidata.Cloud.ExcelLoader"
              mapTo="Medidata.Cloud.ExcelLoader.SheetParser, Medidata.Cloud.ExcelLoader" />
    <register type="Medidata.Cloud.ExcelLoader.IExcelBuilder, Medidata.Cloud.ExcelLoader"
              mapTo="Medidata.Cloud.ExcelLoader.ExcelBuilder, Medidata.Rave.Tsdv.Loader" />
    <register type="Medidata.Cloud.ExcelLoader.IExcelParser, Medidata.Cloud.ExcelLoader"
              mapTo="Medidata.Cloud.ExcelLoader.ExcelParser, Medidata.Cloud.ExcelLoader" />
    <register type="Medidata.Cloud.ExcelLoader.IExcelLoader, Medidata.Cloud.ExcelLoader"
              mapTo="Medidata.Cloud.ExcelLoader.ExcelLoader, Medidata.Cloud.ExcelLoader" />
  </container>
</unity>
```
