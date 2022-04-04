# C# Class / List to excel tools

Simply convert your custom class/list class to excel file.  
  
usage:  

```cs

ExcelObject excel = new ExcelObject();

excel.initExcelPackage();
excel.AddWorkSheet<MyClass>(results, "TestWorkBook");
excel.Save("out\\result.xlsx");

```  

