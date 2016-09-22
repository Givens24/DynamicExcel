# DynamicExcel
A generic .Net Library used to export a collection of any class to an Excel spreadsheet. (NOTE: Nested collections are not supported)

# Setup and Usage for DynamicExcel
* All you need to do in order to execute a simple export with DynamicExcel is create an instance of the **DynamicExcelExport** class
* the DynamicExcelExport constructor requires an IEnumerable<T> which is the data that you intend to export.
* The "Export" method takes one required parameter which is the file path of the new excel spreadsheet. The second and third parameters are optional and are for adding a name to the worksheet or accessing the index of a specific worksheet that exists in the workbook. (The defaults for the **worksheetName** and **worksheetIndex** parameters are "" and 1. If no **worksheetName** is specified, the Excel default "Sheet(n)" will be used)
* The following code demonstrates how to call the class's export method and the class's constructor which takes an IEnumerable<T> as a parameter.

 ```C#
//view model data to pass to the DynamicExcelExport
var contactsViewModel = new List<ContactInfoViewModel>();
var contactInfoViewModel = new ContactInfoViewModel{ FirstName = "John", LastName = "Doe", Age = 36 };
contactsViewModel.Add(contactInfoViewModel);
var testFileName = "C:\\testFile.xlsx";
                                            `
var dynamicExcelExport = new DynamicExcelExport<ContactInfoViewModel>(contactsViewModel);
dynamicExcelExport.Export(testFileName, "Contact Sheet", 1);
```

***

# Formatting the Excel Column Headers
* Formatting the spreadsheet's columns is all done through custom attributes with DynamicExcel. The following attributes need to be added to the class's properties in order to format the column headers.
* In order to change the **entire** column's width, use the **ColumnWidth** attribute which takes an int parameter.
```C#
[ColumnWidth(20)]
public string FirstName { get; set; }
```
* In order to change text alignment, use the **TextAlignment** attribute which takes two parameters. The first parameter is an XlHalign property (**NOTE**: You will need to add a using statement reference Microsoft.Office.Interop.Excel to use the XlHAlign enum.) and the second is a boolean that specifies if you want the entire column aligned a specific way. The default functionality will simply align the column header.
```C#
[TextAlignment(Excel.XlHAlign.xlHAlignCenter, true)]
public Guid ProductId { get; set; }
```
* To change the **column header's** font color, use the **HeaderFontColor** attribute which takes an XlRgbColor.
(**NOTE**: You will need to add a using statement reference Microsoft.Office.Interop.Excel to use the XlRgbColor enum.)
```C#
[HeaderFontColor(XlRgbColor.White)]
public string Address { get; set; }
```
* To change the **column header's** background color, use the **HeaderBackgroundColor** attribute which takes an XlRgbColor parameter.
```C#
[HeaderBackgroundColor(XlRgbColor.NavyBlue)]
public string ZipCode { get; set; }
```
* To change the **column header's** font size, use the **FontSize** attribute which takes an int parameter.
```C#
[FontSize(16)]
public string LastName { get; set; }
```

***

