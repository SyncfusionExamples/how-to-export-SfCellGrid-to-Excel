# How-to-export-SfCellGrid-to-Excel

This example demonstartes that how to export `SfCellGrid` to Excel.

SfCellGrid does not have built-in function to export the grid data to excel. To export the data into excel, create a new excel file using the `ExcelEngine` and set the grid data to worksheet cells using `IRange.Value` property and save that modified workbook.

``` c#
//Convert to excel
ExcelEngine excelEngine = new ExcelEngine();
excelEngine.Excel.DefaultVersion = ExcelVersion.Excel2016;
IApplication application = excelEngine.Excel;
 
IWorkbook workbook = application.Workbooks.Create(1);
 
IWorksheet sheet = workbook.Worksheets[0];
 
for (int i = 0; i < grid.RowCount; i++)
{
    for (int j = 0; j < grid.ColumnCount; j++)
    {
        IRange range = sheet[i + 1, j + 1];
        var style = grid.Model[i, j];
        var brush = (style.Background as SolidColorBrush);
        //Export with style
        if (brush != null)
            range.CellStyle.Color = brush.Color;
        range.Value = style.CellValue.ToString();
    }
}
StorageFile storageFile;
StorageFolder local = ApplicationData.Current.LocalFolder;
storageFile = await local.CreateFileAsync("Sample.xlsx", CreationCollisionOption.ReplaceExisting);
await workbook.SaveAsAsync(storageFile);
Windows.System.Launcher.LaunchFileAsync(storageFile);
```