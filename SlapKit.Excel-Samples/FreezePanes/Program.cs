using Helpers;
using SlapKit.Excel.Excel;

var workbook = new XLWorkbook();
var worksheet = workbook.AddWorksheet();

DataGenerator.GenerateStockData(worksheet);

// Method 1: Freeze both rows and columns
worksheet.SheetView.Freeze(1, 1);

// Method 2: Freeze rows or columns via method
// worksheet.SheetView.FreezeRows(1);
// worksheet.SheetView.FreezeColumns(1);

// Method 3: Freeze rows or columns via property
// worksheet.SheetView.SplitRow = 1;
// worksheet.SheetView.SplitColumn = 1;

// Unfreeze rows or columns
// worksheet.SheetView.SplitRow = 0;
// worksheet.SheetView.SplitColumn = 0;

workbook.SaveAs("workbook.xlsx");
