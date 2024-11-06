using Helpers;
using SlapKit.Excel.Excel;

var workbook = new XLWorkbook();
var worksheet = workbook.AddWorksheet();

DataGenerator.GenerateStockData(worksheet);

var chart = worksheet.Charts.AddContourChart();
chart.MoveTo(worksheet.Cell("H2"));

chart.Series.Add()
    .SetValues(worksheet.Range("B2:B10"));

chart.Series.Add()
    .SetValues(worksheet.Range("C2:C10"));

workbook.SaveAs("workbook.xlsx");
