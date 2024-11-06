using Helpers;
using SlapKit.Excel.Excel;

var workbook = new XLWorkbook();
var worksheet = workbook.AddWorksheet();

DataGenerator.GenerateStockData(worksheet);

var chart = worksheet.Charts.AddContourChart();
chart.Wireframe = true;

chart.MoveTo(worksheet.Cell("H2"));

chart.CreateBandFormat(0).Border.SetColorSolid(XLColor.Red);
chart.CreateBandFormat(1).Border.SetColorSolid(XLColor.Green);
chart.CreateBandFormat(2).Border.SetColorSolid(XLColor.Blue);
chart.CreateBandFormat(3).Border.SetColorSolid(XLColor.Black);

chart.Series.Add()
    .SetValues(worksheet.Range("D2:D10"));

chart.Series.Add()
    .SetValues(worksheet.Range("E2:E10"));

workbook.SaveAs("workbook.xlsx");
