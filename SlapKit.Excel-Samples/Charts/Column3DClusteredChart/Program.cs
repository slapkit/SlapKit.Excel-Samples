using Helpers;
using SlapKit.Excel.Excel;
using SlapKit.Excel.Excel.Charts;
using SlapKit.Excel.Excel.Charts.ChartShapeFeatures;

var workbook = new XLWorkbook();
var worksheet = workbook.AddWorksheet();

DataGenerator.GenerateStockData(worksheet);

var chart = worksheet.Charts.AddBar3DChart(XLChartBarDirection.Column);

chart.MoveTo(worksheet.Cell("H2"));
chart.Legend.LegendPosition = XLChartLegendPosition.None;
chart.VaryColors = true;

chart.Series.Add()
    .SetName(worksheet.Range("B1"))
    .SetValues(worksheet.Range("B2:B10"));

workbook.SaveAs("workbook.xlsx");
