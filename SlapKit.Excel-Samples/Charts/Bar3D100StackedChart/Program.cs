using Helpers;
using SlapKit.Excel.Excel;
using SlapKit.Excel.Excel.Charts.ChartShapeFeatures;

var workbook = new XLWorkbook();
var worksheet = workbook.AddWorksheet();

DataGenerator.GenerateStockData(worksheet);

var chart = worksheet.Charts.AddBar3DChart(XLChartBarDirection.Bar, XLChartBarGrouping.Percent);

chart.MoveTo(worksheet.Cell("H2"));

chart.GapWidth = 80;
chart.GapDepth = 70;
chart.Bar3DShape = XLBar3DShape.Cylinder;

chart.Series.Add()
    .SetCategories(worksheet.Range("B1:F1"))
    .SetValues(worksheet.Range("B2:F2"));

chart.Series.Add()
    .SetBar3DShape(XLBar3DShape.Pyramid)
    .SetCategories(worksheet.Range("B1:F1"))
    .SetValues(worksheet.Range("B3:F3"));

workbook.SaveAs("workbook.xlsx");
