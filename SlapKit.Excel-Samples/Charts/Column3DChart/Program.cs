using Helpers;
using SlapKit.Excel.Excel;
using SlapKit.Excel.Excel.Charts.ChartShapeFeatures;

var workbook = new XLWorkbook();
var worksheet = workbook.AddWorksheet();

DataGenerator.GenerateStockData(worksheet);

var chart = worksheet.Charts.AddBar3DChart(XLChartBarDirection.Column, XLChartBarGrouping.Standard);

chart.MoveTo(worksheet.Cell("H2"));

chart.Properties3D.View3D.RotateX = 83;
chart.Properties3D.View3D.RotateY = 41;

chart.SeriesAxis.MinorTickMark = SlapKit.Excel.Excel.Charts.ChartTypeFeatures.XLAxisTickMark.Inside;
chart.SeriesAxis.MajorTickMark = SlapKit.Excel.Excel.Charts.ChartTypeFeatures.XLAxisTickMark.Outside;

chart.SeriesAxis.LineStyle.SetColorSolid(XLColor.Purple).SetWidth(3);

chart.GapWidth = 92;
chart.GapDepth = 101;
chart.Bar3DShape = XLBar3DShape.PyramidPartial;

chart.Series.Add()
    .SetCategories(worksheet.Range("B1:F1"))
    .SetValues(worksheet.Range("B2:F2"))
    .SetBar3DShape(XLBar3DShape.ConePartial);

chart.Series.Add()
    .SetCategories(worksheet.Range("B1:F1"))
    .SetValues(worksheet.Range("B3:F3"));

workbook.SaveAs("workbook.xlsx");
