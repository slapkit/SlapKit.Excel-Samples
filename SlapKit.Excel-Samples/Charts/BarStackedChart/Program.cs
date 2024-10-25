using Helpers;
using SlapKit.Excel.Excel;
using SlapKit.Excel.Excel.Charts.ChartShapeFeatures;
using SlapKit.Excel.Excel.DrawingML;

var workbook = new XLWorkbook();

var dataWorksheet = workbook.AddWorksheet();
DataGenerator.GenerateStockData(dataWorksheet);

var chartWorksheet = workbook.AddWorksheet();

var chart = chartWorksheet.Charts.AddBarChart(XLChartBarDirection.Bar, XLChartBarGrouping.Stacked);

chart.MoveTo(chartWorksheet.Cell("B2"));

chart.GapWidth = 66;

chart.CategoryAxis.SetMinorGridlines(minorGridlines =>
{
    minorGridlines.LineStyle
        .SetColorSolid(XLColor.FromTheme(XLThemeColor.Accent3))
        .SetWidth(4)
        .SetBeginArrowType(XLDrawingLineArrowType.Oval)
        .SetEndArrowType(XLDrawingLineArrowType.Stealth)
        .SetEndArrowLength(XLDrawingLineArrowSize.Large)
        .SetEndArrowWidth(XLDrawingLineArrowSize.Medium);
});

chart.Series.Add()
    .SetValues(dataWorksheet.Range("B4:F4"));

chart.Series.Add()
    .SetValues(dataWorksheet.Range("B10:F10"));

workbook.SaveAs("workbook.xlsx");
