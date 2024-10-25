using Helpers;
using SlapKit.Excel.Excel;
using SlapKit.Excel.Excel.Charts.ChartShapeFeatures;
using SlapKit.Excel.Excel.DrawingML;

var workbook = new XLWorkbook();
var worksheet = workbook.AddWorksheet();

DataGenerator.GenerateStockData(worksheet);

var chart = worksheet.Charts.AddBarChart(XLChartBarDirection.Column, XLChartBarGrouping.Clustered);

chart.MoveTo(worksheet.Cell("H2"));

chart.GapWidth = 56;
chart.Overlap = 46;

chart.CategoryAxis.SetMajorGridlines(majorGridlines =>
{
    majorGridlines.LineStyle
        .SetWidth(1.2)
        .SetColorSolid(XLColor.CoolGrey)
        .SetBeginArrowType(XLDrawingLineArrowType.Diamond)
        .SetBeginArrowLength(XLDrawingLineArrowSize.Small)
        .SetBeginArrowWidth(XLDrawingLineArrowSize.Small)
        .SetEndArrowType(XLDrawingLineArrowType.Open);
});

chart.Series.Add()
    .SetCategories(worksheet.Range("B1:F1"))
    .SetValues(worksheet.Range("B4:F4"));

chart.Series.Add()
    .SetCategories(worksheet.Range("B1:F1"))
    .SetValues(worksheet.Range("B5:F5"));

workbook.SaveAs("workbook.xlsx");
