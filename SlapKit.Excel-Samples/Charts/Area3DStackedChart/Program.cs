using SlapKit.Excel.Excel;
using SlapKit.Excel.Excel.Charts.ChartShapeFeatures;
using SlapKit.Excel.Excel.Charts.ChartTypeFeatures;
using Helpers;

var workbook = new XLWorkbook();
var worksheet = workbook.AddWorksheet();

DataGenerator.GenerateStockData(worksheet);

var chart = worksheet.Charts.AddArea3DChart(XLChartGrouping.Stacked)
    .MoveTo(worksheet.Cell("H2"), worksheet.Cell("O16"));

chart.GapDepth = 97;

chart.ValueAxis.Scaling.MinAxisValue = 2;
chart.ValueAxis.Scaling.MaxAxisValue = 50.1;
chart.ValueAxis.MajorUnit = 15;
chart.ValueAxis.MinorUnit = 1;
chart.ValueAxis.TickLabelPosition = XLAxisTickLabelPosition.High;
chart.ValueAxis.NumberingFormat.SourceLinked = false;
chart.ValueAxis.NumberingFormat.Format = "#.##0.00";

chart.CategoryAxis.MinorTickMark = XLAxisTickMark.Inside;
chart.CategoryAxis.MajorTickMark = XLAxisTickMark.Outside;
chart.CategoryAxis.Scaling.MinAxisValue = worksheet.Cell(3, 1).Value.GetUnifiedNumber();
chart.CategoryAxis.Scaling.MaxAxisValue = worksheet.Cell(9, 1).Value.GetUnifiedNumber();
chart.CategoryAxis.LineStyle.SetColorSolid(XLColor.Red).SetWidth(2);

chart.Series.Add()
    .SetName(worksheet.Range("B1"))
    .SetCategories(worksheet.Range("A2:A10"))
    .SetValues(worksheet.Range("B2:B10"));

chart.Series.Add()
    .SetName(worksheet.Range("C1"))
    .SetValues(worksheet.Range("C2:C10"));

workbook.SaveAs("workbook.xlsx");
