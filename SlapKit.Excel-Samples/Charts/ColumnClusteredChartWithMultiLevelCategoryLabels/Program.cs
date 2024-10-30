using Helpers;
using SlapKit.Excel.Excel;
using SlapKit.Excel.Excel.Charts;
using SlapKit.Excel.Excel.Charts.ChartShapeFeatures;

var workbook = new XLWorkbook();
var worksheet = workbook.AddWorksheet();

DataGenerator.GenerateFruitExpensesData(worksheet);

var chart = worksheet.Charts.AddBarChart(XLChartBarDirection.Column);

chart.MoveTo(worksheet.Cell("E2"));

chart.Legend.LegendPosition = XLChartLegendPosition.None;
chart.CategoryAxis.LineStyle.SetColorAuto();
chart.CategoryAxis.MultiLevelLabel = true;
chart.CategoryAxis.LabelOffset = 20;
chart.CategoryAxis.TickMarkSkip = 1;

chart.Series.Add()
    .SetName("Monthly expense")
    .SetCategories(worksheet.Range("A1:B9"))
    .SetValues(worksheet.Range("C1:C9"));

workbook.SaveAs("workbook.xlsx");
