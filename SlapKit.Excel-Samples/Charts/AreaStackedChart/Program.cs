using Helpers;
using SlapKit.Excel.Excel;
using SlapKit.Excel.Excel.Charts.ChartShapeFeatures;
using SlapKit.Excel.Excel.Charts.SeriesFeatures;

var workbook = new XLWorkbook();
var worksheet = workbook.AddWorksheet();

DataGenerator.GenerateStockData(worksheet);

var chart = worksheet.Charts
    .AddAreaChart(XLChartGrouping.Stacked)
    .MoveTo(worksheet.Cell("H2"))
    .WithSize(400, 250);

var series = chart.Series.Add()
    .SetValues(worksheet.Range("B2:F2"))
    .AddTrendline(x => x.SetName("toDelete"))
    .AddTrendline(x => x.SetName("default trendline"))
    .AddTrendline(x => x.SetType(XLSeriesTrendlineType.Linear).SetName("toDelete"))
    .AddTrendline(x => x
        .SetName("moving avg[3]")
        .SetType(XLSeriesTrendlineType.MovingAverage)
        .SetPeriod(3)
        .SetLineStyle(l => l
            .SetColorSolid(XLColor.Orange)
            .SetWidth(2)));

chart.Series.Add()
    .SetValues(worksheet.Range("B8:F8"))
    .AddTrendline();

series.RemoveTrendlines(x => x.Name == "toDelete");

workbook.SaveAs("workbook.xlsx");
