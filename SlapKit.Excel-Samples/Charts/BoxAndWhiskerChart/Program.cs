using Helpers;
using SlapKit.Excel.Excel;
using SlapKit.Excel.Excel.Charts.ExtendedChartTypeFeatures;
using SlapKit.Excel.Excel.Charts.ExtendedSeriesFeatures;

var workbook = new XLWorkbook();
var worksheet = workbook.AddWorksheet();

DataGenerator.GenerateSchoolCoursesData(worksheet);

var chart = worksheet.Charts.AddBoxAndWhiskerChart();
chart.MoveTo(worksheet.Cell("F2"), worksheet.Cell("N18"));

chart.CategoryAxis.GapWidthRatio = 0.24;
chart.CategoryAxis.MajorTickMark = XLAxisTickMark.Inside;

chart.ValueAxis.MinorTickMark = XLAxisTickMark.Inside;
chart.ValueAxis.MajorTickMark = XLAxisTickMark.Outside;
chart.ValueAxis.Scaling.MinAxisValue = 45;
chart.ValueAxis.SetDisplayUnit(x =>
{
    x.Type = XLAxisDisplayUnitType.Hundreds;
    x.SetLabel();
});

var categories = worksheet.Range("A2:A16");

chart.Series.Add()
    .SetName(worksheet.Range("B1"))
    .SetCategories(categories)
    .SetValues(worksheet.Range("B2:B16"))
    .SetElementVisibilities(x =>
    {
        x.ShowMeanLine = false;
        x.ShowMeanMarkers = false;
        x.ShowInnerPoints = false;
        x.ShowOutlierPoints = false;
    });

chart.Series.Add()
    .SetName(worksheet.Range("C1"))
    .SetCategories(categories)
    .SetValues(worksheet.Range("C2:C16"))
    .SetQuartileMethod(XLSeriesQuartileMethod.Exclusive)
    .SetElementVisibilities(x =>
    {
        x.ShowMeanLine = true;
        x.ShowMeanMarkers = true;
        x.ShowInnerPoints = true;
        x.ShowOutlierPoints = true;
    });

chart.Series.Add()
    .SetName(worksheet.Range("D1"))
    .SetCategories(categories)
    .SetValues(worksheet.Range("D2:D16"))
    .SetQuartileMethod(XLSeriesQuartileMethod.Inclusive)
    .SetElementVisibilities(x =>
    {
        x.ShowMeanLine = false;
        x.ShowMeanMarkers = true;
        x.ShowInnerPoints = false;
        x.ShowOutlierPoints = false;
    });

workbook.SaveAs("workbook.xlsx");
