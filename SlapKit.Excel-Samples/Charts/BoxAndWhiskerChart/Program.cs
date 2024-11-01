using Helpers;
using SlapKit.Excel.Excel;

var workbook = new XLWorkbook();
var worksheet = workbook.AddWorksheet();

DataGenerator.GenerateSchoolCoursesData(worksheet);

var chart = worksheet.Charts.AddBoxAndWhiskerChart();
chart.MoveTo(worksheet.Cell("F2"), worksheet.Cell("N18"));

var categories = worksheet.Range("A2:A16");

chart.Series.Add()
    .SetName(worksheet.Range("B1"))
    .SetCategories(categories)
    .SetValues(worksheet.Range("B2:B16"));

chart.Series.Add()
    .SetName(worksheet.Range("C1"))
    .SetCategories(categories)
    .SetValues(worksheet.Range("C2:C16"));

chart.Series.Add()
    .SetName(worksheet.Range("D1"))
    .SetCategories(categories)
    .SetValues(worksheet.Range("D2:D16"));

workbook.SaveAs("workbook.xlsx");
