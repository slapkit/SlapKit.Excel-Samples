using Helpers;
using SlapKit.Excel.Excel;

var workbook = new XLWorkbook();
var worksheet = workbook.AddWorksheet();

DataGenerator.GenerateStockData(worksheet);

var chart = worksheet.Charts.AddSurfaceChart();
chart.Wireframe = true;

chart.MoveTo(worksheet.Cell("H2"));

chart.ValueAxis.MajorUnit = 25;
chart.ValueAxis.Scaling.MinAxisValue = 0.5;
chart.Properties3D.View3D.RotateX = 25;
chart.Properties3D.View3D.RotateY = 5;

chart.Series.Add()
    .SetValues(worksheet.Range("B2:B10"));

chart.Series.Add()
    .SetValues(worksheet.Range("D2:D10"));

workbook.SaveAs("workbook.xlsx");
