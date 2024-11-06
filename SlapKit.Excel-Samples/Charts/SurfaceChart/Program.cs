using Helpers;
using SlapKit.Excel.Excel;

var workbook = new XLWorkbook();
var worksheet = workbook.AddWorksheet();

DataGenerator.GenerateStockData(worksheet);

var chart = worksheet.Charts.AddSurfaceChart();
chart.MoveTo(worksheet.Cell("H2"));

var bandFormat = chart.CreateBandFormat(1);
bandFormat.Fill.SetColorSolid(XLColor.Pink);
bandFormat.Border
    .SetColorSolid(XLColor.Black)
    .SetWidth(3)
    .SetDashType(SlapKit.Excel.Excel.DrawingML.XLDrawingLineDashType.Dash);

chart.Series.Add()
    .SetValues(worksheet.Range("B2:B10"));

chart.Series.Add()
    .SetValues(worksheet.Range("D2:D10"));

workbook.SaveAs("workbook.xlsx");
