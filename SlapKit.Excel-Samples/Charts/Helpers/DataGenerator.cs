using SlapKit.Excel.Excel;

namespace Helpers;

public static class DataGenerator
{
    public static void GenerateStockData(IXLWorksheet worksheet)
    {
        worksheet.Cell(2, 1).Value = new DateTime(2020, 1, 1, 0, 0, 0, DateTimeKind.Utc);
        worksheet.Cell(3, 1).Value = new DateTime(2020, 1, 2, 0, 0, 0, DateTimeKind.Utc);
        worksheet.Cell(4, 1).Value = new DateTime(2020, 1, 3, 0, 0, 0, DateTimeKind.Utc);
        worksheet.Cell(5, 1).Value = new DateTime(2020, 1, 4, 0, 0, 0, DateTimeKind.Utc);
        worksheet.Cell(6, 1).Value = new DateTime(2020, 1, 5, 0, 0, 0, DateTimeKind.Utc);
        worksheet.Cell(7, 1).Value = new DateTime(2020, 1, 6, 0, 0, 0, DateTimeKind.Utc);
        worksheet.Cell(8, 1).Value = new DateTime(2020, 1, 7, 0, 0, 0, DateTimeKind.Utc);
        worksheet.Cell(9, 1).Value = new DateTime(2020, 1, 8, 0, 0, 0, DateTimeKind.Utc);
        worksheet.Cell(10, 1).Value = new DateTime(2020, 1, 9, 0, 0, 0, DateTimeKind.Utc);

        worksheet.Cell(1, 2).Value = "Volume";
        worksheet.Cell(2, 2).Value = 2;
        worksheet.Cell(3, 2).Value = 3;
        worksheet.Cell(4, 2).Value = 8;
        worksheet.Cell(5, 2).Value = 5;
        worksheet.Cell(6, 2).Value = 13;
        worksheet.Cell(7, 2).Value = 32;
        worksheet.Cell(8, 2).Value = 25;
        worksheet.Cell(9, 2).Value = 30;
        worksheet.Cell(10, 2).Value = 22;

        worksheet.Cell(1, 3).Value = "Open";
        worksheet.Cell(2, 3).Value = 28.22;
        worksheet.Cell(3, 3).Value = 27.77;
        worksheet.Cell(4, 3).Value = 27.31;
        worksheet.Cell(5, 3).Value = 27.34;
        worksheet.Cell(6, 3).Value = 26.99;
        worksheet.Cell(7, 3).Value = 26.55;
        worksheet.Cell(8, 3).Value = 26.38;
        worksheet.Cell(9, 3).Value = 26.4;
        worksheet.Cell(10, 3).Value = 26.28;

        worksheet.Cell(1, 4).Value = "High";
        worksheet.Cell(2, 4).Value = 28.66;
        worksheet.Cell(3, 4).Value = 28.5;
        worksheet.Cell(4, 4).Value = 27.92;
        worksheet.Cell(5, 4).Value = 27.5;
        worksheet.Cell(6, 4).Value = 27.6;
        worksheet.Cell(7, 4).Value = 26.9;
        worksheet.Cell(8, 4).Value = 26.79;
        worksheet.Cell(9, 4).Value = 26.65;
        worksheet.Cell(10, 4).Value = 26.65;

        worksheet.Cell(1, 5).Value = "Low";
        worksheet.Cell(2, 5).Value = 28.12;
        worksheet.Cell(3, 5).Value = 27.7;
        worksheet.Cell(4, 5).Value = 27.29;
        worksheet.Cell(5, 5).Value = 27.15;
        worksheet.Cell(6, 5).Value = 26.97;
        worksheet.Cell(7, 5).Value = 26.53;
        worksheet.Cell(8, 5).Value = 26.38;
        worksheet.Cell(9, 5).Value = 26.4;
        worksheet.Cell(10, 5).Value = 26.24;

        worksheet.Cell(1, 6).Value = "Close";
        worksheet.Cell(2, 6).Value = 28.35;
        worksheet.Cell(3, 6).Value = 28.35;
        worksheet.Cell(4, 6).Value = 27.77;
        worksheet.Cell(5, 6).Value = 27.32;
        worksheet.Cell(6, 6).Value = 27.41;
        worksheet.Cell(7, 6).Value = 26.9;
        worksheet.Cell(8, 6).Value = 26.77;
        worksheet.Cell(9, 6).Value = 26.47;
        worksheet.Cell(10, 6).Value = 26.6;

        worksheet.Range("A1:A10").Style.Font.Bold = true;
        worksheet.Range("B1:F1").Style.Font.Bold = true;
        worksheet.Range("B1:F1").Style.Fill.BackgroundColor = XLColor.Yellow;
        worksheet.Range("B1:F1").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
    }

    public static void GenerateFruitExpensesData(IXLWorksheet worksheet)
    {
        worksheet.Cell(1, 1).Value = "Fruit";

        worksheet.Cell(1, 2).Value = "Apple";
        worksheet.Cell(1, 3).Value = 168;
        worksheet.Cell(2, 2).Value = "Banana";
        worksheet.Cell(2, 3).Value = 194;
        worksheet.Cell(3, 2).Value = "Orange";
        worksheet.Cell(3, 3).Value = 120;

        worksheet.Cell(4, 1).Value = "Vegetable";

        worksheet.Cell(4, 2).Value = "Carrot";
        worksheet.Cell(4, 3).Value = 90;
        worksheet.Cell(5, 2).Value = "Potato";
        worksheet.Cell(5, 3).Value = 134;
        worksheet.Cell(6, 2).Value = "Onion";
        worksheet.Cell(6, 3).Value = 80;

        worksheet.Cell(7, 1).Value = "Packaged Food";

        worksheet.Cell(7, 2).Value = "Sauce";
        worksheet.Cell(7, 3).Value = 50;
        worksheet.Cell(8, 2).Value = "Noodle";
        worksheet.Cell(8, 3).Value = 78;
        worksheet.Cell(9, 2).Value = "Milk";
        worksheet.Cell(9, 3).Value = 197;

        worksheet.Columns(1, 2).AdjustToContents();
        worksheet.Range("A1:B9").Style.Font.Bold = true;
        worksheet.Range("A1:C3").Style.Fill.BackgroundColor = XLColor.FromTheme(XLThemeColor.Accent1);
        worksheet.Range("A4:C6").Style.Fill.BackgroundColor = XLColor.FromTheme(XLThemeColor.Accent2);
        worksheet.Range("A7:C9").Style.Fill.BackgroundColor = XLColor.FromTheme(XLThemeColor.Accent3);
    }

    public static void GenerateSchoolCoursesData(IXLWorksheet worksheet)
    {
        worksheet.Cell(1, 1).Value = "Course";
        worksheet.Cell(1, 2).Value = "School A";
        worksheet.Cell(1, 3).Value = "School B";
        worksheet.Cell(1, 4).Value = "School C";

        worksheet.Cell(2, 1).Value = "English";
        worksheet.Cell(2, 2).Value = 63;
        worksheet.Cell(2, 3).Value = 53;
        worksheet.Cell(2, 4).Value = 45;

        worksheet.Cell(3, 1).Value = "Physics";
        worksheet.Cell(3, 2).Value = 61;
        worksheet.Cell(3, 3).Value = 55;
        worksheet.Cell(3, 4).Value = 65;

        worksheet.Cell(4, 1).Value = "English";
        worksheet.Cell(4, 2).Value = 63;
        worksheet.Cell(4, 3).Value = 50;
        worksheet.Cell(4, 4).Value = 65;

        worksheet.Cell(5, 1).Value = "Math";
        worksheet.Cell(5, 2).Value = 62;
        worksheet.Cell(5, 3).Value = 51;
        worksheet.Cell(5, 4).Value = 64;

        worksheet.Cell(6, 1).Value = "English";
        worksheet.Cell(6, 2).Value = 46;
        worksheet.Cell(6, 3).Value = 53;
        worksheet.Cell(6, 4).Value = 66;

        worksheet.Cell(7, 1).Value = "English";
        worksheet.Cell(7, 2).Value = 58;
        worksheet.Cell(7, 3).Value = 56;
        worksheet.Cell(7, 4).Value = 67;

        worksheet.Cell(8, 1).Value = "Math";
        worksheet.Cell(8, 2).Value = 60;
        worksheet.Cell(8, 3).Value = 51;
        worksheet.Cell(8, 4).Value = 67;

        worksheet.Cell(9, 1).Value = "Math";
        worksheet.Cell(9, 2).Value = 62;
        worksheet.Cell(9, 3).Value = 53;
        worksheet.Cell(9, 4).Value = 66;

        worksheet.Cell(10, 1).Value = "English";
        worksheet.Cell(10, 2).Value = 63;
        worksheet.Cell(10, 3).Value = 54;
        worksheet.Cell(10, 4).Value = 64;

        worksheet.Cell(11, 1).Value = "English";
        worksheet.Cell(11, 2).Value = 63;
        worksheet.Cell(11, 3).Value = 52;
        worksheet.Cell(11, 4).Value = 67;

        worksheet.Cell(12, 1).Value = "Physics";
        worksheet.Cell(12, 2).Value = 60;
        worksheet.Cell(12, 3).Value = 56;
        worksheet.Cell(12, 4).Value = 64;

        worksheet.Cell(13, 1).Value = "English";
        worksheet.Cell(13, 2).Value = 60;
        worksheet.Cell(13, 3).Value = 56;
        worksheet.Cell(13, 4).Value = 67;

        worksheet.Cell(14, 1).Value = "Math";
        worksheet.Cell(14, 2).Value = 61;
        worksheet.Cell(14, 3).Value = 56;
        worksheet.Cell(14, 4).Value = 45;

        worksheet.Cell(15, 1).Value = "Math";
        worksheet.Cell(15, 2).Value = 63;
        worksheet.Cell(15, 3).Value = 58;
        worksheet.Cell(15, 4).Value = 64;

        worksheet.Cell(16, 1).Value = "English";
        worksheet.Cell(16, 2).Value = 59;
        worksheet.Cell(16, 3).Value = 54;
        worksheet.Cell(16, 4).Value = 65;

        worksheet.Range("A1:A16").Style.Font.Bold = true;
        worksheet.Range("B1:D1").Style.Font.Bold = true;
        worksheet.Range("A1:D1").Style.Fill.BackgroundColor = XLColor.Yellow;
        worksheet.Range("B1:D1").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
    }
}
