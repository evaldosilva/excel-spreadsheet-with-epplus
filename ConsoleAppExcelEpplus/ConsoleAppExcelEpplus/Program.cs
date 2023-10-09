using Domain;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System.Drawing;

Console.WriteLine("Hello, World!");

// If you use EPPlus in a noncommercial context
// according to the Polyform Noncommercial license:
ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

string dirName = "spreadsheets";
string imageLogoPath = "../../../res/abstract-geometric-logo-or-infinity-line-logo-for-your-company-free-vector.jpg";

List<Employee> emps = new()
{
    new Employee()
    {
        Id = 1,
        Name = "Abel"
    }
};

DirectoryInfo dir = new($".\\{dirName}");
if (!dir.Exists)
    Directory.CreateDirectory($".\\{dirName}");

using (var package = new ExcelPackage($".\\{dirName}\\{Guid.NewGuid()}-myWorkbook.xlsx"))
{
    // Spreadsheet name
    var worksheet = package.Workbook.Worksheets.Add("My Sheet NAme");

    // Spreadsheet image
    var excelImage = worksheet.Drawings.AddPicture("My Logo", imageLogoPath);
    excelImage.SetSize(10);

    //add the image to row 1, column A
    excelImage.SetPosition(0, 0, 0, 0);


    // Spreadsheet Title
    worksheet.Cells["A7"].Value = "Hello World!";

    // Spreadsheet Body
    worksheet.Cells["A8"].LoadFromCollection(emps, true);

    // Apply style on a cell group
    worksheet.Cells["A1:ZC7,C377"].Style.Font.Bold = true;
    using (var range = worksheet.Cells[1, 1, 7, 99]) 
    {
        range.Style.Font.Bold = true;
        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
        range.Style.Fill.BackgroundColor.SetColor(Color.DarkBlue);
        range.Style.Font.Color.SetColor(Color.White);
    }

    using (var range = worksheet.Cells[8, 1, 8, 99])  
    {
        range.Style.Font.Bold = true;
        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
        range.Style.Fill.BackgroundColor.SetColor(Color.BlueViolet);
        range.Style.Font.Color.SetColor(Color.White);
    }

    using (var range = worksheet.Cells[9, 1, 9, 99])  
    {
        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
        range.Style.Fill.BackgroundColor.SetColor(Color.PeachPuff);
        range.Style.Font.Color.SetColor(Color.Purple);
    }

    //make the borders of cells A18 - J18 double and with a purple color
    worksheet.Cells["A9:CU99"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
    worksheet.Cells["A9:CU99"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
    worksheet.Cells["A9:CU99"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
    worksheet.Cells["A9:CU99"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        
    worksheet.Cells["A9:CU99"].Style.Border.Top.Color.SetColor(Color.Black);
    worksheet.Cells["A9:CU99"].Style.Border.Bottom.Color.SetColor(Color.Black);
    worksheet.Cells["A9:CU99"].Style.Border.Left.Color.SetColor(Color.Black);
    worksheet.Cells["A9:CU99"].Style.Border.Right.Color.SetColor(Color.Black);

    // Save to file
    package.Save();
}

// Exporting on ASP.NET 
//public IActionResult GetExcel()
//{
//    using (var package = new ExcelPackage())
//    {
//        var worksheet = package.Workbook.Worksheets.Add("Test");
//        var excelData = package.GetAsByteArray();
//        var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
//        var fileName = "MyWorkbook.xlsx";
//        return File(excelData, contentType, fileName);
//    }
//}