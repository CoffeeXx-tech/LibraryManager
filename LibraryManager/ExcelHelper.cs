using System;
using System.IO;
using OfficeOpenXml;

public static class ExcelHelper
{
    public static void CreateBookExcel(string path)
    {
        using (var package = new ExcelPackage())
        {
            var worksheet = package.Workbook.Worksheets.Add("Books");
            worksheet.Cells[1, 1].Value = "ID";
            worksheet.Cells[1, 2].Value = "Title";
            worksheet.Cells[1, 3].Value = "Author";

            package.SaveAs(new FileInfo(path));
            Console.WriteLine("Books file created.");
        }
    }

    public static void CreateStudentExcel(string path)
    {
        using (var package = new ExcelPackage())
        {
            var worksheet = package.Workbook.Worksheets.Add("Students");
            worksheet.Cells[1, 1].Value = "ID";
            worksheet.Cells[1, 2].Value = "Name";
            worksheet.Cells[1, 3].Value = "Email";

            package.SaveAs(new FileInfo(path));
            Console.WriteLine("Students file created.");
        }
    }

    public static void DisplayExcelContent(string path)
    {
        if (!File.Exists(path))
        {
            Console.WriteLine("File does not exist.");
            return;
        }

        using (var package = new ExcelPackage(new FileInfo(path)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            int rowCount = worksheet.Dimension.Rows;
            int colCount = worksheet.Dimension.Columns;

            for (int row = 1; row <= rowCount; row++)
            {
                for (int col = 1; col <= colCount; col++)
                {
                    Console.Write(worksheet.Cells[row, col].Text + "\t");
                }
                Console.WriteLine();
            }
        }

    }
    public static void AddStudent(string path, string firstName, string lastName, string major, int year, string email)
{
    string newId = ""; 

    using (var package = new ExcelPackage(new FileInfo(path)))
    {
        var worksheet = package.Workbook.Worksheets[0];

        int rowCount = worksheet.Dimension?.Rows ?? 1;
        int newRow = rowCount + 1;

        newId = $"S{newRow - 1:000}";

        worksheet.Cells[newRow, 1].Value = newId;
        worksheet.Cells[newRow, 2].Value = firstName;
        worksheet.Cells[newRow, 3].Value = lastName;
        worksheet.Cells[newRow, 4].Value = major;
        worksheet.Cells[newRow, 5].Value = year;
        worksheet.Cells[newRow, 6].Value = email;

        package.Save();
    }

    Console.WriteLine($"✅ Student added with ID {newId}");
}

public static void AddBook(string path, string ibsn, string title, string author, string genere, int publicationYear, int availableCopies)
{
    using var package = new ExcelPackage(new FileInfo(path));
    var ws = package.Workbook.Worksheets[0];
    int lastRow = ws.Dimension?.End.Row ?? 1;

    ws.Cells[lastRow + 1, 1].Value = ibsn;
    ws.Cells[lastRow + 1, 2].Value = title;
    ws.Cells[lastRow + 1, 3].Value = author;
    ws.Cells[lastRow + 1, 4].Value = genere;
    ws.Cells[lastRow + 1, 5].Value = publicationYear;
    ws.Cells[lastRow + 1, 6].Value = availableCopies;

    package.Save();
}
}
