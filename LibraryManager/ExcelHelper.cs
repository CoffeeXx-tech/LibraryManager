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

            newId = $"{newRow - 1:000}";

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
    public static void SearchBooks(string path, string searchBy, string keyword)
    {
        if (!File.Exists(path))
        {
            Console.WriteLine("❌ Book file not found.");
            return;
        }

        using (var package = new ExcelPackage(new FileInfo(path)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            int rowCount = worksheet.Dimension?.Rows ?? 0;
            bool found = false;

            int searchColumn = searchBy.ToLower() switch
            {
                "isbn" => 1,
                "title" => 2,
                "author" => 3,
                _ => -1
            };

            if (searchColumn == -1)
            {
                Console.WriteLine("Invalid search type.");
                return;
            }

            Console.WriteLine($"\n🔍 Search results for \"{keyword}\" in {searchBy}:\n");

            for (int row = 2; row <= rowCount; row++)
            {
                string cellValue = worksheet.Cells[row, searchColumn].Text;

                if (cellValue.Contains(keyword, StringComparison.OrdinalIgnoreCase))
                {
                    for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                    {
                        Console.Write(worksheet.Cells[row, col].Text + "\t");
                    }
                    Console.WriteLine();
                    found = true;
                }
            }

            if (!found)
            {
                Console.WriteLine("No matching books found.");
            }
        }
    }
    public static void SearchStudents(string path, string searchBy, string keyword)
    {
        if (!File.Exists(path))
        {
            Console.WriteLine("❌ Student file not found.");
            return;
        }

        using (var package = new ExcelPackage(new FileInfo(path)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            int rowCount = worksheet.Dimension?.Rows ?? 0;
            bool found = false;

            int searchColumn = searchBy.ToLower() switch
            {
                "id" => 1,
                "first name" => 2,
                "last name" => 3,
                "major" => 4,
                "year" => 5,
                "email" => 6,
                _ => -1
            };

            if (searchColumn == -1)
            {
                Console.WriteLine("Invalid search type.");
                return;
            }

            Console.WriteLine($"\n🔍 Search results for \"{keyword}\" in {searchBy}:\n");

            for (int row = 2; row <= rowCount; row++) // Skip header
            {
                string cellValue = worksheet.Cells[row, searchColumn].Text;

                if (cellValue.Contains(keyword, StringComparison.OrdinalIgnoreCase))
                {
                    for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                    {
                        Console.Write(worksheet.Cells[row, col].Text + "\t");
                    }
                    Console.WriteLine();
                    found = true;
                }
            }

            if (!found)
            {
                Console.WriteLine("No matching students found.");
            }
        }
    }
    public static void CreateLoansExcel(string path)
    {
        using var package = new ExcelPackage();
        var worksheet = package.Workbook.Worksheets.Add("Loans");

        worksheet.Cells[1, 1].Value = "Loan ID";
        worksheet.Cells[1, 2].Value = "Student ID";
        worksheet.Cells[1, 3].Value = "Book ISBN";
        worksheet.Cells[1, 4].Value = "Borrow Date";
        worksheet.Cells[1, 5].Value = "Return Date";

        package.SaveAs(new FileInfo(path));
        Console.WriteLine("Loan file created.");
    }

public static bool StudentExists(string path, string studentId)
{
    using var package = new ExcelPackage(new FileInfo(path));
    var worksheet = package.Workbook.Worksheets[0];
    int rowCount = worksheet.Dimension.Rows;

    for (int row = 2; row <= rowCount; row++)
    {
        if (worksheet.Cells[row, 1].Text == studentId)
            return true;
    }
    return false;
}

public static void RegisterStudent(string path, string studentId, string firstName, string lastName, string major, int year, string email)
{
    if (StudentExists(path, studentId))
    {
        Console.WriteLine("❗ A student with this ID already exists.");
        return;
    }

    AddStudent(path, firstName, lastName, major, year, email);
    Console.WriteLine("✅ Registration successful.");
}

public static bool LoginStudent(string path, string studentId, string lastName)
{
    using var package = new ExcelPackage(new FileInfo(path));
    var worksheet = package.Workbook.Worksheets[0];
    int rowCount = worksheet.Dimension.Rows;

    for (int row = 2; row <= rowCount; row++)
    {
        if (worksheet.Cells[row, 1].Text == studentId &&
            worksheet.Cells[row, 3].Text.Equals(lastName, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }
    }
    return false;
}


}
