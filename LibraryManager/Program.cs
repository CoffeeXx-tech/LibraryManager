using System;
using System.IO;
using OfficeOpenXml;

class Program
{
    static string booksPath = "data/books.xlsx";
    static string studentsPath = "data/students.xlsx";

    static void Main(string[] args)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        ConsoleUI.PrintHeader("Library Management System");

        EnsureDataFiles();

        bool exit = false;
        while (!exit)
        {
            Console.WriteLine();
            ConsoleUI.PrintMenu(
                ("1", "📚 View Books", ConsoleColor.Cyan),
                ("2", "👤 View Students", ConsoleColor.Yellow),
                ("3", "❌ Exit", ConsoleColor.Red)
            );

            Console.Write("\nChoose an option: ");
            string input = Console.ReadLine();

            switch (input)
            {
                case "1":
                    ConsoleUI.PrintLine("\n📚 Book List:\n", ConsoleColor.Cyan);
                    ExcelHelper.DisplayExcelContent(booksPath);
                    break;
                case "2":
                    ConsoleUI.PrintLine("\n👤 Student List:\n", ConsoleColor.Yellow);
                    ExcelHelper.DisplayExcelContent(studentsPath);
                    break;
                case "3":
                    exit = true;
                    ConsoleUI.PrintLine("\nExiting program. Goodbye!", ConsoleColor.Green);
                    break;
                    case "4":
    Console.Write("Enter IBSN: ");
    string ibsn = Console.ReadLine() ?? "";
    Console.Write("Enter book title: ");
    string bookTitle = Console.ReadLine() ?? "";
    Console.Write("Enter book author: ");
    string bookAuthor = Console.ReadLine() ?? "";
    Console.Write("Enter book genere: ");
    string bookGenere = Console.ReadLine() ?? "";
    Console.Write("Enter publication year (number): ");
    int pubYear = int.TryParse(Console.ReadLine(), out var py) ? py : 0;
    Console.Write("Enter available copies (number): ");
    int copies = int.TryParse(Console.ReadLine(), out var ac) ? ac : 0;

    ExcelHelper.AddBook(booksPath, ibsn, bookTitle, bookAuthor, bookGenere, pubYear, copies);
    Console.WriteLine("Book added.");
    break;

case "5":
    Console.Write("Enter first name: ");
    string firstName = Console.ReadLine() ?? "";
    Console.Write("Enter last name: ");
    string lastName = Console.ReadLine() ?? "";
    Console.Write("Enter major: ");
    string major = Console.ReadLine() ?? "";
    Console.Write("Enter year (number): ");
    int year = int.TryParse(Console.ReadLine(), out var y) ? y : 0;
    Console.Write("Enter email: ");
    string email = Console.ReadLine() ?? "";

    ExcelHelper.AddStudent(studentsPath, firstName, lastName, major, year, email);
    Console.WriteLine("Student added.");
    break;

                default:
                    ConsoleUI.PrintLine("Invalid option. Please try again.", ConsoleColor.Red);
                    break;
            }
        }
    }

    static void EnsureDataFiles()
    {
        if (!Directory.Exists("data"))
            Directory.CreateDirectory("data");

        if (!File.Exists(booksPath))
        {
            Console.Write("Books file not found. Create a new one? (y/n): ");
            if (Console.ReadLine()?.Trim().ToLower() == "y")
                ExcelHelper.CreateBookExcel(booksPath);
        }

        if (!File.Exists(studentsPath))
        {
            Console.Write("Students file not found. Create a new one? (y/n): ");
            if (Console.ReadLine()?.Trim().ToLower() == "y")
                ExcelHelper.CreateStudentExcel(studentsPath);
        }
    }
}
