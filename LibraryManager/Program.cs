using System;
using System.IO;
using OfficeOpenXml;

class Program
{
    static string booksPath = "data/books.xlsx";
    static string studentsPath = "data/students.xlsx";
    static string loansPath = "data/loans.xlsx";


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
                ("3", "❌ Exit", ConsoleColor.Red),
                ("4", "➕ Add Book", ConsoleColor.Green),
                ("5", "➕ Add Student", ConsoleColor.Green),
                ("6", "🔍 Search Book", ConsoleColor.Blue),
                ("7", "🔎 Search Student", ConsoleColor.Magenta),
                ("8", "📝 Register Student", ConsoleColor.DarkGreen),
                ("9", "🔐 Student Login", ConsoleColor.DarkBlue)

            );

            Console.Write("\nChoose an option: ");
            string input = Console.ReadLine();

            switch (input)
            {
                case "0":
                    ConsoleUI.PrintLine("Exiting program. Goodbye!", ConsoleColor.Green);
                    exit = true;
                    break;
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
                case "6":
                    Console.WriteLine("\nSearch by:");
                    Console.WriteLine("1. ISBN");
                    Console.WriteLine("2. Title");
                    Console.WriteLine("3. Author");
                    Console.Write("Choose option (1-3): ");
                    string choice = Console.ReadLine();

                    string searchBy = choice switch
                    {
                        "1" => "ISBN",
                        "2" => "Title",
                        "3" => "Author",
                        _ => ""
                    };

                    if (string.IsNullOrEmpty(searchBy))
                    {
                        Console.WriteLine("Invalid option.");
                        break;
                    }

                    Console.Write($"Enter {searchBy}: ");
                    string keyword = Console.ReadLine();

                    ExcelHelper.SearchBooks(booksPath, searchBy, keyword);
                    break;
                case "7":
                    Console.WriteLine("\nSearch student by:");
                    Console.WriteLine("1. ID");
                    Console.WriteLine("2. First Name");
                    Console.WriteLine("3. Last Name");
                    Console.WriteLine("4. Major");
                    Console.WriteLine("5. Year");
                    Console.WriteLine("6. Email");
                    Console.Write("Choose option (1-6): ");
                    string studentChoice = Console.ReadLine();

                    string studentSearchBy = studentChoice switch
                    {
                        "1" => "ID",
                        "2" => "First Name",
                        "3" => "Last Name",
                        "4" => "Major",
                        "5" => "Year",
                        "6" => "Email",
                        _ => ""
                    };

                    if (string.IsNullOrEmpty(studentSearchBy))
                    {
                        Console.WriteLine("Invalid option.");
                        break;
                    }

                    Console.Write($"Enter {studentSearchBy}: ");
                    string studentKeyword = Console.ReadLine();

                    ExcelHelper.SearchStudents(studentsPath, studentSearchBy, studentKeyword);
                    break;
                case "8":
                    Console.Write("Student ID: ");
                    string regId = Console.ReadLine();
                    Console.Write("First name: ");
                    string regFirst = Console.ReadLine();
                    Console.Write("Last name: ");
                    string regLast = Console.ReadLine();
                    Console.Write("Major: ");
                    string regMajor = Console.ReadLine();
                    Console.Write("Year: ");
                    int regYear = int.Parse(Console.ReadLine() ?? "1");
                    Console.Write("Email: ");
                    string regEmail = Console.ReadLine();

                    ExcelHelper.RegisterStudent(studentsPath, regId, regFirst, regLast, regMajor, regYear, regEmail);
                    break;

                case "9":
                    Console.Write("Student ID: ");
                    string loginId = Console.ReadLine();
                    Console.Write("Last name: ");
                    string loginLast = Console.ReadLine();

                    if (ExcelHelper.LoginStudent(studentsPath, loginId, loginLast))
                    {
                        Console.WriteLine("✅ Login successful.");
                        ShowStudentLoans(loginId);
                    }
                    else
                    {
                        Console.WriteLine("❌ Invalid ID or last name.");
                    }
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
        if (!File.Exists(loansPath))
        {
            Console.Write("Loan file not found. Create a new one? (y/n): ");
            if (Console.ReadLine()?.ToLower() == "y")
                ExcelHelper.CreateLoansExcel(loansPath);
        }
    }
    static void ShowStudentLoans(string studentId)
{
    Console.WriteLine($"\n📖 Books borrowed by Student ID: {studentId}\n");

    using var package = new ExcelPackage(new FileInfo(loansPath));
    var worksheet = package.Workbook.Worksheets[0];
    int rowCount = worksheet.Dimension?.Rows ?? 0;
    bool any = false;

    for (int row = 2; row <= rowCount; row++)
    {
        if (worksheet.Cells[row, 2].Text == studentId)
        {
            Console.WriteLine($"Book ISBN: {worksheet.Cells[row, 3].Text}");
            Console.WriteLine($"Borrowed: {worksheet.Cells[row, 4].Text}");
            Console.WriteLine($"Returned: {worksheet.Cells[row, 5].Text}");
            Console.WriteLine(new string('-', 30));
            any = true;
        }
    }

    if (!any)
        Console.WriteLine("No borrow records found.");
}
}
