using System;

public static class ConsoleUI
{
    public static void PrintHeader(string title)
    {
        Console.ForegroundColor = ConsoleColor.Green;
        string line = new string('═', title.Length + 4);
        Console.WriteLine($"╔{line}╗");
        Console.WriteLine($"║  {title}  ║");
        Console.WriteLine($"╚{line}╝");
        Console.ResetColor();
    }

    public static void PrintMenu(params (string key, string label, ConsoleColor color)[] options)
    {
        foreach (var (key, label, color) in options)
        {
            Console.ForegroundColor = color;
            Console.WriteLine($" {key}. {label}");
        }
        Console.ResetColor();
    }

    public static void PrintLine(string text, ConsoleColor color = ConsoleColor.Gray)
    {
        Console.ForegroundColor = color;
        Console.WriteLine(text);
        Console.ResetColor();
    }
}
