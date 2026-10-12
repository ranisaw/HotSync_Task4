namespace HotsyncApp.Helpers;

/// <summary>
/// Lightweight structured console logger.
/// Prefixes every line with a UTC timestamp and severity level.
/// No external logging framework required — keeps the app dependency-lean.
/// </summary>
public static class Logger
{
    private static readonly object _lock = new();

    public static void Info(string message) => Write("INFO ", ConsoleColor.Cyan, message);
    public static void Success(string message) => Write("OK   ", ConsoleColor.Green, message);
    public static void Warn(string message) => Write("WARN ", ConsoleColor.Yellow, message);
    public static void Error(string message, Exception? ex = null)
    {
        Write("ERROR", ConsoleColor.Red, message);
        if (ex is not null)
            Write("     ", ConsoleColor.DarkRed, $"  {ex.GetType().Name}: {ex.Message}");
    }

    public static void Section(string title)
    {
        Console.WriteLine();
        Write("=====", ConsoleColor.Magenta, $"=== {title} ===");
    }

    private static void Write(string level, ConsoleColor colour, string message)
    {
        lock (_lock)
        {
            var timestamp = DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss");
            Console.Write($"[{timestamp}] ");

            var prev = Console.ForegroundColor;
            Console.ForegroundColor = colour;
            Console.Write($"[{level}]");
            Console.ForegroundColor = prev;

            Console.WriteLine($" {message}");
        }
    }
}
