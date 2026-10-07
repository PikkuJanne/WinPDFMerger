// Development-only process fixture. This is not a PDF engine or validator.
// Compatible with the Windows .NET Framework C# compiler (no runtime packages).
using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;

internal static class FakeNative
{
    private const int UsageExitCode = 64;
    private const int MaximumSleepMilliseconds = 300000;
    private const int MaximumFloodLines = 65536;
    private const int MaximumFloodWidth = 1024;
    private const long MaximumFloodCharactersPerStream = 16 * 1024 * 1024;

    private static int Main(string[] args)
    {
        Console.OutputEncoding = new UTF8Encoding(false);
        try
        {
            if (args.Length == 0)
                return Usage("A mode is required: echo, streams, fail, sleep, or flood.");

            switch (args[0])
            {
                case "echo":
                    Console.Out.WriteLine(JsonArguments(args, 1));
                    return 0;
                case "streams":
                    if (args.Length > 3)
                        return Usage("streams [stdout text] [stderr text]");
                    Console.Out.WriteLine(args.Length > 1 ? args[1] : "fake stdout");
                    Console.Error.WriteLine(args.Length > 2 ? args[2] : "fake stderr");
                    return 0;
                case "fail":
                    return Fail(args);
                case "sleep":
                    if (args.Length != 2)
                        return Usage("sleep <milliseconds 0..300000>");
                    Thread.Sleep(ParseBoundedInteger(args[1], 0, MaximumSleepMilliseconds));
                    return 0;
                case "flood":
                    return Flood(args);
                default:
                    return Usage("Unknown mode. Use echo, streams, fail, sleep, or flood.");
            }
        }
        catch (Exception error)
        {
            // Expected invalid input/IO failures are visible and never successful.
            Console.Error.WriteLine("fake-native: " + error.GetType().Name + ": " + error.Message);
            return UsageExitCode;
        }
    }

    private static int Fail(string[] args)
    {
        if (args.Length < 2 || args.Length > 3)
            return Usage("fail <exit code 1..255> [absolute new partial-file path]");
        int exitCode = ParseBoundedInteger(args[1], 1, 255);
        if (args.Length == 3)
        {
            // The harness must supply a run-owned path. Never create directories or
            // overwrite an existing file; bytes deliberately do not form a PDF.
            if (!Path.IsPathRooted(args[2]) || Path.GetFullPath(args[2]) != args[2])
                return Usage("Partial-file path must be absolute and normalized.");
            byte[] marker = Encoding.UTF8.GetBytes("FAKE-NATIVE-PARTIAL: not a PDF\r\n");
            using (FileStream output = new FileStream(args[2], FileMode.CreateNew, FileAccess.Write, FileShare.None))
                output.Write(marker, 0, marker.Length);
        }
        Console.Out.WriteLine("fake-native: controlled failure");
        Console.Error.WriteLine("fake-native: requested exit " + exitCode.ToString(CultureInfo.InvariantCulture));
        return exitCode;
    }

    private static int Flood(string[] args)
    {
        if (args.Length < 2 || args.Length > 3)
            return Usage("flood <lines 0..65536> [padding width 1..1024; default 128]");
        int lines = ParseBoundedInteger(args[1], 0, MaximumFloodLines);
        int width = args.Length == 3 ? ParseBoundedInteger(args[2], 1, MaximumFloodWidth) : 128;
        // Include prefixes and CRLF in the limit. Enough output to exercise pipe
        // draining, but invalid combinations fail before writing any flood data.
        if ((long)lines * (width + 16) > MaximumFloodCharactersPerStream)
            return Usage("Requested flood exceeds 16 MiB per stream.");
        string padding = new string('x', width);
        for (int i = 0; i < lines; i++)
        {
            string index = i.ToString("D6", CultureInfo.InvariantCulture);
            Console.Out.WriteLine("stdout:" + index + ":" + padding);
            Console.Error.WriteLine("stderr:" + index + ":" + padding);
        }
        return 0;
    }

    private static int ParseBoundedInteger(string value, int minimum, int maximum)
    {
        int parsed;
        if (!Int32.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out parsed)
            || parsed < minimum || parsed > maximum)
            throw new ArgumentException("Expected integer from " + minimum.ToString(CultureInfo.InvariantCulture)
                + " to " + maximum.ToString(CultureInfo.InvariantCulture) + ".");
        return parsed;
    }

    private static int Usage(string message)
    {
        Console.Error.WriteLine("fake-native: " + message);
        return UsageExitCode;
    }

    private static string JsonArguments(string[] args, int start)
    {
        StringBuilder json = new StringBuilder("[");
        for (int i = start; i < args.Length; i++)
        {
            if (i > start)
                json.Append(',');
            json.Append('"');
            foreach (char value in args[i])
            {
                switch (value)
                {
                    case '"': json.Append("\\\""); break;
                    case '\\': json.Append("\\\\"); break;
                    case '\b': json.Append("\\b"); break;
                    case '\f': json.Append("\\f"); break;
                    case '\n': json.Append("\\n"); break;
                    case '\r': json.Append("\\r"); break;
                    case '\t': json.Append("\\t"); break;
                    default:
                        if (value < 0x20)
                            json.Append("\\u" + ((int)value).ToString("x4", CultureInfo.InvariantCulture));
                        else
                            json.Append(value);
                        break;
                }
            }
            json.Append('"');
        }
        return json.Append(']').ToString();
    }
}
