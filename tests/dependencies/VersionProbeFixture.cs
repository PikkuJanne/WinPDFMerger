// Development-only controlled version/fault process. This is not PDFtk,
// Ghostscript, a PDF generator, or evidence of native PDF compatibility.
using System;
using System.Diagnostics;
using System.IO;
using System.Threading;

internal static class VersionProbeFixture
{
    private static int Main(string[] arguments)
    {
        string receipt = Environment.GetEnvironmentVariable("WINPDFMERGER_T07_VERSION_RECEIPT");
        if (!String.IsNullOrEmpty(receipt))
        {
            File.WriteAllLines(receipt, arguments);
        }
        string pidReceipt = Environment.GetEnvironmentVariable("WINPDFMERGER_T07_VERSION_PID_RECEIPT");
        if (!String.IsNullOrEmpty(pidReceipt))
        {
            File.WriteAllText(pidReceipt, Process.GetCurrentProcess().Id.ToString());
        }
        string gsOptionsReceipt = Environment.GetEnvironmentVariable("WINPDFMERGER_T07_GS_OPTIONS_RECEIPT");
        if (!String.IsNullOrEmpty(gsOptionsReceipt))
        {
            File.WriteAllText(gsOptionsReceipt, Environment.GetEnvironmentVariable("GS_OPTIONS") ?? "<unset>");
        }

        if (arguments.Length != 1 || arguments[0] != "--version")
        {
            Console.Error.WriteLine("T07 controlled fixture refuses every command except --version.");
            return 64;
        }

        string mode = Environment.GetEnvironmentVariable("WINPDFMERGER_T07_VERSION_MODE");
        if (mode == "nonzero")
        {
            Console.WriteLine("pdftk 2.02 a Handy Tool for Manipulating PDF Documents");
            Console.Error.WriteLine("T07 controlled version failure.");
            return 7;
        }
        if (mode == "unrecognized")
        {
            Console.WriteLine("T07 unrelated executable, version 2.02");
            return 0;
        }
        if (mode == "timeout")
        {
            // Finite even if the application probe regresses to an unbounded wait.
            // The application has a five-second version-probe bound.
            Thread.Sleep(7000);
            Console.WriteLine("pdftk 2.02 a Handy Tool for Manipulating PDF Documents");
            return 0;
        }

        Console.Error.WriteLine("T07 controlled fixture requires an explicit fault mode.");
        return 65;
    }
}
