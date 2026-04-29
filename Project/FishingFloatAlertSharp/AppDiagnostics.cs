using System;

namespace FishingFloatAlertSharp
{
    internal static class AppDiagnostics
    {
        public static void LogError(string context, Exception? ex = null)
        {
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] [ERROR] {context}");
            if (ex != null)
            {
                Console.WriteLine(ex);
                if (ex.InnerException != null)
                    Console.WriteLine("Inner: " + ex.InnerException);
            }
        }

        public static void LogWarning(string context, Exception? ex = null)
        {
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] [WARN] {context}");
            if (ex != null)
                Console.WriteLine(ex);
        }
    }
}
