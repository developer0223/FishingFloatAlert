using System.Runtime.InteropServices;
using System.Text;
using System.Windows;

namespace FishingFloatAlertSharp
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
#if DEBUG
            if (AllocConsole())
            {
                try
                {
                    Console.OutputEncoding = Encoding.UTF8;
                }
                catch
                {
                    // ignore
                }
            }
#endif

            base.OnStartup(e);
        }

#if DEBUG
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool AllocConsole();
#endif
    }
}
