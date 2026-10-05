using System;
using System.Text;

namespace slmgr
{
#if !LIBRARY_BUILD
    public static class Program
    {
        public static int Main(string[] args)
        {
            if (!OutputMode.IsWindowsGui)
            {
                try { Console.OutputEncoding = Encoding.UTF8; }
                catch { }
            }

            try
            {
                return Commands.ExecuteCommandLine(args);
            }
            catch (ExitSignalException ex)
            {
                return ex.ExitCode;
            }
            catch (Exception ex)
            {
                if (!string.IsNullOrEmpty(ex.Message))
                    OutputMode.Write(ex.Message);
                return ex.HResult;
            }
        }
    }
#endif
}