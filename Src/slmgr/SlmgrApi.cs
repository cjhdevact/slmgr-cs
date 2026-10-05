using System;
using System.Collections.Generic;
using System.Text;

namespace slmgr
{
    /// <summary>
    /// Public entry point for external callers. Executes an slmgr command
    /// and returns its textual output instead of writing to the console or
    /// showing a MessageBox.
    ///
    /// Usage:
    ///     slmgr.SlmgrResult r = slmgr.Slmgr.Run("/dlv");
    ///     Console.WriteLine(r.Output);
    /// </summary>
    public static class Slmgr
    {
        private static readonly object _gate = new object();

        /// <summary>
        /// Executes the given command line string and returns the captured
        /// output. Option names start with '/' or '-'. Up to three leading
        /// arguments are treated as remote machine name, user and password.
        /// </summary>
        public static SlmgrResult Run(string commandLine)
        {
            return Run(SplitCommandLine(commandLine));
        }

        /// <summary>Executes the given argument array and returns the output.</summary>
        public static SlmgrResult Run(string[] args)
        {
            lock (_gate)
            {
                StringBuilder buffer = new StringBuilder();

                OutputMode.SetRedirect(delegate (string s) { buffer.Append(s); });

                int exitCode = 0;
                try
                {
                    exitCode = Commands.ExecuteCommandLine(args);
                }
                catch (ExitSignalException ex)
                {
                    exitCode = ex.ExitCode;
                }
                catch (Exception ex)
                {
                    if (!string.IsNullOrEmpty(ex.Message))
                        buffer.Append(ex.Message);
                    exitCode = ex.HResult;
                }
                finally
                {
                    OutputMode.SetRedirect(null);
                }

                return new SlmgrResult
                {
                    Output = buffer.ToString(),
                    ExitCode = exitCode
                };
            }
        }

        /// <summary>
        /// Splits a command line string into an argument array. Handles
        /// double-quoted arguments so paths with spaces work correctly.
        /// </summary>
        private static string[] SplitCommandLine(string commandLine)
        {
            if (string.IsNullOrEmpty(commandLine)) return new string[0];

            List<string> parts = new List<string>();
            StringBuilder current = new StringBuilder();
            bool inQuotes = false;

            for (int i = 0; i < commandLine.Length; i++)
            {
                char c = commandLine[i];

                if (c == '"') { inQuotes = !inQuotes; continue; }

                if (!inQuotes && (c == ' ' || c == '\t'))
                {
                    if (current.Length > 0)
                    {
                        parts.Add(current.ToString());
                        current.Length = 0;
                    }
                    continue;
                }

                current.Append(c);
            }

            if (current.Length > 0) parts.Add(current.ToString());
            return parts.ToArray();
        }
    }
}