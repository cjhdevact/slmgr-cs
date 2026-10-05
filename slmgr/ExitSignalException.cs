using System;

namespace slmgr
{
    /// <summary>
    /// Control-flow exception used to unwind out of a command handler.
    /// Replaces the process-terminating behaviour of the original
    /// slmgr.vbs so the same command logic can run inside a hosting
    /// process or inside the library build.
    /// </summary>
    public sealed class ExitSignalException : Exception
    {
        public int ExitCode { get; private set; }

        public ExitSignalException(int exitCode)
        {
            ExitCode = exitCode;
        }
    }
}