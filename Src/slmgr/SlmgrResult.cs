namespace slmgr
{
    /// <summary>Result of an slmgr command executed through Slmgr.Run.</summary>
    public class SlmgrResult
    {
        /// <summary>Text output that would have been displayed.</summary>
        public string Output { get; set; }

        /// <summary>Exit code that the command line tool would return.</summary>
        public int ExitCode { get; set; }

        /// <summary>True when the command completed without an error code.</summary>
        public bool Success
        {
            get { return ExitCode == 0; }
        }
    }
}