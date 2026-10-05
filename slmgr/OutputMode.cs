using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace slmgr
{
    /// <summary>
    /// Decides where command output goes:
    ///   - When invoked through Slmgr.Run, output is captured into a string.
    ///   - Otherwise, output goes to Console (CUI) or MessageBox (GUI),
    ///     depending on the PE subsystem of the running executable.
    /// </summary>
    public static class OutputMode
    {
        private enum Subsystem { Unknown = 0, WindowsGui = 2, WindowsCui = 3 }

        private static Subsystem _detected = Subsystem.Unknown;
        private static bool _detectionDone;

        // Redirect target. When set, all output goes here instead of the
        // console or a MessageBox. Used by Slmgr.Run.
        private static Action<string> _redirect;

        /// <summary>Installs a redirect target. Pass null to restore default.</summary>
        public static void SetRedirect(Action<string> redirect)
        {
            _redirect = redirect;
        }

        /// <summary>True when the executing process is a Windows GUI app.</summary>
        public static bool IsWindowsGui
        {
            get
            {
                if (!_detectionDone)
                {
                    _detected = DetectSubsystem();
                    _detectionDone = true;
                }
                return _detected == Subsystem.WindowsGui;
            }
        }

        private static Subsystem DetectSubsystem()
        {
            try
            {
                Assembly entry = Assembly.GetEntryAssembly();
                if (entry == null) return Subsystem.Unknown;

                string path = entry.Location;
                using (FileStream fs = new FileStream(path, FileMode.Open,
                                                      FileAccess.Read, FileShare.Read))
                {
                    byte[] buf = new byte[4];
                    if (fs.Length < 0x40) return Subsystem.Unknown;

                    fs.Seek(0x3C, SeekOrigin.Begin);
                    if (fs.Read(buf, 0, 4) != 4) return Subsystem.Unknown;

                    int peOffset = BitConverter.ToInt32(buf, 0);
                    if (peOffset <= 0) return Subsystem.Unknown;

                    long subsystemOffset = (long)peOffset + 0x5C;
                    if (subsystemOffset + 2 > fs.Length) return Subsystem.Unknown;

                    fs.Seek(subsystemOffset, SeekOrigin.Begin);
                    if (fs.Read(buf, 0, 2) != 2) return Subsystem.Unknown;

                    ushort subsystem = BitConverter.ToUInt16(buf, 0);
                    if (subsystem == 2) return Subsystem.WindowsGui;
                    if (subsystem == 3) return Subsystem.WindowsCui;
                    return Subsystem.Unknown;
                }
            }
            catch
            {
                return Subsystem.Unknown;
            }
        }

        /// <summary>Writes text to the current output target.</summary>
        public static void Write(string text)
        {
            if (_redirect != null)
            {
                _redirect(text);
                return;
            }

            if (IsWindowsGui)
            {
                MessageBox.Show(text, "slmgr", MessageBoxButtons.OK, MessageBoxIcon.None);
            }
            else
            {
                Console.Write(text);
            }
        }
    }
}