using System;
using System.Globalization;
using System.Reflection;
using System.Resources;

namespace slmgr
{
    /// <summary>
    /// Localized string lookup. English strings live in Strings.en-US.resx and
    /// serve as the default fallback for any culture that has no translation.
    /// Missing keys raise an exception so typos in code surface immediately.
    /// </summary>
    public static class Resources
    {
        // Logical name of the compiled resource. Must match the namespace,
        // the "Resources" folder and the base file name (without culture
        // suffix). Change "slmgr" here if your namespace differs.
        private const string ResourceBaseName = "slmgr.Resources.Strings";

        private static readonly ResourceManager _rm =
            new ResourceManager(ResourceBaseName, Assembly.GetExecutingAssembly());

        // The default culture used when no translation matches. English is
        // stored under en-US in this project.
        private static readonly CultureInfo DefaultCulture =
            new CultureInfo("en-US");

        /// <summary>Returns the localized string for the given key.</summary>
        public static string Get(string key)
        {
            // 1. Current UI culture.
            string v = _rm.GetString(key, CultureInfo.CurrentUICulture);
            if (v != null) return v;

            // 2. Explicit fallback to en-US.
            if (!CultureInfo.CurrentUICulture.Equals(DefaultCulture))
            {
                v = _rm.GetString(key, DefaultCulture);
                if (v != null) return v;
            }

            // 3. Not found anywhere: this indicates a key that does not exist
            //    in Strings.en-US.resx, i.e. a bug in the code.
            throw new MissingManifestResourceException(
                "Missing resource key: " + key);
        }

        /// <summary>
        /// Returns the error message for the given HRESULT. The key is built
        /// dynamically as "L_MsgError_" + uppercase hex without leading
        /// zeros, matching the VBScript Hex() output. Returns an empty
        /// string when no such key exists, so callers can decide what to do.
        /// </summary>
        public static string GetError(uint hr)
        {
            string key = "L_MsgError_" + hr.ToString("X");

            string v = _rm.GetString(key, CultureInfo.CurrentUICulture);
            if (v != null) return v;

            if (!CultureInfo.CurrentUICulture.Equals(DefaultCulture))
            {
                v = _rm.GetString(key, DefaultCulture);
                if (v != null) return v;
            }

            return "";
        }
    }
}