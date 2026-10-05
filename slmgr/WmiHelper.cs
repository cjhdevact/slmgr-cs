using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;

namespace slmgr
{
    /// <summary>
    /// Thin wrapper around the WMI classes SoftwareLicensingService and
    /// SoftwareLicensingProduct. Provides helpers to read properties and to
    /// invoke WMI methods by explicit parameter names (avoiding any reliance
    /// on property enumeration order).
    /// </summary>
    public static class WmiHelper
    {
        private static ManagementScope _scope;
        private static bool _isRemote;
        private static string _computerName = ".";

        public static ManagementScope Scope { get { return _scope; } }
        public static bool IsRemote { get { return _isRemote; } }
        public static string ComputerName { get { return _computerName; } }

        /// <summary>
        /// Connects to the local or remote WMI namespace root\cimv2.
        /// </summary>
        public static void Connect(string computer, string user, string password)
        {
            _computerName = string.IsNullOrEmpty(computer) ? "." : computer;

            ConnectionOptions opts = new ConnectionOptions();
            opts.Impersonation = ImpersonationLevel.Impersonate;
            opts.Authentication = AuthenticationLevel.PacketPrivacy;

            if (!string.IsNullOrEmpty(user))
            {
                opts.Username = user;
                opts.Password = password;
            }

            _scope = new ManagementScope(
                string.Format(@"\\{0}\root\cimv2", _computerName), opts);
            _scope.Connect();

            _isRemote = _computerName != ".";
        }

        /// <summary>Returns the single SoftwareLicensingService instance.</summary>
        public static ManagementObject GetService(string select)
        {
            if (string.IsNullOrEmpty(select)) select = "*";

            using (ManagementObjectSearcher s = new ManagementObjectSearcher(
                _scope,
                new ObjectQuery("SELECT " + select + " FROM SoftwareLicensingService")))
            {
                foreach (ManagementObject o in s.Get()) return o;
            }
            return null;
        }

        /// <summary>Returns a collection of SoftwareLicensingProduct instances.</summary>
        public static ManagementObjectCollection GetProducts(string select, string where)
        {
            string q = string.IsNullOrEmpty(where)
                ? "SELECT " + select + " FROM SoftwareLicensingProduct"
                : "SELECT " + select + " FROM SoftwareLicensingProduct WHERE " + where;

            ManagementObjectSearcher s = new ManagementObjectSearcher(_scope, new ObjectQuery(q));
            return s.Get();
        }

        /// <summary>
        /// Returns exactly one SoftwareLicensingProduct instance, or null if
        /// none matched. Throws if more than one matched.
        /// </summary>
        public static ManagementObject GetProduct(string select, string where)
        {
            List<ManagementObject> list = GetProducts(select, where).Cast<ManagementObject>().ToList();
            if (list.Count == 0) return null;
            if (list.Count > 1) throw new InvalidOperationException("Multiple products matched.");
            return list[0];
        }

        /// <summary>Reads a string property; returns "" when null.</summary>
        public static string Str(ManagementObject o, string prop)
        {
            object v = o[prop];
            return v == null ? "" : v.ToString();
        }

        /// <summary>Reads a UInt32 property; returns 0 when null.</summary>
        public static uint U32(ManagementObject o, string prop)
        {
            object v = o[prop];
            return v == null ? 0u : Convert.ToUInt32(v);
        }

        /// <summary>Reads a boolean property; returns false when null.</summary>
        public static bool B(ManagementObject o, string prop)
        {
            object v = o[prop];
            return v is bool && (bool)v;
        }

        /// <summary>
        /// Reads a WMI datetime property and converts it to DateTime.
        /// Returns null if the property is missing or cannot be parsed.
        /// </summary>
        public static DateTime? Date(ManagementObject o, string prop)
        {
            object v = o[prop];
            if (v == null) return null;

            string s = v.ToString();
            if (string.IsNullOrEmpty(s)) return null;

            try { return ManagementDateTimeConverter.ToDateTime(s); }
            catch { return null; }
        }

        /// <summary>
        /// Invokes a WMI method with positional arguments. The caller supplies
        /// the parameter names in order so the code does not depend on the
        /// enumeration order of the WMI class properties.
        /// </summary>
        public static void Invoke(ManagementObject o, string method,
                                  string[] paramNames, object[] args)
        {
            ManagementBaseObject inParams = o.GetMethodParameters(method);
            for (int i = 0; i < args.Length && i < paramNames.Length; i++)
                inParams[paramNames[i]] = args[i];
            o.InvokeMethod(method, inParams, null);
        }

        /// <summary>Invokes a WMI method with no parameters.</summary>
        public static void Invoke(ManagementObject o, string method)
        {
            o.InvokeMethod(method, null);
        }

        /// <summary>
        /// Invokes a WMI method and returns the output object. Used for methods
        /// that produce a return value such as
        /// GenerateActiveDirectoryOfflineActivationId.
        /// </summary>
        public static ManagementBaseObject InvokeWithResult(
            ManagementObject o, string method, string[] paramNames, object[] args)
        {
            ManagementBaseObject inParams = o.GetMethodParameters(method);
            for (int i = 0; i < args.Length && i < paramNames.Length; i++)
                inParams[paramNames[i]] = args[i];
            return o.InvokeMethod(method, inParams, null);
        }
    }
}