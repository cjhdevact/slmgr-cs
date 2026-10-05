using System;
using System.Management;
using System.Reflection;

namespace slmgr
{
    /// <summary>
    /// Exception that carries a custom HRESULT value. The base Exception class
    /// exposes HResult with a protected setter in .NET Framework 4.5, so a
    /// derived type is required in order to assign it from user code.
    /// </summary>
    internal sealed class HResultException : Exception
    {
        public HResultException(string message, int hresult)
            : base(message)
        {
            HResult = hresult;
        }
    }

    /// <summary>
    /// Token-based activation helpers. Mirrors the TKA commands of slmgr.vbs:
    /// /lil, /ril, /ltc, /fta.
    /// </summary>
    internal static class TkaHelper
    {
        private const uint HR_SL_E_PKEY_NOT_INSTALLED = 0xC004F014;

        /// <summary>
        /// Creates the SPPWMI token activation signer COM object. Equivalent to
        /// the TkaGetSigner() function in slmgr.vbs.
        /// </summary>
        public static object CreateSigner()
        {
            Type t = Type.GetTypeFromProgID("SPPWMI.SppWmiTokenActivationSigner");
            if (t == null)
                throw new InvalidOperationException(
                    "SPPWMI.SppWmiTokenActivationSigner not registered.");
            return Activator.CreateInstance(t);
        }

        /// <summary>
        /// Returns the Windows SoftwareLicensingProduct that has a product key
        /// installed. Equivalent to the TkaGetProduct() function in slmgr.vbs.
        /// </summary>
        public static ManagementObject GetProduct()
        {
            ManagementObject p = WmiHelper.GetProduct(
                "ID, Name, ApplicationId, PartialProductKey, Description, LicenseIsAddon",
                "ApplicationId = '" + WindowsAppIdHolder.WindowsAppId + "' " +
                "AND PartialProductKey <> NULL " +
                "AND LicenseIsAddon = FALSE");

            if (p == null)
                throw new HResultException(
                    "Product not found.",
                    unchecked((int)HR_SL_E_PKEY_NOT_INSTALLED));

            return p;
        }

        /// <summary>
        /// Calls Sign(challenge, thumbprint, pin, out auth2) on the signer and
        /// returns auth1. The fourth argument is an output parameter.
        ///
        /// The COM interface exposes Sign with a by-reference final argument.
        /// We first attempt a reflective lookup that matches the expected
        /// signature; if that fails we fall back to a plain InvokeMember call
        /// so the caller still gets a meaningful result where possible.
        /// </summary>
        public static object Sign(object signer, string challenge,
                                  string thumbprint, string pin, out object auth2)
        {
            auth2 = null;

            object[] args = new object[] { challenge, thumbprint, pin, null };
            Type t = signer.GetType();

            MethodInfo[] methods = t.GetMethods();
            for (int i = 0; i < methods.Length; i++)
            {
                MethodInfo m = methods[i];
                if (m.Name != "Sign") continue;

                ParameterInfo[] ps = m.GetParameters();
                if (ps.Length != 4) continue;

                try
                {
                    object result = m.Invoke(signer, args);
                    auth2 = args[3] != null ? args[3] : new object();
                    return result != null ? result : new object();
                }
                catch
                {
                    // Try the next overload, if any.
                }
            }

            // Fall back to a dynamic-style invocation. The fourth argument may
            // or may not be reported back depending on the COM interop layer.
            object a1 = t.InvokeMember(
                "Sign",
                BindingFlags.InvokeMethod,
                null, signer,
                new object[] { challenge, thumbprint, pin, null });

            auth2 = new object();
            return a1 != null ? a1 : new object();
        }

        /// <summary>Lists installed Token-based Activation licenses.</summary>
        public static void ListILs(Action<string> lineOut)
        {
            lineOut(Resources.Get("L_MsgTkaLicenses"));
            lineOut("");

            int n = 0;

            ManagementObjectSearcher searcher = new ManagementObjectSearcher(
                WmiHelper.Scope,
                new ObjectQuery("SELECT * FROM SoftwareLicensingTokenActivationLicense"));

            foreach (ManagementObject lic in searcher.Get())
            {
                string ilid = WmiHelper.Str(lic, "ILID");
                string ilvid = WmiHelper.Str(lic, "ILVID");

                lineOut(Resources.Get("L_MsgTkaLicenseHeader")
                    .Replace("%ILID%", ilid)
                    .Replace("%ILVID%", ilvid));

                lineOut("    " + Resources.Get("L_MsgTkaLicenseILID")
                    .Replace("%ILID%", ilid));

                lineOut("    " + Resources.Get("L_MsgTkaLicenseILVID")
                    .Replace("%ILVID%", ilvid));

                DateTime? exp = WmiHelper.Date(lic, "ExpirationDate");
                if (exp.HasValue)
                {
                    lineOut("    " + Resources.Get("L_MsgTkaLicenseExpiration")
                        .Replace("%TODATE%", exp.Value.ToString()));
                }

                string more = WmiHelper.Str(lic, "AdditionalInfo");
                if (more != "")
                {
                    lineOut("    " + Resources.Get("L_MsgTkaLicenseAdditionalInfo")
                        .Replace("%MOREINFO%", more));
                }

                uint az = WmiHelper.U32(lic, "AuthorizationStatus");
                if (az != 0)
                {
                    // VBScript Hex() has no leading zeros.
                    lineOut("    " + Resources.Get("L_MsgTkaLicenseAuthZStatus")
                        .Replace("%ERRCODE%", "0x" + az.ToString("X")));
                }
                else
                {
                    lineOut("    " + Resources.Get("L_MsgTkaLicenseDescr")
                        .Replace("%DESC%", WmiHelper.Str(lic, "Description")));
                }

                lineOut("");
                n++;
            }

            if (n == 0) lineOut(Resources.Get("L_MsgTkaLicenseNone"));
        }

        /// <summary>Removes an installed Token-based Activation license.</summary>
        public static void RemoveIL(Action<string> lineOut, string ilid, string ilvid)
        {
            int targetIlvid = int.Parse(ilvid);

            lineOut(Resources.Get("L_MsgTkaRemoving"));
            lineOut("");

            int n = 0;

            ManagementObjectSearcher searcher = new ManagementObjectSearcher(
                WmiHelper.Scope,
                new ObjectQuery("SELECT * FROM SoftwareLicensingTokenActivationLicense"));

            foreach (ManagementObject lic in searcher.Get())
            {
                string licIlid = WmiHelper.Str(lic, "ILID");
                int licIlvid = (int)WmiHelper.U32(lic, "ILVID");

                if (licIlid == ilid && licIlvid == targetIlvid)
                {
                    string msg = Resources.Get("L_MsgTkaRemovedItem")
                        .Replace("%SLID%", WmiHelper.Str(lic, "ID"));

                    WmiHelper.Invoke(lic, "Uninstall");

                    lineOut(msg);
                    n++;
                }
            }

            if (n == 0) lineOut(Resources.Get("L_MsgTkaRemovedNone"));
        }

        /// <summary>Lists Token-based Activation certificates.</summary>
        public static void ListCerts(Action<string> lineOut)
        {
            object signer = CreateSigner();
            ManagementObject product = GetProduct();

            ManagementBaseObject grantsOut = WmiHelper.InvokeWithResult(
                product, "GetTokenActivationGrants",
                new string[0], new object[0]);

            object grants = grantsOut["Grants"];

            object thumbsObj = signer.GetType().InvokeMember(
                "GetCertificateThumbprints",
                BindingFlags.InvokeMethod,
                null, signer,
                new object[] { grants });

            object[] thumbs = thumbsObj as object[];
            if (thumbs != null)
            {
                for (int i = 0; i < thumbs.Length; i++)
                    PrintCert(lineOut, (string)thumbs[i]);
            }
        }

        private static void PrintCert(Action<string> lineOut, string thumbprint)
        {
            string[] a = thumbprint.Split('|');

            lineOut("");
            lineOut(Resources.Get("L_MsgTkaCertThumbprint")
                .Replace("%THUMBPRINT%", a[0]));
            lineOut(Resources.Get("L_MsgTkaCertSubject")
                .Replace("%SUBJECT%", a[1]));
            lineOut(Resources.Get("L_MsgTkaCertIssuer")
                .Replace("%ISSUER%", a[2]));
            lineOut(Resources.Get("L_MsgTkaCertValidFrom")
                .Replace("%FROMDATE%", a[3]));
            lineOut(Resources.Get("L_MsgTkaCertValidTo")
                .Replace("%TODATE%", a[4]));
        }

        /// <summary>Forces a Token-based Activation.</summary>
        public static void Activate(Action<string> lineOut,
                                    Action<ManagementObject> printStatus,
                                    string thumbprint, string pin)
        {
            object signer = CreateSigner();
            ManagementObject product = GetProduct();
            ManagementObject svc = WmiHelper.GetService("Version");

            lineOut(Resources.Get("L_MsgActivating")
                .Replace("%PRODUCTNAME%", WmiHelper.Str(product, "Name"))
                .Replace("%PRODUCTID%", WmiHelper.Str(product, "ID")));

            ManagementBaseObject chOut = WmiHelper.InvokeWithResult(
                product, "GenerateTokenActivationChallenge",
                new string[0], new object[0]);

            string challenge = (string)chOut["Challenge"];

            object auth2;
            object auth1 = Sign(signer, challenge, thumbprint, pin, out auth2);

            WmiHelper.Invoke(product, "DepositTokenActivationResponse",
                new string[] { "Challenge", "AuthInfo1", "AuthInfo2" },
                new object[] { challenge, auth1, auth2 });

            WmiHelper.Invoke(svc, "RefreshLicenseStatus");
            product.Get();
            printStatus(product);
        }
    }
}