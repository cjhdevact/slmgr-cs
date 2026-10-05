using System;
using System.Globalization;
using System.IO;
using System.Management;
using System.Text;
using Microsoft.Win32;

namespace slmgr
{
    /// <summary>
    /// Implements all slmgr.vbs commands. Option names, output strings and
    /// behaviour match the VBScript source exactly.
    /// </summary>
    public static class Commands
    {
        // -----------------------------------------------------------------
        // Global state
        // -----------------------------------------------------------------

        private static readonly StringBuilder _buf = new StringBuilder();
        private static string _computer = ".";
        private static bool _isRemote;

        /// <summary>Appends a line to the output buffer.</summary>
        private static void LineOut(string s)
        {
            _buf.Append(s).Append('\n');
        }

        private static void LineFlush(string s)
        {
            string combined = _buf.ToString() + s;
            _buf.Clear();
            if (combined.Length > 0)
                OutputMode.Write(combined);
        }

        private static void Exit(int code)
        {
            if (_buf.Length > 0)
            {
                OutputMode.Write(_buf.ToString());
                _buf.Clear();
            }
            throw new ExitSignalException(code);
        }
        // -----------------------------------------------------------------
        // Command-line parsing (mirrors ExecCommandLine in slmgr.vbs)
        // -----------------------------------------------------------------

        public static int ExecuteCommandLine(string[] args)
        {
            string[] remoteInfo = new string[] { ".", "", "", "" };
            bool known = false;
            int indexOption = 0;

            for (indexOption = 0; indexOption <= 3; indexOption++)
            {
                if (indexOption >= args.Length) break;

                string s = args[indexOption];
                if (s.Length == 0) break;

                char ch = s[0];
                if (ch == '/' || ch == '-') { known = true; break; }

                remoteInfo[indexOption] = s;
            }

            if (!known || indexOption == 2)
            {
                _computer = ".";
                known = false;
            }
            else
            {
                _computer = remoteInfo[0];
            }

            try
            {
                WmiHelper.Connect(
                    _computer,
                    remoteInfo.Length > 1 ? remoteInfo[1] : "",
                    remoteInfo.Length > 2 ? remoteInfo[2] : "");
            }
            catch (Exception ex)
            {
                ShowError(Resources.Get("L_MsgErrorLocalWMI"),
                    unchecked((uint)ex.HResult), ex.Message);
                Exit(unchecked((int)ex.HResult));
            }

            _isRemote = WmiHelper.IsRemote;

            if (!known)
            {
                LineOut(Resources.Get("L_MsgInvalidOptions"));
                LineOut("");
                DisplayUsage();
            }

            if (indexOption >= args.Length)
            {
                LineOut(Resources.Get("L_MsgInvalidOptions"));
                LineOut("");
                DisplayUsage();
            }

            if (!ParseCommandLine(args, indexOption))
            {
                LineOut(Resources.Get("L_MsgUnrecognizedOption") + args[indexOption]);
                LineOut("");
                DisplayUsage();
            }

            Exit(0);
            return 0;
        }

        private static bool ParseCommandLine(string[] args, int index)
        {
            string opt = args[index].Substring(1).ToLowerInvariant();

            switch (opt)
            {
                // ---- Global options ----
                case "ipk":
                    if (!Handle(args, index + 1, true, "/ipk", "<Product Key>")) return true;
                    InstallProductKey(args[index + 1]);
                    return true;

                case "upk":
                    UninstallProductKey(Handle(args, index + 1, false, "/upk", "")
                        ? args[index + 1] : "");
                    return true;

                case "ato":
                    ActivateProduct(Handle(args, index + 1, false, "/ato", "")
                        ? args[index + 1] : "");
                    return true;

                case "dli":
                    DisplayAllInformation(Handle(args, index + 1, false, "/dli", "")
                        ? args[index + 1] : "", false);
                    return true;

                case "dlv":
                    DisplayAllInformation(Handle(args, index + 1, false, "/dlv", "")
                        ? args[index + 1] : "", true);
                    return true;

                case "xpr":
                    ExpirationDatetime(Handle(args, index + 1, false, "/xpr", "")
                        ? args[index + 1] : "");
                    return true;

                // ---- Advanced options ----
                case "cpky":
                    ClearPKeyFromRegistry();
                    return true;

                case "ilc":
                    if (!Handle(args, index + 1, true, "/ilc", "<License file>")) return true;
                    InstallLicense(args[index + 1]);
                    return true;

                case "rilc":
                    ReinstallLicenses();
                    return true;

                case "rearm":
                    ReArmWindows();
                    return true;

                case "rearm-app":
                    if (!Handle(args, index + 1, true, "/rearm-app", "<Application ID>")) return true;
                    ReArmApp(args[index + 1]);
                    return true;

                case "rearm-sku":
                    if (!Handle(args, index + 1, true, "/rearm-sku", "<Activation ID>")) return true;
                    ReArmSku(args[index + 1]);
                    return true;

                case "dti":
                    DisplayIID(Handle(args, index + 1, false, "/dti", "")
                        ? args[index + 1] : "");
                    return true;

                case "atp":
                    if (!Handle(args, index + 1, true, "/atp", "<Confirmation ID>")) return true;
                    PhoneActivateProduct(args[index + 1],
                        Handle(args, index + 2, false, "/atp", "")
                            ? args[index + 2] : "");
                    return true;

                // ---- KMS client options ----
                case "skms":
                    if (!Handle(args, index + 1, true, "/skms", "<Name[:Port] | : port>")) return true;
                    SetKmsMachineName(args[index + 1],
                        Handle(args, index + 2, false, "/skms", "")
                            ? args[index + 2] : "");
                    return true;

                case "ckms":
                    ClearKms(Handle(args, index + 1, false, "/ckms", "")
                        ? args[index + 1] : "");
                    return true;

                case "skms-domain":
                    if (!Handle(args, index + 1, true, "/skms-domain", "<FQDN>")) return true;
                    SetKmsLookupDomain(args[index + 1],
                        Handle(args, index + 2, false, "/skms-domain", "")
                            ? args[index + 2] : "");
                    return true;

                case "ckms-domain":
                    ClearKmsLookupDomain(Handle(args, index + 1, false, "/ckms-domain", "")
                        ? args[index + 1] : "");
                    return true;

                case "skhc":
                    SetHostCachingDisable(false);
                    return true;

                case "ckhc":
                    SetHostCachingDisable(true);
                    return true;

                case "sai":
                    if (!Handle(args, index + 1, true, "/sai", "<Activation Interval>")) return true;
                    SetActivationInterval(int.Parse(args[index + 1]));
                    return true;

                case "sri":
                    if (!Handle(args, index + 1, true, "/sri", "<Renewal Interval>")) return true;
                    SetRenewalInterval(int.Parse(args[index + 1]));
                    return true;

                case "sprt":
                    if (!Handle(args, index + 1, true, "/sprt", "<Port>")) return true;
                    SetKmsListenPort(args[index + 1]);
                    return true;

                case "sdns":
                    SetDnsPublishingDisabled(false);
                    return true;

                case "cdns":
                    SetDnsPublishingDisabled(true);
                    return true;

                case "spri":
                    SetKmsLowPriority(false);
                    return true;

                case "cpri":
                    SetKmsLowPriority(true);
                    return true;

                case "act-type":
                    {
                        int? type = null;
                        if (Handle(args, index + 1, false, "/act-type", ""))
                            type = int.Parse(args[index + 1]);
                        string actId = Handle(args, index + 2, false, "/act-type", "")
                            ? args[index + 2] : "";
                        SetVLActivationType(type, actId);
                    }
                    return true;

                // ---- Token-based activation ----
                case "lil":
                    TkaHelper.ListILs(LineOut);
                    return true;

                case "ril":
                    if (!Handle(args, index + 2, true, "/ril", "<ILID> <ILvID>")) return true;
                    TkaHelper.RemoveIL(LineOut, args[index + 1], args[index + 2]);
                    return true;

                case "ltc":
                    TkaHelper.ListCerts(LineOut);
                    return true;

                case "fta":
                    if (Handle(args, index + 2, false, "/fta", ""))
                        TkaHelper.Activate(LineOut, PrintActivatedStatus,
                            args[index + 1], args[index + 2]);
                    else if (Handle(args, index + 1, true, "/fta", ""))
                        TkaHelper.Activate(LineOut, PrintActivatedStatus,
                            args[index + 1], "");
                    return true;

                // ---- Active Directory activation ----
                case "ad-activation-get-iid":
                    if (!Handle(args, index + 1, true, "/ad-activation-get-iid", "<Product Key>")) return true;
                    AdGetIID(args[index + 1]);
                    return true;

                case "ad-activation-online":
                    if (!Handle(args, index + 1, true, "/ad-activation-online", "<Product Key>")) return true;
                    AdActivateOnline(args[index + 1],
                        Handle(args, index + 2, false, "/ad-activation-online", "")
                            ? args[index + 2] : "");
                    return true;

                case "ad-activation-apply-cid":
                    if (!Handle(args, index + 1, true, "/ad-activation-apply-cid", "<Product Key>")) return true;
                    if (!Handle(args, index + 2, true, "/ad-activation-apply-cid", "<Confirmation ID>")) return true;
                    AdActivatePhone(args[index + 1], args[index + 2],
                        Handle(args, index + 3, false, "/ad-activation-apply-cid", "")
                            ? args[index + 3] : "");
                    return true;

                case "ao-list":
                    AdHelper.ListActivationObjects(LineOut);
                    return true;

                case "del-ao":
                    if (!Handle(args, index + 1, true, "/del-ao", "<Activation Object DN | RDN>")) return true;
                    AdHelper.DeleteActivationObject(LineOut, args[index + 1]);
                    return true;

                default:
                    return false;
            }
        }

        private static bool Handle(string[] args, int idx, bool mustProvide,
                                   string opt, string param)
        {
            if (idx >= args.Length)
            {
                if (mustProvide)
                {
                    LineOut("");
                    LineOut(Resources.Get("L_MsgErrorText_9")
                        .Replace("%OPTION%", opt).Replace("%PARAM%", param));
                    DisplayUsage();
                }
                return false;
            }
            return true;
        }

        // -----------------------------------------------------------------
        // Usage text
        // -----------------------------------------------------------------

        private static void DisplayUsage()
        {
            LineOut(Resources.Get("L_MsgHelp_1"));
            LineOut(Resources.Get("L_MsgHelp_2"));
            LineOut("           " + Resources.Get("L_MsgHelp_3"));
            LineOut("           " + Resources.Get("L_MsgHelp_4"));
            LineOut("           " + Resources.Get("L_MsgHelp_5"));
            LineOut("");
            LineOut(Resources.Get("L_MsgGlobalOptions"));
            OptLine("ipk", "<Product Key>", "L_optInstallProductKeyUsage");
            OptLine("ato", "[Activation ID]", "L_optActivateProductUsage");
            OptLine("dli", "[Activation ID | All]", "L_optDisplayInformationUsage");
            OptLine("dlv", "[Activation ID | All]", "L_optDisplayInformationUsageVerbose");
            OptLine("xpr", "[Activation ID]", "L_optExpirationDatimeUsage");

            LineFlush("");

            LineOut(Resources.Get("L_MsgAdvancedOptions"));
            OptLine("cpky", "", "L_optClearPKeyFromRegistryUsage");
            OptLine("ilc", "<License file>", "L_optInstallLicenseUsage");
            OptLine("rilc", "", "L_optReinstallLicensesUsage");
            OptLine("rearm", "", "L_optReArmWindowsUsage");
            OptLine("rearm-app", "<Application ID>", "L_optReArmApplicationUsage");
            OptLine("rearm-sku", "<Activation ID>", "L_optReArmSkuUsage");
            OptLine("upk", "[Activation ID]", "L_optUninstallProductKeyUsage");
            LineOut("");
            OptLine("dti", "[Activation ID]", "L_optDisplayIIDUsage");
            OptLine2("atp", "<Confirmation ID>", "[Activation ID]", "L_optPhoneActivateProductUsage");

            LineOut("");
            LineOut(Resources.Get("L_MsgKmsClientOptions"));
            OptLine2("skms", "<Name[:Port] | : port>", "[Activation ID]", "L_optSetKmsNameUsage");
            OptLine("ckms", "[Activation ID]", "L_optClearKmsNameUsage");
            OptLine2("skms-domain", "<FQDN>", "[Activation ID]", "L_optSetKmsLookupDomainUsage");
            OptLine("ckms-domain", "[Activation ID]", "L_optClearKmsLookupDomainUsage");
            OptLine("skhc", "", "L_optSetKmsHostCachingUsage");
            OptLine("ckhc", "", "L_optClearKmsHostCachingUsage");

            LineFlush("");

            LineOut(Resources.Get("L_MsgTkaClientOptions"));
            OptLine("lil", "", "L_optListInstalledILsUsage");
            OptLine("ril", "<ILID> <ILvID>", "L_optRemoveInstalledILUsage");
            OptLine("ltc", "", "L_optListTkaCertsUsage");
            OptLine("fta", "<Certificate Thumbprint> [<PIN>]", "L_optForceTkaActivationUsage");

            LineFlush("");

            LineOut(Resources.Get("L_MsgKmsOptions"));
            OptLine("sprt", "<Port>", "L_optSetKmsListenPortUsage");
            OptLine("sai", "<Activation Interval>", "L_optSetActivationIntervalUsage");
            OptLine("sri", "<Renewal Interval>", "L_optSetRenewalIntervalUsage");
            OptLine("sdns", "", "L_optSetDNSUsage");
            OptLine("cdns", "", "L_optClearDNSUsage");
            OptLine("spri", "", "L_optSetNormalPriorityUsage");
            OptLine("cpri", "", "L_optClearNormalPriorityUsage");
            OptLine2("act-type", "[Activation-Type]", "[Activation ID]", "L_optSetVLActivationTypeUsage");

            LineFlush("");

            LineOut(Resources.Get("L_MsgADOptions"));
            OptLine2("ad-activation-online", "<Product Key>", "[Activation Object name]", "L_optADActivateUsage");
            OptLine("ad-activation-get-iid", "<Product Key>", "L_optADGetIIDUsage");
            OptLine3("ad-activation-apply-cid", "<Product Key>", "<Confirmation ID>", "[Activation Object name]", "L_optADApplyCIDUsage");
            OptLine("ao-list", "", "L_optADListAOsUsage");
            OptLine("del-ao", "<Activation Object DN | RDN>", "L_optADDeleteAOsUsage");

            Exit(1);
        }

        private static void OptLine(string opt, string p, string usageKey)
        {
            LineOut("/" + opt + " " + p);
            LineOut("    " + Resources.Get(usageKey));
        }

        private static void OptLine2(string opt, string p1, string p2, string usageKey)
        {
            LineOut("/" + opt + " " + p1 + " " + p2);
            LineOut("    " + Resources.Get(usageKey));
        }

        private static void OptLine3(string opt, string p1, string p2, string p3, string usageKey)
        {
            LineOut("/" + opt + " " + p1 + " " + p2 + " " + p3);
            LineOut("    " + Resources.Get(usageKey));
        }

        // -----------------------------------------------------------------
        // Classification helpers (mirrors slmgr.vbs)
        // -----------------------------------------------------------------

        private static bool IsKmsClient(string d) { return d.Contains("VOLUME_KMSCLIENT"); }
        private static bool IsKmsServer(string d) { return !IsKmsClient(d) && d.Contains("VOLUME_KMS"); }
        private static bool IsTbl(string d) { return d.Contains("TIMEBASED_"); }
        private static bool IsAvma(string d) { return d.Contains("VIRTUAL_MACHINE_ACTIVATION"); }
        private static bool IsMak(string d) { return d.Contains("MAK"); }

        private static int GetPrimaryWindowsSku(ManagementObject p)
        {
            if (WmiHelper.Str(p, "ApplicationId").ToLowerInvariant() != WindowsAppIdHolder.WindowsAppId)
                return 0;
            if (WmiHelper.Str(p, "PartialProductKey") == "") return 0;

            object addon = p["LicenseIsAddon"];
            if (addon is bool) return ((bool)addon) ? 0 : 1;

            string d = WmiHelper.Str(p, "Description");
            return (IsKmsClient(d) || IsKmsServer(d)) ? 1 : 2;
        }

        private static bool CheckProductForCommand(ManagementObject p, string actId)
        {
            if (actId == "" &&
                WmiHelper.Str(p, "ApplicationId").ToLowerInvariant() == WindowsAppIdHolder.WindowsAppId &&
                !WmiHelper.B(p, "LicenseIsAddon"))
                return true;

            return WmiHelper.Str(p, "ID").ToLowerInvariant() == actId;
        }

        private static void OutputIndeterminateOperationWarning(ManagementObject p)
        {
            LineOut(Resources.Get("L_MsgUndeterminedPrimaryKeyOperation"));
            LineOut(Resources.Get("L_MsgUndeterminedOperationFormat")
                .Replace("%PRODUCTDESCRIPTION%", WmiHelper.Str(p, "Description"))
                .Replace("%PRODUCTID%", WmiHelper.Str(p, "ID")));
        }

        // -----------------------------------------------------------------
        // Error output
        // -----------------------------------------------------------------

        private static void ShowError(string message, uint hr, string errText)
        {
            string strNumber;
            if (hr >= 0x80000000)
                strNumber = "0x" + hr.ToString("X");   // VBScript Hex() has no leading zeros
            else
                strNumber = hr.ToString();

            string strDescription = Resources.GetError(hr);
            if (string.IsNullOrEmpty(strDescription) || strDescription.StartsWith("["))
            {
                if (string.IsNullOrEmpty(errText))
                    strDescription = Resources.Get("L_MsgErrorText_6")
                        .Replace("0x%ERRCODE%", strNumber);
                else
                    strDescription = errText;
            }

            if (message.IndexOf("0x%ERRCODE%") < 0) message += "0x%ERRCODE%";
            if (message.IndexOf("%ERRTEXT%") < 0) message += " %ERRTEXT%";

            message = message
                .Replace("%COMPUTERNAME%", _computer)
                .Replace("0x%ERRCODE%", strNumber)
                .Replace("%ERRTEXT%", strDescription);

            LineOut(message);
        }

        // -----------------------------------------------------------------
        // /ipk
        // -----------------------------------------------------------------

        private static void InstallProductKey(string key)
        {
            try
            {
                ManagementObject svc = WmiHelper.GetService("Version");
                string ver = WmiHelper.Str(svc, "Version");

                WmiHelper.Invoke(svc, "InstallProductKey",
                    new string[] { "ProductKey" }, new object[] { key });
                WmiHelper.Invoke(svc, "RefreshLicenseStatus");

                bool isKms = false;
                foreach (ManagementObject p in WmiHelper.GetProducts(
                    "ID, ApplicationId, PartialProductKey, LicenseIsAddon, Description, Name",
                    "PartialProductKey <> null"))
                {
                    if (GetPrimaryWindowsSku(p) == 2) OutputIndeterminateOperationWarning(p);
                    if (IsKmsServer(WmiHelper.Str(p, "Description"))) { isKms = true; break; }
                }

                if (isKms) SetKmsVersion(ver);
                else ClearKmsVersion();

                LineOut(Resources.Get("L_MsgInstalledPKey").Replace("%PKEY%", key));
            }
            catch (Exception ex)
            {
                ShowError(Resources.Get("L_MsgErrorText_8"),
                    unchecked((uint)ex.HResult), ex.Message);
                Exit(unchecked((int)ex.HResult));
            }
        }

        // -----------------------------------------------------------------
        // /upk
        // -----------------------------------------------------------------

        private static void UninstallProductKey(string actId)
        {
            actId = actId.ToLowerInvariant();

            ManagementObject svc = WmiHelper.GetService("Version");
            string ver = WmiHelper.Str(svc, "Version");
            bool kmsFound = false;
            bool done = false;

            foreach (ManagementObject p in WmiHelper.GetProducts(
                "ID, ApplicationId, PartialProductKey, LicenseIsAddon, Description, Name, ProductKeyID",
                "PartialProductKey <> null"))
            {
                string desc = WmiHelper.Str(p, "Description");

                if (CheckProductForCommand(p, actId))
                {
                    int pri = GetPrimaryWindowsSku(p);
                    if (actId == "" && pri == 2) OutputIndeterminateOperationWarning(p);

                    WmiHelper.Invoke(p, "UninstallProductKey");
                    WmiHelper.Invoke(svc, "RefreshLicenseStatus");

                    if (actId != "" || pri == 1) done = true;
                    LineOut(Resources.Get("L_MsgUninstalledPKey"));
                }
                else if (IsKmsServer(desc))
                {
                    kmsFound = true;
                }

                if (kmsFound && done) break;
            }

            if (kmsFound) SetKmsVersion(ver);
            else ClearKmsVersion();

            if (!done) LineOut(Resources.Get("L_MsgErrorPKey"));
        }

        // -----------------------------------------------------------------
        // /ato
        // -----------------------------------------------------------------

        private static void ActivateProduct(string actId)
        {
            actId = actId.ToLowerInvariant();

            ManagementObject svc = WmiHelper.GetService("Version");
            bool found = false;

            foreach (ManagementObject p in WmiHelper.GetProducts(
                "ID, ApplicationId, PartialProductKey, LicenseIsAddon, Description, Name, LicenseStatus, VLActivationTypeEnabled",
                "PartialProductKey <> null"))
            {
                if (!CheckProductForCommand(p, actId)) continue;

                int pri = GetPrimaryWindowsSku(p);
                if (actId == "" && pri == 2) OutputIndeterminateOperationWarning(p);

                if (WmiHelper.U32(p, "VLActivationTypeEnabled") == 3)
                {
                    LineOut(Resources.Get("L_MsgTokenBasedActivationMustBeDone"));
                    return;
                }

                LineOut(Resources.Get("L_MsgActivating")
                    .Replace("%PRODUCTNAME%", WmiHelper.Str(p, "Name"))
                    .Replace("%PRODUCTID%", WmiHelper.Str(p, "ID")));

                string desc = WmiHelper.Str(p, "Description");
                if (!IsMak(desc) || WmiHelper.U32(p, "LicenseStatus") != 1)
                {
                    WmiHelper.Invoke(p, "Activate");
                    WmiHelper.Invoke(svc, "RefreshLicenseStatus");
                    p.Get();
                }
                PrintActivatedStatus(p);

                found = true;
                if (actId != "" || pri == 1) return;
            }

            if (!found) LineOut(Resources.Get("L_MsgErrorProductNotFound"));
        }

        private static void PrintActivatedStatus(ManagementObject p)
        {
            uint s = WmiHelper.U32(p, "LicenseStatus");
            uint r = WmiHelper.U32(p, "LicenseStatusReason");

            if (s == 1)
                LineOut(Resources.Get("L_MsgActivated"));
            else if (s == 4)
                LineOut(Resources.Get("L_MsgErrorText_8") + Resources.Get("L_MsgErrorText_11"));
            else if (s == 5 && r == 0xC004F200)
                LineOut(Resources.Get("L_MsgErrorText_8") + Resources.Get("L_MsgErrorText_12"));
            else if (s == 6)
            {
                LineOut(Resources.Get("L_MsgActivated"));
                LineOut(Resources.Get("L_MsgLicenseStatusExtendedGrace_1"));
            }
            else
                LineOut(Resources.Get("L_MsgActivated_Failed"));
        }

        // -----------------------------------------------------------------
        // /atp
        // -----------------------------------------------------------------

        private static void PhoneActivateProduct(string cid, string actId)
        {
            actId = actId.ToLowerInvariant();

            ManagementObject svc = WmiHelper.GetService("Version");
            bool found = false;

            foreach (ManagementObject p in WmiHelper.GetProducts(
                "ID, ApplicationId, PartialProductKey, LicenseIsAddon, Description, Name, " +
                "OfflineInstallationId, LicenseStatus, LicenseStatusReason",
                "PartialProductKey <> null"))
            {
                if (!CheckProductForCommand(p, actId)) continue;

                int pri = GetPrimaryWindowsSku(p);
                if (actId == "" && pri == 2) OutputIndeterminateOperationWarning(p);

                WmiHelper.Invoke(p, "DepositOfflineConfirmationId",
                    new string[] { "InstallationId", "ConfirmationId" },
                    new object[] { WmiHelper.Str(p, "OfflineInstallationId"), cid });

                WmiHelper.Invoke(svc, "RefreshLicenseStatus");
                p.Get();

                uint s = WmiHelper.U32(p, "LicenseStatus");
                if (s == 1)
                    LineOut(Resources.Get("L_MsgConfID").Replace("%ACTID%", WmiHelper.Str(p, "ID")));
                else
                    PrintActivatedStatus(p);

                found = true;
                if (actId != "" || pri == 1) return;
            }

            if (!found) LineOut(Resources.Get("L_MsgErrorProductNotFound"));
        }

        // -----------------------------------------------------------------
        // /dti
        // -----------------------------------------------------------------

        private static void DisplayIID(string actId)
        {
            actId = actId.ToLowerInvariant();
            bool found = false;

            foreach (ManagementObject p in WmiHelper.GetProducts(
                "ID, ApplicationId, PartialProductKey, LicenseIsAddon, Description, Name, OfflineInstallationId",
                "PartialProductKey <> null"))
            {
                if (!CheckProductForCommand(p, actId)) continue;

                int pri = GetPrimaryWindowsSku(p);
                if (actId == "" && pri == 2) OutputIndeterminateOperationWarning(p);

                LineOut(Resources.Get("L_MsgInstallationID") +
                        WmiHelper.Str(p, "OfflineInstallationId"));
                found = true;

                if (actId != "" || pri == 1) return;
            }

            if (found)
            {
                LineOut("");
                LineOut(Resources.Get("L_MsgPhoneNumbers"));
            }
            else
            {
                LineOut(Resources.Get("L_MsgErrorProductNotFound"));
            }
        }

        // -----------------------------------------------------------------
        // /dli /dlv
        // -----------------------------------------------------------------

        private static void DisplayAllInformation(string parm, bool verbose)
        {
            parm = parm.ToLowerInvariant();
            bool productKeyFound = false;

            string svcSelect =
                "SubscriptionType, SubscriptionStatus, SubscriptionEdition, SubscriptionExpiry, " +
                "KeyManagementServiceListeningPort, KeyManagementServiceDnsPublishing, " +
                "KeyManagementServiceLowPriority, ClientMachineId, KeyManagementServiceHostCaching, Version";

            string prodSelect =
                "ID, ApplicationId, PartialProductKey, LicenseIsAddon, Description, Name, " +
                "ProductKeyID, ProductKeyChannel, OfflineInstallationId, " +
                "ProcessorURL, MachineURL, UseLicenseURL, ProductKeyURL, ValidationURL, " +
                "GracePeriodRemaining, LicenseStatus, LicenseStatusReason, EvaluationEndDate, " +
                "VLRenewalInterval, VLActivationInterval, KeyManagementServiceLookupDomain, " +
                "KeyManagementServiceMachine, KeyManagementServicePort, " +
                "DiscoveredKeyManagementServiceMachineName, DiscoveredKeyManagementServiceMachinePort, " +
                "DiscoveredKeyManagementServiceMachineIpAddress, KeyManagementServiceProductKeyID, " +
                "TokenActivationILID, TokenActivationILVID, TokenActivationGrantNumber, " +
                "TokenActivationCertificateThumbprint, TokenActivationAdditionalInfo, TrustedTime, " +
                "ADActivationObjectName, ADActivationObjectDN, ADActivationCsvlkPid, ADActivationCsvlkSkuId, " +
                "VLActivationTypeEnabled, VLActivationType, IAID, " +
                "AutomaticVMActivationHostMachineName, AutomaticVMActivationLastActivationTime, " +
                "AutomaticVMActivationHostDigitalPid2";

            if (verbose)
            {
                svcSelect = "RemainingWindowsReArmCount, " + svcSelect;
                prodSelect = "RemainingAppReArmCount, RemainingSkuReArmCount, " + prodSelect;
            }

            ManagementObject svc = WmiHelper.GetService(svcSelect);
            if (verbose)
                LineOut(Resources.Get("L_MsgServiceVersion") + WmiHelper.Str(svc, "Version"));

            string iterSelect = parm == "all"
                ? prodSelect
                : "ID, ApplicationId, PartialProductKey, LicenseIsAddon, Description, Name";

            foreach (ManagementObject pi in WmiHelper.GetProducts(iterSelect, ""))
            {
                string id = WmiHelper.Str(pi, "ID");
                int pri = GetPrimaryWindowsSku(pi);
                bool useDefault = false;
                bool show = false;

                if (parm == "" && (pri == 1 || pri == 2)) { useDefault = true; show = true; }
                if (parm == "" && WmiHelper.B(pi, "LicenseIsAddon") &&
                    WmiHelper.Str(pi, "PartialProductKey") != "") show = true;
                if (parm == "all") show = true;
                if (parm == id.ToLowerInvariant()) show = true;
                if (!show) continue;

                ManagementObject p = parm == "all"
                    ? pi
                    : WmiHelper.GetProduct(prodSelect, "id = '" + id + "'");

                if (useDefault && pri == 2) OutputIndeterminateOperationWarning(p);

                productKeyFound = true;
                PrintProduct(p, svc, id, verbose, pri);

                if (parm != "all" && parm == id.ToLowerInvariant()) break;
                LineOut("");
            }

            PrintSubscription(svc);
            if (!productKeyFound) LineOut(Resources.Get("L_MsgErrorPKey"));
        }

        private static void PrintProduct(ManagementObject p, ManagementObject svc,
                                         string id, bool verbose, int pri)
        {
            LineOut("");
            LineOut(Resources.Get("L_MsgProductName") + WmiHelper.Str(p, "Name"));
            LineOut(Resources.Get("L_MsgProductDesc") + WmiHelper.Str(p, "Description"));

            if (WmiHelper.Str(p, "TokenActivationAdditionalInfo") != "")
                LineOut(Resources.Get("L_MsgTkaInfoAdditionalInfo")
                    .Replace("%MOREINFO%", WmiHelper.Str(p, "TokenActivationAdditionalInfo")));

            string d = WmiHelper.Str(p, "Description");
            bool kmsClient = IsKmsClient(d);
            bool kmsServer = IsKmsServer(d);
            bool tbl = IsTbl(d);
            bool avma = IsAvma(d);

            if (verbose)
            {
                LineOut(Resources.Get("L_MsgActID") + id);
                LineOut(Resources.Get("L_MsgAppID") + WmiHelper.Str(p, "ApplicationID"));
                LineOut(Resources.Get("L_MsgPID4") + WmiHelper.Str(p, "ProductKeyID"));
                LineOut(Resources.Get("L_MsgChannel") + WmiHelper.Str(p, "ProductKeyChannel"));
                LineOut(Resources.Get("L_MsgInstallationID") + WmiHelper.Str(p, "OfflineInstallationId"));

                if (!kmsClient && !avma)
                {
                    PrintUrlIfPresent(p, "ProcessorURL", "L_MsgProcessorCertUrl");
                    PrintUrlIfPresent(p, "MachineURL", "L_MsgMachineCertUrl");
                    PrintUrlIfPresent(p, "UseLicenseURL", "L_MsgUseLicenseCertUrl");
                    PrintUrlIfPresent(p, "ProductKeyURL", "L_MsgPKeyCertUrl");
                    PrintUrlIfPresent(p, "ValidationURL", "L_MsgValidationUrl");
                }
            }

            if (WmiHelper.Str(p, "PartialProductKey") != "")
                LineOut(Resources.Get("L_MsgPartialPKey") + WmiHelper.Str(p, "PartialProductKey"));
            else
                LineOut(Resources.Get("L_MsgErrorLicenseNotInUse"));

            PrintLicenseStatus(p, tbl, avma, verbose);

            if (verbose)
            {
                if (WmiHelper.Str(p, "ApplicationId").ToLowerInvariant() == WindowsAppIdHolder.WindowsAppId)
                    LineOut(Resources.Get("L_MsgRemainingWindowsRearmCount")
                        .Replace("%COUNT%", WmiHelper.Str(svc, "RemainingWindowsReArmCount")));
                else
                    LineOut(Resources.Get("L_MsgRemainingAppRearmCount")
                        .Replace("%COUNT%", WmiHelper.Str(p, "RemainingAppReArmCount")));

                LineOut(Resources.Get("L_MsgRemainingSkuRearmCount")
                    .Replace("%COUNT%", WmiHelper.Str(p, "RemainingSkuReArmCount")));

                DateTime? tt = WmiHelper.Date(p, "TrustedTime");
                if (tt.HasValue)
                    LineOut(Resources.Get("L_MsgCurrentTrustedTime") + tt.Value);
            }

            if (kmsClient) PrintKmsClientSection(svc, p);
            if (kmsServer || pri == 1 || pri == 2) PrintKmsServerSection(svc, p);
            if (avma) PrintAvmaSection(p);
        }

        private static void PrintUrlIfPresent(ManagementObject p, string prop, string labelKey)
        {
            string v = WmiHelper.Str(p, prop);
            if (v != "") LineOut(Resources.Get(labelKey) + v);
        }

        private static void PrintLicenseStatus(ManagementObject p, bool tbl, bool avma, bool verbose)
        {
            uint ls = WmiHelper.U32(p, "LicenseStatus");
            uint r = WmiHelper.U32(p, "LicenseStatusReason");
            uint gp = WmiHelper.U32(p, "GracePeriodRemaining");

            if (ls == 0)
            {
                LineOut(Resources.Get("L_MsgLicenseStatusUnlicensed_1"));
            }
            else if (ls == 1)
            {
                LineOut(Resources.Get("L_MsgLicenseStatusLicensed_1"));
                if (gp != 0)
                {
                    uint days = DaysFromMins(gp);
                    string fmtKey = tbl ? "L_MsgLicenseStatusTBL_1"
                                  : avma ? "L_MsgLicenseStatusAVMA_1"
                                  : "L_MsgLicenseStatusVL_1";
                    LineOut(Resources.Get(fmtKey)
                        .Replace("%MINUTE%", gp.ToString())
                        .Replace("%DAY%", days.ToString()));
                }
            }
            else if (ls == 2 || ls == 3 || ls == 4 || ls == 6)
            {
                string key;
                if (ls == 2) key = "L_MsgLicenseStatusInitialGrace_1";
                else if (ls == 3) key = "L_MsgLicenseStatusAdditionalGrace_1";
                else if (ls == 4) key = "L_MsgLicenseStatusNonGenuineGrace_1";
                else key = "L_MsgLicenseStatusExtendedGrace_1";

                LineOut(Resources.Get(key));
                LineOut(Resources.Get("L_MsgLicenseStatusTimeRemaining")
                    .Replace("%MINUTE%", gp.ToString())
                    .Replace("%DAY%", DaysFromMins(gp).ToString()));
            }
            else if (ls == 5)
            {
                LineOut(Resources.Get("L_MsgLicenseStatusNotification_1"));
                string hex = "0x" + r.ToString("X");   // VBScript Hex() has no leading zeros
                if (r == 0xC004F200)
                    LineOut(Resources.Get("L_MsgNotificationErrorReasonNonGenuine")
                        .Replace("%ERRCODE%", hex));
                else if (r == 0xC004F009)
                    LineOut(Resources.Get("L_MsgNotificationErrorReasonExpiration")
                        .Replace("%ERRCODE%", hex));
                else
                    LineOut(Resources.Get("L_MsgNotificationErrorReasonOther")
                        .Replace("%ERRCODE%", hex));
            }
            else
            {
                LineOut(Resources.Get("L_MsgLicenseStatusUnknown"));
            }

            if (ls != 0 && verbose)
            {
                DateTime? ev = WmiHelper.Date(p, "EvaluationEndDate");
                if (ev.HasValue)
                    LineOut(Resources.Get("L_MsgLicenseStatusEvalEndData") + ev.Value);
            }
        }

        private static uint DaysFromMins(uint mins)
        {
            // Round up to the nearest day (ceil).
            return (mins + 1439) / 1440;
        }

        private static void PrintKmsClientSection(ManagementObject svc, ManagementObject p)
        {
            uint t = WmiHelper.U32(p, "VLActivationTypeEnabled");
            if (t == 1) LineOut(Resources.Get("L_MsgVLActivationTypeAD"));
            else if (t == 2) LineOut(Resources.Get("L_MsgVLActivationTypeKMS"));
            else if (t == 3) LineOut(Resources.Get("L_MsgVLActivationTypeToken"));
            else LineOut(Resources.Get("L_MsgVLActivationTypeAll"));

            if (WmiHelper.U32(p, "VLActivationType") == 1)
                DisplayAdClientInfo(svc, p);
            else if (IsTokenActivated(p))
                DisplayTkaClientInfo(svc, p);
            else if (WmiHelper.U32(p, "LicenseStatus") != 1)
                LineOut(Resources.Get("L_MsgPleaseActivateRefreshKMSInfo"));
            else
                DisplayKmsClientInfo(svc, p);
        }

        private static bool IsTokenActivated(ManagementObject p)
        {
            object v = p["TokenActivationILVID"];
            if (v == null) return false;
            return Convert.ToUInt32(v) != 0xFFFFFFFF;
        }

        private static void DisplayKmsClientInfo(ManagementObject svc, ManagementObject p)
        {
            LineOut("");
            LineOut(Resources.Get("L_MsgVLMostRecentActivationInfo"));
            LineOut(Resources.Get("L_MsgKmsInfo"));
            LineOut("    " + Resources.Get("L_MsgCmid") + WmiHelper.Str(svc, "ClientMachineID"));

            string lookup = WmiHelper.Str(p, "KeyManagementServiceLookupDomain");
            bool bLookup = lookup != "";
            if (bLookup)
                LineOut("    " + Resources.Get("L_MsgKmsLookupDomain") + lookup);

            string kms = WmiHelper.Str(p, "KeyManagementServiceMachine");
            bool bFixed = kms != "";
            string port;

            if (bFixed)
            {
                port = WmiHelper.Str(p, "KeyManagementServicePort");
                if (port == "0" || port == "") port = "1688";
                LineOut("    " + Resources.Get("L_MsgRegisteredKmsName") + kms + ":" + port);
            }
            else
            {
                kms = WmiHelper.Str(p, "DiscoveredKeyManagementServiceMachineName");
                port = WmiHelper.Str(p, "DiscoveredKeyManagementServiceMachinePort");
                if (kms == "" || port == "0" || port == "")
                    LineOut("    " + Resources.Get("L_MsgKmsFromDnsUnavailable"));
                else
                    LineOut("    " + Resources.Get("L_MsgKmsFromDns") + kms + ":" + port);
            }

            string ip = WmiHelper.Str(p, "DiscoveredKeyManagementServiceMachineIpAddress");
            if (ip == "")
                LineOut("    " + Resources.Get("L_MsgKmsIpAddressUnavailable"));
            else
                LineOut("    " + Resources.Get("L_MsgKmsIpAddress") + ip);

            LineOut("    " + Resources.Get("L_MsgKmsPID4") +
                    WmiHelper.Str(p, "KeyManagementServiceProductKeyID"));
            LineOut("    " + Resources.Get("L_MsgActivationInterval")
                .Replace("%INTERVAL%", WmiHelper.Str(p, "VLActivationInterval")));
            LineOut("    " + Resources.Get("L_MsgRenewalInterval")
                .Replace("%INTERVAL%", WmiHelper.Str(p, "VLRenewalInterval")));

            if (WmiHelper.B(svc, "KeyManagementServiceHostCaching"))
                LineOut("    " + Resources.Get("L_MsgKmsHostCachingEnabled"));
            else
                LineOut("    " + Resources.Get("L_MsgKmsHostCachingDisabled"));

            if (bLookup && bFixed)
            {
                LineOut("");
                LineOut(Resources.Get("L_MsgKmsUseMachineNameOverrides")
                    .Replace("%KMS%", kms + ":" + port));
            }
        }

        private static void DisplayAdClientInfo(ManagementObject svc, ManagementObject p)
        {
            LineOut("");
            LineOut(Resources.Get("L_MsgVLMostRecentActivationInfo"));
            LineOut(Resources.Get("L_MsgADInfo"));
            LineOut("    " + Resources.Get("L_MsgADInfoAOName") + WmiHelper.Str(p, "ADActivationObjectName"));
            LineOut("    " + Resources.Get("L_MsgADInfoAODN") + WmiHelper.Str(p, "ADActivationObjectDN"));
            LineOut("    " + Resources.Get("L_MsgADInfoExtendedPid") + WmiHelper.Str(p, "ADActivationCsvlkPid"));
            LineOut("    " + Resources.Get("L_MsgADInfoActID") + WmiHelper.Str(p, "ADActivationCsvlkSkuId"));
        }

        private static void DisplayTkaClientInfo(ManagementObject svc, ManagementObject p)
        {
            LineOut("");
            LineOut(Resources.Get("L_MsgVLMostRecentActivationInfo"));
            LineOut(Resources.Get("L_MsgTkaInfo"));
            LineOut("    " + Resources.Get("L_MsgTkaInfoILID")
                .Replace("%ILID%", WmiHelper.Str(p, "TokenActivationILID")));
            LineOut("    " + Resources.Get("L_MsgTkaInfoILVID")
                .Replace("%ILVID%", WmiHelper.Str(p, "TokenActivationILVID")));
            LineOut("    " + Resources.Get("L_MsgTkaInfoGrantNo")
                .Replace("%GRANTNO%", WmiHelper.Str(p, "TokenActivationGrantNumber")));
            LineOut("    " + Resources.Get("L_MsgTkaInfoThumbprint")
                .Replace("%THUMBPRINT%", WmiHelper.Str(p, "TokenActivationCertificateThumbprint")));
        }

        private static void PrintKmsServerSection(ManagementObject svc, ManagementObject p)
        {
            string id = WmiHelper.Str(p, "ID");
            ManagementObject q = WmiHelper.GetProduct(
                "IsKeyManagementServiceMachine, KeyManagementServiceCurrentCount, " +
                "KeyManagementServiceTotalRequests, KeyManagementServiceFailedRequests, " +
                "KeyManagementServiceUnlicensedRequests, KeyManagementServiceLicensedRequests, " +
                "KeyManagementServiceOOBGraceRequests, KeyManagementServiceOOTGraceRequests, " +
                "KeyManagementServiceNonGenuineGraceRequests, KeyManagementServiceNotificationRequests",
                "id = '" + id + "'");

            if (q == null || WmiHelper.U32(q, "IsKeyManagementServiceMachine") == 0) return;

            LineOut("");
            LineOut(Resources.Get("L_MsgKmsEnabled"));
            LineOut("    " + Resources.Get("L_MsgKmsCurrentCount") +
                    WmiHelper.Str(q, "KeyManagementServiceCurrentCount"));

            uint lp = WmiHelper.U32(svc, "KeyManagementServiceListeningPort");
            LineOut("    " + Resources.Get("L_MsgKmsListeningOnPort") +
                    (lp == 0 ? "1688" : lp.ToString()));

            LineOut("    " + (WmiHelper.B(svc, "KeyManagementServiceDnsPublishing")
                ? Resources.Get("L_MsgKmsDnsPublishingEnabled")
                : Resources.Get("L_MsgKmsDnsPublishingDisabled")));

            LineOut("    " + (WmiHelper.B(svc, "KeyManagementServiceLowPriority")
                ? Resources.Get("L_MsgKmsPriLow")
                : Resources.Get("L_MsgKmsPriNormal")));

            LineOut("");
            LineOut(Resources.Get("L_MsgKmsCumulativeRequestsFromClients"));
            LineOut("    " + Resources.Get("L_MsgKmsTotalRequestsRecieved") +
                    WmiHelper.Str(q, "KeyManagementServiceTotalRequests"));
            LineOut("    " + Resources.Get("L_MsgKmsFailedRequestsReceived") +
                    WmiHelper.Str(q, "KeyManagementServiceFailedRequests"));
            LineOut("    " + Resources.Get("L_MsgKmsRequestsWithStatusUnlicensed") +
                    WmiHelper.Str(q, "KeyManagementServiceUnlicensedRequests"));
            LineOut("    " + Resources.Get("L_MsgKmsRequestsWithStatusLicensed") +
                    WmiHelper.Str(q, "KeyManagementServiceLicensedRequests"));
            LineOut("    " + Resources.Get("L_MsgKmsRequestsWithStatusInitialGrace") +
                    WmiHelper.Str(q, "KeyManagementServiceOOBGraceRequests"));
            LineOut("    " + Resources.Get("L_MsgKmsRequestsWithStatusLicenseExpiredOrHwidOot") +
                    WmiHelper.Str(q, "KeyManagementServiceOOTGraceRequests"));
            LineOut("    " + Resources.Get("L_MsgKmsRequestsWithStatusNonGenuineGrace") +
                    WmiHelper.Str(q, "KeyManagementServiceNonGenuineGraceRequests"));
            LineOut("    " + Resources.Get("L_MsgKmsRequestsWithStatusNotification") +
                    WmiHelper.Str(q, "KeyManagementServiceNotificationRequests"));
        }

        private static void PrintAvmaSection(ManagementObject p)
        {
            string iaid = WmiHelper.Str(p, "IAID");
            LineOut(Resources.Get("L_MsgAVMAID") +
                (iaid != "" ? iaid : Resources.Get("L_MsgNotAvailable")));

            LineOut("");
            LineOut(Resources.Get("L_MsgVLMostRecentActivationInfo"));
            LineOut(Resources.Get("L_MsgAVMAInfo"));

            string host = WmiHelper.Str(p, "AutomaticVMActivationHostMachineName");
            string pid = WmiHelper.Str(p, "AutomaticVMActivationHostDigitalPid2");
            DateTime? time = WmiHelper.Date(p, "AutomaticVMActivationLastActivationTime");

            LineOut("    " + Resources.Get("L_MsgAVMAHostMachineName") +
                (host != "" ? host : Resources.Get("L_MsgNotAvailable")));
            LineOut("    " + Resources.Get("L_MsgAVMALastActTime") +
                (time.HasValue ? time.Value.ToString() : Resources.Get("L_MsgNotAvailable")));
            LineOut("    " + Resources.Get("L_MsgAVMAHostPid2") +
                (pid != "" ? pid : Resources.Get("L_MsgNotAvailable")));
        }

        private static void PrintSubscription(ManagementObject svc)
        {
            // Only print the subscription block when SubscriptionType is a known
            // value (0 = user based, 1 = device based, 2 = Azure based). Any other
            // value — including a missing property on older Windows builds, or the
            // special 120 used by slmgr.vbs to mean "expired" — suppresses the
            // entire block. This avoids printing "Unknown" placeholders when the
            // underlying WMI property does not exist on this machine.
            uint t = WmiHelper.U32(svc, "SubscriptionType");
            if (t != 0 && t != 1 && t != 2) return;

            LineOut(Resources.Get("L_MsgSubscriptionEdition") +
                    WmiHelper.Str(svc, "SubscriptionEdition"));

            string typeStr;
            if (t == 0) typeStr = Resources.Get("L_MsgSubscriptionTypeUBS");
            else if (t == 1) typeStr = Resources.Get("L_MsgSubscriptionTypeDBS");
            else typeStr = Resources.Get("L_MsgSubscriptionTypeABS");
            LineOut(Resources.Get("L_MsgSubscriptionType") + typeStr);

            uint st = WmiHelper.U32(svc, "SubscriptionStatus");
            string statusStr;
            if (st == 120) statusStr = Resources.Get("L_MsgSubscriptionStatusExpired");
            else if (st == 100) statusStr = Resources.Get("L_MsgSubscriptionStatusDisabled");
            else if (st == 1) statusStr = Resources.Get("L_MsgSubscriptionStatusActive");
            else statusStr = Resources.Get("L_MsgSubscriptionStatusNotActive");
            LineOut(Resources.Get("L_MsgSubscriptionStatus") + statusStr);

            string exp = WmiHelper.Str(svc, "SubscriptionExpiry");
            LineOut(Resources.Get("L_MsgSubscriptionExpiry") +
                (exp != "" && exp != "<unspecified>"
                    ? exp
                    : Resources.Get("L_MsgSubscriptionExpiryUnknown")));
        }

        // -----------------------------------------------------------------
        // /xpr
        // -----------------------------------------------------------------

        private static void ExpirationDatetime(string actId)
        {
            actId = actId.ToLowerInvariant();

            string where = actId == ""
                ? "ApplicationId = '" + WindowsAppIdHolder.WindowsAppId + "'"
                : "ID = '" + actId.Replace("'", "") + "'";
            where += " AND PartialProductKey <> null";

            bool found = false;

            foreach (ManagementObject p in WmiHelper.GetProducts(
                "ID, ApplicationId, PartialProductKey, LicenseIsAddon, Description, Name, " +
                "LicenseStatus, GracePeriodRemaining", where))
            {
                uint ls = WmiHelper.U32(p, "LicenseStatus");
                uint gp = WmiHelper.U32(p, "GracePeriodRemaining");
                DateTime ends = DateTime.Now.AddMinutes(gp);
                found = true;

                int pri = GetPrimaryWindowsSku(p);
                if (actId == "" && pri == 2) OutputIndeterminateOperationWarning(p);

                string output = "";
                string desc = WmiHelper.Str(p, "Description");

                if (ls == 0)
                {
                    output = Resources.Get("L_MsgLicenseStatusUnlicensed");
                }
                else if (ls == 1)
                {
                    if (gp != 0)
                    {
                        string key = IsTbl(desc) ? "L_MsgLicenseStatusTBL"
                                   : IsAvma(desc) ? "L_MsgLicenseStatusAVMA"
                                   : "L_MsgLicenseStatusVL";
                        output = Resources.Get(key)
                            .Replace("%ENDDATE%", ends.ToString(CultureInfo.CurrentCulture));
                    }
                    else
                    {
                        output = Resources.Get("L_MsgLicenseStatusLicensed");
                    }
                }
                else if (ls == 2)
                    output = Resources.Get("L_MsgLicenseStatusInitialGrace")
                        .Replace("%ENDDATE%", ends.ToString(CultureInfo.CurrentCulture));
                else if (ls == 3)
                    output = Resources.Get("L_MsgLicenseStatusAdditionalGrace")
                        .Replace("%ENDDATE%", ends.ToString(CultureInfo.CurrentCulture));
                else if (ls == 4)
                    output = Resources.Get("L_MsgLicenseStatusNonGenuineGrace")
                        .Replace("%ENDDATE%", ends.ToString(CultureInfo.CurrentCulture));
                else if (ls == 5)
                    output = Resources.Get("L_MsgLicenseStatusNotification");
                else if (ls == 6)
                    output = Resources.Get("L_MsgLicenseStatusExtendedGrace")
                        .Replace("%ENDDATE%", ends.ToString(CultureInfo.CurrentCulture));

                if (output != "")
                {
                    LineOut(WmiHelper.Str(p, "Name") + ":");
                    LineOut("    " + output);
                }
            }

            if (!found) LineOut(Resources.Get("L_MsgErrorPKey"));
        }

        // -----------------------------------------------------------------
        // /cpky
        // -----------------------------------------------------------------

        private static void ClearPKeyFromRegistry()
        {
            ManagementObject svc = WmiHelper.GetService("Version");
            WmiHelper.Invoke(svc, "ClearProductKeyFromRegistry");
            LineOut(Resources.Get("L_MsgClearedPKey"));
        }

        // -----------------------------------------------------------------
        // /ilc /rilc
        // -----------------------------------------------------------------

        private static void InstallLicense(string file)
        {
            try
            {
                byte[] bytes = File.ReadAllBytes(file);
                string text = Encoding.Unicode.GetString(bytes);

                ManagementObject svc = WmiHelper.GetService("Version");
                WmiHelper.Invoke(svc, "InstallLicense",
                    new string[] { "License" }, new object[] { text });

                LineOut(Resources.Get("L_MsgLicenseFile").Replace("%LICENSEFILE%", file));
                LineOut("");
            }
            catch (Exception ex)
            {
                ShowError(Resources.Get("L_MsgErrorText_8"),
                    unchecked((uint)ex.HResult), ex.Message);
                Exit(unchecked((int)ex.HResult));
            }
        }

        private static void ReinstallLicenses()
        {
            string root = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
            string spp = Path.Combine(Path.Combine(Path.Combine(root, "system32"), "spp"), "tokens");
            string oem = Path.Combine(Path.Combine(root, "system32"), "oem");

            LineOut(Resources.Get("L_MsgReinstallingLicenses"));

            if (Directory.Exists(spp))
            {
                string[] dirs = Directory.GetDirectories(spp);
                for (int i = 0; i < dirs.Length; i++) InstallLicenseTree(dirs[i]);
            }

            if (Directory.Exists(oem)) InstallLicenseTree(oem);

            LineOut(Resources.Get("L_MsgLicensesReinstalled"));
        }

        private static void InstallLicenseTree(string dir)
        {
            string[] files = Directory.GetFiles(dir, "*.xrm-ms");
            for (int i = 0; i < files.Length; i++) InstallLicense(files[i]);

            string[] dirs = Directory.GetDirectories(dir);
            for (int i = 0; i < dirs.Length; i++) InstallLicenseTree(dirs[i]);
        }

        // -----------------------------------------------------------------
        // /rearm /rearm-app /rearm-sku
        // -----------------------------------------------------------------

        private static void ReArmWindows()
        {
            ManagementObject svc = WmiHelper.GetService("Version");
            WmiHelper.Invoke(svc, "ReArmWindows");
            LineOut(Resources.Get("L_MsgRearm_1"));
            LineOut(Resources.Get("L_MsgRearm_2"));
        }

        private static void ReArmApp(string appId)
        {
            ManagementObject svc = WmiHelper.GetService("Version");
            WmiHelper.Invoke(svc, "ReArmApp",
                new string[] { "ApplicationId" }, new object[] { appId });
            LineOut(Resources.Get("L_MsgRearm_1"));
        }

        private static void ReArmSku(string skuId)
        {
            skuId = skuId.ToLowerInvariant();
            bool found = false;

            foreach (ManagementObject p in WmiHelper.GetProducts("ID", "ID = '" + skuId + "'"))
            {
                found = true;
                WmiHelper.Invoke(p, "ReArmSku");
                LineOut(Resources.Get("L_MsgRearm_1"));
                break;
            }

            if (!found) LineOut(Resources.Get("L_MsgErrorProductNotFound"));
        }

        // -----------------------------------------------------------------
        // KMS client management
        // -----------------------------------------------------------------

        private static ManagementObject GetKmsTarget(string actId)
        {
            actId = actId.ToLowerInvariant();

            if (actId == "")
                return WmiHelper.GetService(
                    "Version, KeyManagementServiceMachine, KeyManagementServicePort, " +
                    "KeyManagementServiceLookupDomain");

            foreach (ManagementObject p in WmiHelper.GetProducts(
                "ID, KeyManagementServiceMachine, KeyManagementServicePort, KeyManagementServiceLookupDomain",
                ""))
            {
                if (WmiHelper.Str(p, "ID").ToLowerInvariant() == actId) return p;
            }

            LineOut(Resources.Get("L_MsgErrorActivationID").Replace("%ActID%", actId));
            return null;
        }

        private static void SetKmsMachineName(string namePort, string actId)
        {
            ManagementObject target = GetKmsTarget(actId);
            if (target == null) return;

            string name;
            string port;

            int bracketEnd = namePort.IndexOf(']');
            if (namePort.StartsWith("[") && bracketEnd > 0)
            {
                // IPv6 form: [host]:port
                if (namePort.Length == bracketEnd + 1)
                {
                    name = namePort;
                    port = "";
                }
                else
                {
                    name = namePort.Substring(0, bracketEnd + 1);
                    port = namePort.Substring(bracketEnd + 2);
                }
            }
            else
            {
                int colon = namePort.IndexOf(':');
                if (colon >= 0)
                {
                    name = namePort.Substring(0, colon);
                    port = namePort.Substring(colon + 1);
                }
                else
                {
                    name = namePort;
                    port = "";
                }
            }

            if (name != "")
                WmiHelper.Invoke(target, "SetKeyManagementServiceMachine",
                    new string[] { "MachineName" }, new object[] { name });

            if (port != "")
                WmiHelper.Invoke(target, "SetKeyManagementServicePort",
                    new string[] { "Port" }, new object[] { uint.Parse(port) });
            else
                WmiHelper.Invoke(target, "ClearKeyManagementServicePort");

            LineOut(Resources.Get("L_MsgKmsNameSet").Replace("%KMS%", namePort));

            if (WmiHelper.Str(target, "KeyManagementServiceLookupDomain") != "")
                LineOut(Resources.Get("L_MsgKmsUseMachineNameOverrides").Replace("%KMS%", namePort));
        }

        private static void ClearKms(string actId)
        {
            ManagementObject t = GetKmsTarget(actId);
            if (t == null) return;

            WmiHelper.Invoke(t, "ClearKeyManagementServiceMachine");
            WmiHelper.Invoke(t, "ClearKeyManagementServicePort");
            LineOut(Resources.Get("L_MsgKmsNameCleared"));

            string d = WmiHelper.Str(t, "KeyManagementServiceLookupDomain");
            if (d != "") LineOut(Resources.Get("L_MsgKmsUseLookupDomain").Replace("%FQDN%", d));
        }

        private static void SetKmsLookupDomain(string domain, string actId)
        {
            ManagementObject t = GetKmsTarget(actId);
            if (t == null) return;

            WmiHelper.Invoke(t, "SetKeyManagementServiceLookupDomain",
                new string[] { "LookupDomain" }, new object[] { domain });

            LineOut(Resources.Get("L_MsgKmsLookupDomainSet").Replace("%FQDN%", domain));

            if (WmiHelper.Str(t, "KeyManagementServiceMachine") != "")
                LineOut(Resources.Get("L_MsgKmsUseMachineNameOverrides").Replace("%KMS%",
                    WmiHelper.Str(t, "KeyManagementServiceMachine") + ":" +
                    WmiHelper.Str(t, "KeyManagementServicePort")));
        }

        private static void ClearKmsLookupDomain(string actId)
        {
            ManagementObject t = GetKmsTarget(actId);
            if (t == null) return;

            WmiHelper.Invoke(t, "ClearKeyManagementServiceLookupDomain");
            LineOut(Resources.Get("L_MsgKmsLookupDomainCleared"));

            if (WmiHelper.Str(t, "KeyManagementServiceMachine") != "")
                LineOut(Resources.Get("L_MsgKmsUseMachineName").Replace("%KMS%",
                    WmiHelper.Str(t, "KeyManagementServiceMachine") + ":" +
                    WmiHelper.Str(t, "KeyManagementServicePort")));
        }

        private static void SetHostCachingDisable(bool disable)
        {
            ManagementObject svc = WmiHelper.GetService("Version");
            WmiHelper.Invoke(svc, "DisableKeyManagementServiceHostCaching",
                new string[] { "Disable" }, new object[] { disable });

            if (disable)
                LineOut(Resources.Get("L_MsgKmsHostCachingDisabled"));
            else
                LineOut(Resources.Get("L_MsgKmsHostCachingEnabled"));
        }

        // -----------------------------------------------------------------
        // KMS host configuration
        // -----------------------------------------------------------------

        private static void SetActivationInterval(int interval)
        {
            if (interval < 0)
            {
                LineOut(Resources.Get("L_MsgInvalidDataError"));
                return;
            }

            ManagementObject svc = WmiHelper.GetService("Version");
            bool kmsFlag = false;

            foreach (ManagementObject p in WmiHelper.GetProducts(
                "ID, IsKeyManagementServiceMachine", "PartialProductKey <> null"))
            {
                if (WmiHelper.U32(p, "IsKeyManagementServiceMachine") == 1)
                {
                    kmsFlag = true;
                    WmiHelper.Invoke(svc, "SetVLActivationInterval",
                        new string[] { "ActivationInterval" },
                        new object[] { (uint)interval });
                    LineOut(Resources.Get("L_MsgActivationSet")
                        .Replace("%ACTIVATION%", interval.ToString()));
                    LineOut(Resources.Get("L_MsgWarningKmsReboot"));
                    break;
                }
            }

            if (!kmsFlag) LineOut(Resources.Get("L_MsgWarningActivation"));
        }

        private static void SetRenewalInterval(int interval)
        {
            if (interval < 0)
            {
                LineOut(Resources.Get("L_MsgInvalidDataError"));
                return;
            }

            ManagementObject svc = WmiHelper.GetService("Version");
            bool kmsFlag = false;

            foreach (ManagementObject p in WmiHelper.GetProducts(
                "ID, IsKeyManagementServiceMachine", "PartialProductKey <> null"))
            {
                if (WmiHelper.U32(p, "IsKeyManagementServiceMachine") == 1)
                {
                    kmsFlag = true;
                    WmiHelper.Invoke(svc, "SetVLRenewalInterval",
                        new string[] { "RenewalInterval" },
                        new object[] { (uint)interval });
                    LineOut(Resources.Get("L_MsgRenewalSet")
                        .Replace("%RENEWAL%", interval.ToString()));
                    LineOut(Resources.Get("L_MsgWarningKmsReboot"));
                    break;
                }
            }

            if (!kmsFlag) LineOut(Resources.Get("L_MsgWarningRenewal"));
        }

        private static void SetKmsListenPort(string port)
        {
            ManagementObject svc = WmiHelper.GetService("Version");
            bool kmsFlag = false;

            foreach (ManagementObject p in WmiHelper.GetProducts(
                "ID, IsKeyManagementServiceMachine", "PartialProductKey <> null"))
            {
                if (WmiHelper.U32(p, "IsKeyManagementServiceMachine") == 1)
                {
                    kmsFlag = true;
                    WmiHelper.Invoke(svc, "SetKeyManagementServiceListeningPort",
                        new string[] { "Port" },
                        new object[] { uint.Parse(port) });
                    LineOut(Resources.Get("L_MsgKmsPortSet").Replace("%PORT%", port));
                    LineOut(Resources.Get("L_MsgWarningKmsReboot"));
                    break;
                }
            }

            if (!kmsFlag) LineOut(Resources.Get("L_MsgWarningKmsPort"));
        }

        private static void SetDnsPublishingDisabled(bool disable)
        {
            ManagementObject svc = WmiHelper.GetService("Version");
            bool kmsFlag = false;

            foreach (ManagementObject p in WmiHelper.GetProducts(
                "ID, IsKeyManagementServiceMachine", "PartialProductKey <> null"))
            {
                if (WmiHelper.U32(p, "IsKeyManagementServiceMachine") == 1)
                {
                    kmsFlag = true;
                    WmiHelper.Invoke(svc, "DisableKeyManagementServiceDnsPublishing",
                        new string[] { "Disable" }, new object[] { disable });

                    if (disable)
                        LineOut(Resources.Get("L_MsgKmsDnsPublishingDisabled"));
                    else
                        LineOut(Resources.Get("L_MsgKmsDnsPublishingEnabled"));

                    LineOut(Resources.Get("L_MsgWarningKmsReboot"));
                    break;
                }
            }

            if (!kmsFlag) LineOut(Resources.Get("L_MsgKmsDnsPublishingWarning"));
        }

        private static void SetKmsLowPriority(bool low)
        {
            ManagementObject svc = WmiHelper.GetService("Version");
            bool kmsFlag = false;

            foreach (ManagementObject p in WmiHelper.GetProducts(
                "ID, IsKeyManagementServiceMachine", "PartialProductKey <> null"))
            {
                if (WmiHelper.U32(p, "IsKeyManagementServiceMachine") == 1)
                {
                    kmsFlag = true;
                    WmiHelper.Invoke(svc, "EnableKeyManagementServiceLowPriority",
                        new string[] { "LowPriority" }, new object[] { low });

                    if (low)
                        LineOut(Resources.Get("L_MsgKmsPriSetToLow"));
                    else
                        LineOut(Resources.Get("L_MsgKmsPriSetToNormal"));

                    LineOut(Resources.Get("L_MsgWarningKmsReboot"));
                    break;
                }
            }

            if (!kmsFlag) LineOut(Resources.Get("L_MsgWarningKmsPri"));
        }

        private static void SetVLActivationType(int? type, string actId)
        {
            int t = type.HasValue ? type.Value : 0;
            if (t < 0 || t > 3)
            {
                LineOut(Resources.Get("L_MsgInvalidDataError"));
                return;
            }

            ManagementObject target = GetKmsTarget(actId);
            if (target == null) return;

            if (t != 0)
                WmiHelper.Invoke(target, "SetVLActivationTypeEnabled",
                    new string[] { "ActivationType" }, new object[] { (uint)t });
            else
                WmiHelper.Invoke(target, "ClearVLActivationTypeEnabled");

            LineOut(Resources.Get("L_MsgVLActivationTypeSet"));
        }

        // -----------------------------------------------------------------
        // AD activation (WMI-backed commands)
        // -----------------------------------------------------------------

        private static void FailRemoteExec()
        {
            if (_isRemote)
            {
                LineOut(Resources.Get("L_MsgRemoteExecNotSupported"));
                Exit(1);
            }
        }

        private static void AdActivateOnline(string key, string aoName)
        {
            FailRemoteExec();

            ManagementObject svc = WmiHelper.GetService("Version");
            WmiHelper.Invoke(svc, "DoActiveDirectoryOnlineActivation",
                new string[] { "ProductKey", "ActivationObjectName" },
                new object[] { key, aoName });

            LineOut(Resources.Get("L_MsgActivated"));
        }

        private static void AdGetIID(string key)
        {
            FailRemoteExec();

            ManagementObject svc = WmiHelper.GetService("Version");
            ManagementBaseObject r = WmiHelper.InvokeWithResult(svc,
                "GenerateActiveDirectoryOfflineActivationId",
                new string[] { "ProductKey" }, new object[] { key });

            LineOut(Resources.Get("L_MsgInstallationID") + (string)r["InstallationId"]);
            LineOut("");
            LineOut(Resources.Get("L_MsgPhoneNumbers"));
        }

        private static void AdActivatePhone(string key, string cid, string aoName)
        {
            FailRemoteExec();

            ManagementObject svc = WmiHelper.GetService("Version");
            WmiHelper.Invoke(svc, "DepositActiveDirectoryOfflineActivationConfirmation",
                new string[] { "ProductKey", "ConfirmationId", "ActivationObjectName" },
                new object[] { key, cid, aoName });

            LineOut(Resources.Get("L_MsgActivated"));
        }

        // -----------------------------------------------------------------
        // Registry helpers
        // -----------------------------------------------------------------

        private const string SlKeyPath =
            @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform";
        private const string SlKeyPath32 =
            @"SOFTWARE\Wow6432Node\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform";

        private static void SetKmsVersion(string ver)
        {
            try
            {
                using (RegistryKey k = Registry.LocalMachine.CreateSubKey(SlKeyPath))
                {
                    if (k != null)
                        k.SetValue("KeyManagementServiceVersion", ver, RegistryValueKind.String);
                }
            }
            catch { }

            try
            {
                using (RegistryKey k = Registry.LocalMachine.OpenSubKey(SlKeyPath32, true))
                {
                    if (k != null)
                        k.SetValue("KeyManagementServiceVersion", ver, RegistryValueKind.String);
                }
            }
            catch { }
        }

        private static void ClearKmsVersion()
        {
            try
            {
                using (RegistryKey k = Registry.LocalMachine.OpenSubKey(SlKeyPath, true))
                {
                    if (k != null)
                        k.DeleteValue("KeyManagementServiceVersion", false);
                }
            }
            catch { }

            try
            {
                using (RegistryKey k = Registry.LocalMachine.OpenSubKey(SlKeyPath32, true))
                {
                    if (k != null)
                        k.DeleteValue("KeyManagementServiceVersion", false);
                }
            }
            catch { }
        }
    }
}