using System;
using System.DirectoryServices;
using System.Runtime.InteropServices;

namespace slmgr
{
    /// <summary>
    /// Active Directory based activation helpers. Mirrors the AD commands of
    /// slmgr.vbs (/ao-list, /del-ao).
    /// </summary>
    internal static class AdHelper
    {
        private const string ADActObjContainer =
            "CN=Activation Objects,CN=Microsoft SPP,CN=Services,";
        private const string ADActObjClass = "msSPP-ActivationObject";
        private const string ADActObjAttribSkuId = "msSPP-CSVLKSkuId";
        private const string ADActObjAttribPid = "msSPP-CSVLKPid";
        private const string ADActObjAttribPartialPkey = "msSPP-CSVLKPartialProductKey";
        private const string ADActObjDisplayName = "displayName";
        private const string ADActObjAttribDN = "distinguishedName";

        private const uint HR_ERROR_DS_NO_SUCH_OBJECT = 0x80072030;

        /// <summary>Formats a byte as two uppercase hex digits.</summary>
        private static string HexByte(byte b)
        {
            return b.ToString("X2");
        }

        /// <summary>
        /// Converts a 16-byte GUID to the string form used by slmgr.vbs.
        /// Data1, Data2 and Data3 are little endian; Data4 is big endian.
        /// </summary>
        public static string GuidToString(byte[] b)
        {
            if (b.Length < 16)
                return BitConverter.ToString(b).Replace("-", "");

            return "{" +
                HexByte(b[3]) + HexByte(b[2]) + HexByte(b[1]) + HexByte(b[0]) + "-" +
                HexByte(b[5]) + HexByte(b[4]) + "-" +
                HexByte(b[7]) + HexByte(b[6]) + "-" +
                HexByte(b[8]) + HexByte(b[9]) + "-" +
                HexByte(b[10]) + HexByte(b[11]) + HexByte(b[12]) +
                HexByte(b[13]) + HexByte(b[14]) + HexByte(b[15]) +
                "}";
        }

        /// <summary>Returns the DNS name of the machine's domain plus a trailing slash.</summary>
        public static string GetMachineDomain()
        {
            Type t = Type.GetTypeFromProgID("ADSystemInfo");
            if (t == null) throw new InvalidOperationException("ADSystemInfo not available.");

            object info = Activator.CreateInstance(t);
            return info.GetType().InvokeMember(
                "DomainDNSName",
                System.Reflection.BindingFlags.GetProperty,
                null, info, null).ToString() + "/";
        }

        /// <summary>
        /// Lists all Activation Objects in the local forest. Output goes
        /// through the supplied lineOut callback.
        /// </summary>
        public static void ListActivationObjects(Action<string> lineOut)
        {
            try
            {
                string domain = GetMachineDomain();

                using (DirectoryEntry root = new DirectoryEntry("LDAP://" + domain + "rootDSE"))
                {
                    string config = (string)root.Properties["configurationNamingContext"][0];
                    string path = "LDAP://" + domain + ADActObjContainer + config;

                    using (DirectoryEntry container = new DirectoryEntry(path))
                    {
                        lineOut(Resources.Get("L_MsgActObjAvailable"));
                        bool found = false;

                        foreach (DirectoryEntry child in container.Children)
                        {
                            if (child.SchemaClassName != ADActObjClass) continue;
                            found = true;

                            object skuVal = child.Properties[ADActObjAttribSkuId].Value;
                            string skuId;
                            byte[] skuBytes = skuVal as byte[];
                            if (skuBytes != null)
                                skuId = GuidToString(skuBytes);
                            else
                                skuId = skuVal == null ? "" : skuVal.ToString();

                            lineOut("    " + Resources.Get("L_MsgADInfoAOName") +
                                    child.Properties[ADActObjDisplayName].Value);
                            lineOut("    " + "    " + Resources.Get("L_MsgActID") + skuId);
                            lineOut("    " + "    " + Resources.Get("L_MsgPartialPKey") +
                                    child.Properties[ADActObjAttribPartialPkey].Value);
                            lineOut("    " + "    " + Resources.Get("L_MsgADInfoExtendedPid") +
                                    child.Properties[ADActObjAttribPid].Value);
                            lineOut("    " + "    " + Resources.Get("L_MsgADInfoAODN") +
                                    child.Properties[ADActObjAttribDN].Value);
                            lineOut("");
                        }

                        if (!found)
                            lineOut("    " + Resources.Get("L_MsgActObjNoneFound"));
                    }
                }
            }
            catch (COMException ex)
            {
                if ((uint)ex.HResult == HR_ERROR_DS_NO_SUCH_OBJECT)
                    lineOut(Resources.Get("L_MsgADSchemaNotSupported"));
                else
                    throw;
            }
        }

        /// <summary>
        /// Deletes an Activation Object given either its RDN or full DN.
        /// </summary>
        public static void DeleteActivationObject(Action<string> lineOut, string name)
        {
            try
            {
                string domain = GetMachineDomain();

                using (DirectoryEntry root = new DirectoryEntry("LDAP://" + domain + "rootDSE"))
                {
                    string config = (string)root.Properties["configurationNamingContext"][0];

                    string dn;
                    if (name.IndexOf(",cn=", StringComparison.OrdinalIgnoreCase) > 0)
                        dn = name;
                    else if (name.StartsWith("cn=", StringComparison.OrdinalIgnoreCase))
                        dn = name + "," + ADActObjContainer + config;
                    else
                        dn = "CN=" + name + "," + ADActObjContainer + config;

                    using (DirectoryEntry obj = new DirectoryEntry("LDAP://" + dn))
                    {
                        object parentObj = obj.Parent;
                        DirectoryEntry parent = parentObj as DirectoryEntry;
                        if (parent != null)
                            parent.Children.Remove(obj);
                        else
                            obj.DeleteTree();
                    }

                    lineOut(Resources.Get("L_MsgSucess"));
                }
            }
            catch (COMException ex)
            {
                if ((uint)ex.HResult == HR_ERROR_DS_NO_SUCH_OBJECT)
                    lineOut(Resources.Get("L_MsgADSchemaNotSupported"));
                else
                    throw;
            }
        }
    }
}