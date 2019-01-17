using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ActiveX.NET.Common
{
    internal static class RegistrationUtility
    {
        private static void GuardNullType(Type t, String param)
        {
            if (t == null)
            {
                throw new ArgumentException("The CLR type must be specified.", param);
            }
        }

        private static bool ValidateActiveXServerAttribute(Type t)
        {
            return t.GetCustomAttributes<ActiveXServerAttribute>().Any();
        }

        /// <summary>
        /// Register the component as a local server.
        /// </summary>
        /// <param name="t"></param>
        internal static void RegisterLocalServer(Type t)
        {
            GuardNullType(t, "t");  // Check the argument

            if (!ValidateActiveXServerAttribute(t)) return;

            // Open the CLSID key of the component.
            using (RegistryKey keyCLSID = Registry.ClassesRoot.OpenSubKey(
                @"CLSID\" + t.GUID.ToString("B"), /*writable*/true))
            {
                // Remove the auto-generated InprocServer32 key after registration
                // (REGASM puts it there but we are going out-of-proc).
                keyCLSID.DeleteSubKeyTree("InprocServer32");

                // Create "LocalServer32" under the CLSID key
                using (RegistryKey subkey = keyCLSID.CreateSubKey("LocalServer32"))
                {
                    subkey.SetValue("", Assembly.GetEntryAssembly().Location, RegistryValueKind.String);
                }
            }
        }

        /// <summary>
        /// Unregister the component.
        /// </summary>
        /// <param name="t"></param>
        internal static void UnregisterLocalServer(Type t)
        {
            GuardNullType(t, "t");  // Check the argument

            if (!ValidateActiveXServerAttribute(t)) return;

            // Delete the CLSID key of the component
            Registry.ClassesRoot.DeleteSubKeyTree(@"CLSID\" + t.GUID.ToString("B"));
        }
    };
};

