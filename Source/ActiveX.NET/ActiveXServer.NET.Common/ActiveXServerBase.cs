using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.Composition;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ActiveX.NET.Common
{
    public abstract class ActiveXServerBase : IActiveXServer
    {
        [Import("LockActiveXServer",typeof(Func<int>))]
        protected Func<int> LockActiveXServerDelegate;

        [Import("UnLockActiveXServer", typeof(Func<int>))]
        protected Func<int> UnLockActiveXServerDelegate;

        public ActiveXServerBase()
        {
            // Increment the lock count of objects in the COM server.
            if (LockActiveXServerDelegate != null)
            {
                LockActiveXServerDelegate();
            }
        }

         ~ActiveXServerBase()
        {
            // Decrement the lock count of objects in the COM server.
            if (UnLockActiveXServerDelegate != null)
            {
                UnLockActiveXServerDelegate();
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
         [EditorBrowsable(EditorBrowsableState.Never)]
         [ComRegisterFunction()]
         public static void RegasmRegisterLocalServer(Type t)
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
         [EditorBrowsable(EditorBrowsableState.Never)]
         [ComUnregisterFunction()]
         public static void RegasmUnregisterLocalServer(Type t)
         {
             GuardNullType(t, "t");  // Check the argument

             if (!ValidateActiveXServerAttribute(t)) return;

             // Delete the CLSID key of the component
             Registry.ClassesRoot.DeleteSubKeyTree(@"CLSID\" + t.GUID.ToString("B"));
         }

         private static void GuardNullType(Type t, String param)
         {
             if (t == null)
             {
                 throw new ArgumentException("The CLR type must be specified.", param);
             }
         }
    };
};
