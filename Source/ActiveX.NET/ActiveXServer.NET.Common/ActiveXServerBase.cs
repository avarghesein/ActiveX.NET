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
    public abstract class ActiveXServerBase : StandardOleMarshalObject, IActiveXServer
    {
        [Import("LockActiveXServer", typeof(Func<int>))]
        public Func<int> LockActiveXServer { get; set; }

        [Import("UnLockActiveXServer", typeof(Func<int>))]
        public Func<int> UnLockActiveXServer { get; set; }

        public ActiveXServerBase()
        {
            // Increment the lock count of objects in the COM server. (This will not work, We've to call it manually)
            // MEF will resolve dependencies only after creation (i.e. After Constructor Calls)
            // i.e LockActiveXServer, will be injected only after Constructor invocation
            /*if (LockActiveXServer != null)
            {
                LockActiveXServer();
            }*/
        }

        ~ActiveXServerBase()
        {
            // Decrement the lock count of objects in the COM server.
            if (UnLockActiveXServer != null)
            {
                UnLockActiveXServer();
            }
        }

         /// <summary>
         /// Register the component as a local server.
         /// </summary>
         /// <param name="t"></param>
         [EditorBrowsable(EditorBrowsableState.Never)]
         [ComRegisterFunction()]
         public static void RegasmRegisterLocalServer(Type t)
         {
             RegistrationUtility.RegisterLocalServer(t);
         }

         /// <summary>
         /// Unregister the component.
         /// </summary>
         /// <param name="t"></param>
         [EditorBrowsable(EditorBrowsableState.Never)]
         [ComUnregisterFunction()]
         public static void RegasmUnregisterLocalServer(Type t)
         {
             RegistrationUtility.UnregisterLocalServer(t);
         }
    };
};
