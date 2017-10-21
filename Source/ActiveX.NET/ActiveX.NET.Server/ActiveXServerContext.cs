using ActiveX.NET.Server.Utility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ActiveX.NET.Server
{
    internal class ActiveXServerContext : ApplicationContext
    {
        private object syncRoot = new Object(); // For thread-sync in lock
        private bool _bRunning = false; // Whether the server is running
        // The lock count (the number of active COM objects) in the server
        private int _nLockCnt = 0;
        // The timer to trigger GC every 5 seconds
        private System.Threading.Timer _gcTimer;

        private List<uint> _registrationCookies = new List<uint>();

        private ActiveXServers _comServers;


        public ActiveXServerContext()
        {
            _comServers = new ActiveXServers(this); 
        }

        /// <summary>
        /// The method is call every 5 seconds to GC the managed heap after 
        /// the COM server is started.
        /// </summary>
        /// <param name="stateInfo"></param>
        private static void GarbageCollect(object stateInfo)
        {
            GC.Collect();   // GC
        }

        public bool RegisterPlugins(bool unregister = false)
        {
            var assemblies = _comServers.GetServers().Select(server => server.CoClassType.Assembly).Distinct();
            
            var regAsm = new RegistrationServices();

            if (unregister)
            {
                assemblies.TryForEach(assembly =>
                {
                    regAsm.UnregisterAssembly(assembly);
                });
            }
            else
            {
                assemblies.TryForEach(assembly =>
                {
                    regAsm.RegisterAssembly(assembly, AssemblyRegistrationFlags.SetCodeBase);
                });
            }

            return true;
        }

        private bool PreMessageLoop()
        {
            int hResult = 0;

             _comServers.GetServers().TryForEach(comServer =>
            {
                uint comRegCookie;

                Guid clsidSimpleObj = comServer.ClassID;

                // Register the SimpleObject class object
                hResult = ActiveXNative.CoRegisterClassObject(
                    ref clsidSimpleObj,                 // CLSID to be registered
                    new DefaultActiveXFactory(comServer),     // Class factory
                    CLSCTX.LOCAL_SERVER,                // Context to run
                    REGCLS.MULTIPLEUSE | REGCLS.SUSPENDED,
                    out comRegCookie);

                if (hResult != 0)
                {
                    throw new ApplicationException("CoRegisterClassObject failed w/err 0x" + hResult.ToString("X"));
                }
                else
                {
                    _registrationCookies.Add(comRegCookie);
                }
            });

            // Inform the SCM about all the registered classes, and begins 
            // letting activation requests into the server process.
            hResult = ActiveXNative.CoResumeClassObjects();
            if (hResult != 0)
            {
                RevokeRegistrations();
                throw new ApplicationException("CoResumeClassObjects failed w/err 0x" + hResult.ToString("X"));
            }

            // Records the count of the active COM objects in the server. 
            // When _nLockCnt drops to zero, the server can be shut down.
            _nLockCnt = 0;

            // Start the GC timer to trigger GC every 5 seconds.
            _gcTimer = new System.Threading.Timer(new TimerCallback(GarbageCollect), null, 5000, 5000);
            return false;
        }

        private void RevokeRegistrations()
        {
            // Revoke the registration of the COM classes.
            _registrationCookies.TryForEach(cookie =>
            {
                ActiveXNative.CoRevokeClassObject(cookie);
            });
        }

        private bool PostMessageLoop()
        {
            RevokeRegistrations();

            // Dispose the GC timer.
            if (_gcTimer != null)
            {
                _gcTimer.Dispose();
            }

            // Wait for any threads to finish.
            Thread.Sleep(1000);

            return true;
        }

        private bool RunMessageLoop()
        {
            Application.Run(this);
            return true;
        }

        public void Run()
        {            
            lock (syncRoot) // Ensure thread-safe
            {
                // If the server is running, return directly.
                if (_bRunning) return;

                // Indicate that the server is running now.
                _bRunning = true;
            }
            try
            {
                // Call PreMessageLoop to initialize the member variables 
                // and register the class factories.
                PreMessageLoop();
                try
                {
                    RunMessageLoop();
                }
                finally
                {
                    // Call PostMessageLoop to revoke the registration.
                    PostMessageLoop();
                }
            }
            finally
            {
                _bRunning = false;
            }
        }

        /// <summary>
        /// Increase the lock count
        /// </summary>
        /// <returns>The new lock count after the increment</returns>
        /// <remarks>The method is thread-safe.</remarks>
        public int Lock()
        {
            return Interlocked.Increment(ref _nLockCnt);
        }

        /// <summary>
        /// Decrease the lock count. When the lock count drops to zero, post 
        /// the WM_QUIT message to the message loop in the main thread to 
        /// shut down the COM server.
        /// </summary>
        /// <returns>The new lock count after the increment</returns>
        public int Unlock()
        {
            int nRet = Interlocked.Decrement(ref _nLockCnt);

            // If lock drops to zero, attempt to terminate the server.
            if (nRet == 0)
            {
                ExitThread();
            }

            return nRet;
        }

        /// <summary>
        /// Get the current lock count.
        /// </summary>
        /// <returns></returns>
        public int GetLockCount()
        {
            return _nLockCnt;
        }
    };
};
