using ActiveX.NET.Server.Utility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ActiveX.NET.Server
{
    internal class DefaultActiveXFactory : IClassFactory
    {
        private ActiveXServer _server;
        public DefaultActiveXFactory(ActiveXServer server)
        {
            _server = server;
        }
        public int CreateInstance(IntPtr pUnkOuter, ref Guid riid,
            out IntPtr ppvObject)
        {
            ppvObject = IntPtr.Zero;

            if (pUnkOuter != IntPtr.Zero)
            {
                // The pUnkOuter parameter was non-NULL and the object does 
                // not support aggregation.
                Marshal.ThrowExceptionForHR(ActiveXNative.CLASS_E_NOAGGREGATION);
            }

            if (riid == _server.ClassID ||
                riid == new Guid(ActiveXNative.IID_IDispatch) ||
                riid == new Guid(ActiveXNative.IID_IUnknown))
            {
                // Create the instance of the .NET object
                ppvObject = Marshal.GetComInterfaceForObject(Activator.CreateInstance(_server.CoClassType), _server.PrimaryInterface);
            }
            else
            {
                // The object that ppvObject points to does not support the 
                // interface identified by riid.
                Marshal.ThrowExceptionForHR(ActiveXNative.E_NOINTERFACE);
            }

            return 0;   // S_OK
        }

        public int LockServer(bool fLock)
        {
            return 0;   // S_OK
        }
    }
};
