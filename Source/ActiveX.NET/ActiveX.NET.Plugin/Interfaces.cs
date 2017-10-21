using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ActiveX.NET.Plugin
{
    [Guid(SimpleObject.InterfaceId), ComVisible(true)]
    public interface ISimpleObject
    {
        #region Properties

        float FloatProperty { get; set; }

        #endregion

        #region Methods

        string HelloWorld();

        void GetProcessThreadID(out uint processId, out uint threadId);

        #endregion
    }

    [Guid(SimpleObject.EventsId), ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ISimpleObjectEvents
    {
        #region Events

        [DispId(1)]
        void FloatPropertyChanging(float NewValue, ref bool Cancel);

        #endregion
    }
};
