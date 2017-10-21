using ActiveX.NET.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ActiveX.NET.Plugin
{
    [ClassInterface(ClassInterfaceType.None)]           // No ClassInterface
    [ComSourceInterfaces(typeof(ISimpleObjectEvents))]
    [Guid(SimpleObject.ClassId), ComVisible(true)]
    [ActiveXServer(CoClassType = typeof(SimpleObject), PrimaryInterface = typeof(ISimpleObject))]
    public class SimpleObject : ActiveXServerBase, ISimpleObject
    {
        #region COM Component Registration

        internal const string ClassId =
            "DB9935C1-19C5-4ed2-ADD2-9A57E19F53A3";
        internal const string InterfaceId =
            "941D219B-7601-4375-B68A-61E23A4C8425";
        internal const string EventsId =
            "014C067E-660D-4d20-9952-CD973CE50436";

   
        #endregion

        #region Properties

        private float fField = 0;

        public float FloatProperty
        {
            get { return this.fField; }
            set
            {
                bool cancel = false;
                // Raise the event FloatPropertyChanging
                if (null != FloatPropertyChanging)
                    FloatPropertyChanging(value, ref cancel);
                if (!cancel)
                    this.fField = value;
            }
        }

        #endregion

        #region Methods

        public string HelloWorld()
        {
            return "HelloWorld";
        }

        public void GetProcessThreadID(out uint processId, out uint threadId)
        {
            processId = 10;
            threadId = 10;
        }

        #endregion

        #region Events

        [ComVisible(false)]
        public delegate void FloatPropertyChangingEventHandler(float NewValue, ref bool Cancel);
        public event FloatPropertyChangingEventHandler FloatPropertyChanging;

        #endregion
    }
};
