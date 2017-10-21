using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ActiveX.NET.Common
{
    public interface IActiveXServerAttribute
    {
        Type CoClassType { get; }
        Type PrimaryInterface { get; }
    }

    [MetadataAttribute]
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false)]
    public class ActiveXServerAttribute : ExportAttribute, IActiveXServerAttribute
    {
        public ActiveXServerAttribute() : base(typeof(IActiveXServer))
        {            
        }

        public Type CoClassType { get; set; }
        public Type PrimaryInterface { get; set; }
    }  
}
