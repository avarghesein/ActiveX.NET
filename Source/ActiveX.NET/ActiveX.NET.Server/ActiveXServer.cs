using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ActiveX.NET.Server
{
    internal class ActiveXServer
    {
        public Guid ClassID { get; set; }
        public Type PrimaryInterface { get; set; }
        public Type CoClassType { get; set; }
    };
};
