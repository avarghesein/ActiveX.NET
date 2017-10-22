using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ActiveX.NET.Common
{
    public interface IActiveXServer
    {
        Func<int> LockActiveXServer { get; set; }

        Func<int> UnLockActiveXServer { get; set; }
    };
};
