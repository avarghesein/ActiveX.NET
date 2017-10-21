using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ActiveX.NET.Server
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] cmdLine)
        {
            ActiveXServerContext activeXContext = new ActiveXServerContext();

            if(cmdLine != null && cmdLine.Length > 0)
            {
                bool isRegister = false;
                //Condition Order is important
                if( cmdLine.Contains("/unregserver") || (isRegister = cmdLine.Contains("/regserver")))
                {                    
                    activeXContext.RegisterPlugins(!isRegister);
                    return;
                }
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            activeXContext.Run();
        }
    }
}
