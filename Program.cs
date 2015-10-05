using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace UploadApplication
{
    static class Program
    {
        public static string strUserName = "";
        public static string UserId = "0";
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new frmlogin());
        }
    }
}
