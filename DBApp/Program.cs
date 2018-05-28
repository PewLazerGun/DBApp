using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DBApp
{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {
            //string dbPathMyDocs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            //string dbPath = Path.Combine(dbPathMyDocs, "FDB");
            //AppDomain.CurrentDomain.SetData("DataDirectory", dbPath);


            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
