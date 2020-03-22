using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GetMail
{
    class SimpleLogger
    {
        private static String LogFile = "\\" + DateTime.Now.ToString("dd-MM-yyyy") + "-" + DateTime.Now.ToString("hhmmss") + ".log";
 
        public static Boolean WriteLog(String pTexto) {

            StreamWriter sw = new StreamWriter(Application.StartupPath + Convert.ToString(LogFile), true);
            sw.WriteLine(DateTime.Now.ToString("dd/MM/yyyy") + " " + DateTime.Now.ToString("hh:mm:ss") + " - " + pTexto);
            sw.Close();
            return false;

        }
    }
}
