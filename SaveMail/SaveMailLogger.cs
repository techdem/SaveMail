using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SaveMail
{
    public class SaveMailLogger
    {
        readonly string userProfilePath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        static string logFilePath;
        static StreamWriter sw;

        public SaveMailLogger(string applicationName)
        {
            logFilePath = userProfilePath + "\\AppData\\Local\\" + applicationName + "\\Log-" + Environment.MachineName + ".txt";
        }

        public static void CreateLogFile()
        {
            if (!File.Exists(logFilePath))
            {
                using (sw = File.CreateText(logFilePath))
                {
                    sw.WriteLine("LogFile for user: {0} on machine: {1}\r\n", Environment.UserName, Environment.MachineName);
                }
            }
        }

        public static void LogAction(string logMessage)
        {
            CreateLogFile();

            using (sw = File.AppendText(logFilePath))
            {
                sw.WriteLine(DateTime.Now + " :\t" + logMessage);
            }
        }
    }
}
