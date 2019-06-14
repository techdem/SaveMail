using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SaveMail
{
    public class SaveMailLogging
    {
        readonly string userProfilePath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        string logFilePath;
        StreamWriter sw;

        public SaveMailLogging(string applicationName)
        {
            this.logFilePath = userProfilePath + "\\AppData\\Local\\" + applicationName + "\\Log-" + Environment.MachineName + ".txt";
        }

        public void CreateLogFile()
        {
            if (!File.Exists(logFilePath))
            {
                using (sw = File.CreateText(logFilePath))
                {
                    sw.WriteLine("LogFile for user: {0} on machine: {1}\r\n", Environment.UserName, Environment.MachineName);
                }
            }
        }

        public void LogAction(string logMessage)
        {
            CreateLogFile();

            using (sw = File.AppendText(logFilePath))
            {
                sw.WriteLine(logMessage);
            }
        }
    }
}
