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
        // Set application name
        readonly static string applicationName = "SaveMail";

        // Set the location for the application folder and log file
        readonly static string userProfilePath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        static string applicationFolderPath = userProfilePath + "\\AppData\\Local\\" + applicationName;
        static string logFilePath = applicationFolderPath + "\\Log-" + Environment.MachineName + ".txt";

        static StreamWriter sw;

        // Create new application folder
        public static void CreateApplicationFolder()
        {
            logFilePath = userProfilePath + "\\AppData\\Local\\" + applicationName + "\\Log-" + Environment.MachineName + ".txt";
            Directory.CreateDirectory(applicationFolderPath);
        }

        // Create new log file if it doesn't already exist
        public static void CreateLogFile()
        {
            CreateApplicationFolder();

            if (!File.Exists(logFilePath))
            {
                using (sw = File.CreateText(logFilePath))
                {
                    sw.WriteLine("LogFile for user: {0} on machine: {1}\r\n", Environment.UserName, Environment.MachineName);
                }
            }
        }

        // Log new action with timestamp
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
