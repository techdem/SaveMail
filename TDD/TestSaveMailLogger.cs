using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;

namespace UnitTestsForSaveMail
{
    [TestClass]
    public class TestSaveMailLogging
    {
        readonly static string applicationName = "SaveMail";
        readonly static string userProfilePath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        static string applicationFolderPath = userProfilePath + "\\AppData\\Local\\" + applicationName;
        readonly string logFilePath = applicationFolderPath + "\\Log-" + Environment.MachineName + ".txt";

        [TestMethod]
        public void TestCreateApplicationFolder()
        {
            SaveMail.SaveMailLogger.CreateLogFile();

            Assert.IsTrue(Directory.Exists(applicationFolderPath));
        }

        [TestMethod]
        public void TestCreateLogFile()
        {
            SaveMail.SaveMailLogger.CreateLogFile();

            Assert.IsTrue(File.Exists(logFilePath));
        }

        [TestMethod]
        public void TestLogAction()
        {
            string logFileContent;
            SaveMail.SaveMailLogger.LogAction("TestLogAction");

            using (StreamReader sr = new StreamReader(logFilePath))
            {
                logFileContent = sr.ReadToEnd();
            }

            Assert.IsTrue(logFileContent.Contains("TestLogAction"));
        }
    }
}
