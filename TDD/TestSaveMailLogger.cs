using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;

namespace UnitTestsForSaveMail
{
    [TestClass]
    public class TestSaveMailLogging
    {
        readonly string logFile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\AppData\\Local\\SaveMail\\Log-" + Environment.MachineName + ".txt";

        [TestMethod]
        public void TestCreateLogFile()
        {
            new SaveMail.SaveMailLogger("SaveMail");
            SaveMail.SaveMailLogger.CreateLogFile();

            Assert.IsTrue(File.Exists(logFile));
        }

        [TestMethod]
        public void TestLogAction()
        {
            new SaveMail.SaveMailLogger("SaveMail");
            string logFileContent;
            SaveMail.SaveMailLogger.LogAction("TestLogAction");

            using (StreamReader sr = new StreamReader(logFile))
            {
                logFileContent = sr.ReadToEnd();
            }

            Assert.IsTrue(logFileContent.Contains("TestLogAction"));
        }
    }
}
