using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;

namespace UnitTestsForSaveMail
{
    [TestClass]
    public class TestSaveMailLogger
    {
        readonly string logFile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\AppData\\Local\\SaveMail\\Log-" + Environment.MachineName + ".txt";
        SaveMail.SaveMailLogging sml = new SaveMail.SaveMailLogging("SaveMail");

        [TestMethod]
        public void TestCreateLogFile()
        {
            sml.CreateLogFile();

            Assert.IsTrue(File.Exists(logFile));
        }

        [TestMethod]
        public void TestLogAction()
        {
            string logFileContent;
            sml.LogAction("IT'S FRIDAY !!!");

            using (StreamReader sr = new StreamReader(logFile))
            {
                logFileContent = sr.ReadToEnd();
            }

            Assert.IsTrue(logFileContent.Contains("IT'S FRIDAY !!!"));
        }
    }
}
