using System;
using System.Diagnostics;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SaveMail;

namespace UnitTests
{
    [TestClass]
    public class TestMainRibbon
    {
        [TestMethod]
        public void TestGetPath()
        {
            object[] testGetPath = MainRibbon.GetPath(new FolderBrowserDialog());
            Assert.AreEqual(DialogResult.OK, testGetPath[0]);
            Assert.IsFalse(string.IsNullOrWhiteSpace((String)testGetPath[1]));
        }
    }
}
