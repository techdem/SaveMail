using System;
using System.Diagnostics;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace UnitTests
{
    [TestClass]
    public class TestSaveMail
    {
        [TestMethod]
        public void TestGetPath()
        {
            object[] testGetPath = SaveMail.SaveMail.GetPath(new FolderBrowserDialog());
            Assert.AreEqual(DialogResult.OK, testGetPath[0]);
            Assert.IsFalse(string.IsNullOrWhiteSpace((String)testGetPath[1]));
        }

        [TestMethod]
        public void TestSaveSelected()
        {
            object[] testSaveSelected = new object[] { DialogResult.OK, "savePath" };
            Assert.IsTrue(SaveMail.SaveMail.SaveSelected(testSaveSelected));
        }
    }
}
