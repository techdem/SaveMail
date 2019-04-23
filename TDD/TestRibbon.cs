using System;
using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SaveMail;

namespace TDD
{
    [TestClass]
    public class MockRibbon
    {
        [TestMethod]
        public void TestBrowsers()
        {
            SaveMail.MockRibbon mockRibbon = new SaveMail.MockRibbon();

            Assert.IsInstanceOfType(mockRibbon.TestBrowser(), typeof(System.Windows.Forms.FolderBrowserDialog));
        }
    }
}
