using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SaveMail;

namespace UnitTests
{
    [TestClass]
    public class TestUI
    {
        [TestMethod]
        public void SaveMailUI()
        {
            SaveMailUI saveMailUI = new SaveMailUI();
            IDictionary<String,String> showDialog = saveMailUI.ShowDialog();
            Assert.IsInstanceOfType(showDialog, typeof(IDictionary<String,String>));
            Assert.AreEqual(showDialog["dialogResult"], "OK");
            Assert.AreEqual(showDialog["selectedPath"], "C:\\");
        }
    }
}
