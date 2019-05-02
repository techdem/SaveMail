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
        public void TestShowDialog()
        {
            Dictionary<object, object> showDialog = SaveMailUI.ShowDialog();
            Assert.IsInstanceOfType(showDialog, typeof(Dictionary<object, object>));
            
            //MessageBox.Show(showDialog["dialogResult"]);
            Assert.AreEqual(showDialog["dialogResult"], DialogResult.OK);
            Assert.IsTrue(showDialog["selectedPath"].Equals("C:\\"));
        }
    }
}
