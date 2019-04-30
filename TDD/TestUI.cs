using System;
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
            
            Assert.IsInstanceOfType(saveMailUI.GetPath(), typeof(object));
        }
    }
}
