using Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Collections.Generic;

namespace UnitTestsForSaveMail
{
    [TestClass]
    public class TestSaveMailController
    {
        static Microsoft.Office.Interop.Outlook.Application outlookApplication = new Microsoft.Office.Interop.Outlook.Application();
        Recipient outlookAddress = outlookApplication.Session.CreateRecipient("test@internal.address");
        static string savePath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\AppData\\Local\\SaveMail";
        readonly Dictionary<object, object> okResult = new Dictionary<object, object> { { "dialogResult", DialogResult.OK }, { "selectedPath", savePath } };
        readonly Dictionary<object, object> negativeResult = new Dictionary<object, object> { { "dialogResult", DialogResult.Cancel }, { "selectedPath", " " } };

        [TestMethod]
        public void TestSaveSelected()
        {
            MailItem validMailItem = (MailItem)outlookApplication.CreateItem(OlItemType.olMailItem);
            validMailItem.Sender = outlookAddress.AddressEntry;
            validMailItem.Subject = "valid email subject";
            List<object> selectedItems = new List<object> { validMailItem };
            Directory.CreateDirectory(savePath);

            Assert.IsTrue(SaveMail.SaveMailController.SaveSelected(okResult, selectedItems).Equals("saveSuccess"));
            Assert.IsTrue(File.Exists(savePath + "\\4501-01-01 test@internal.address valid email subject.msg"));
            Assert.IsTrue(SaveMail.SaveMailController.SaveSelected(negativeResult, selectedItems).Equals("saveCancelled"));
            File.Delete(savePath + "\\4501-01-01 test@internal.address valid email subject.msg");
        }
    }
}
