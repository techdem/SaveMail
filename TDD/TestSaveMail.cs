using Microsoft.Office.Interop.Outlook;
using Moq;
using System;
using System.Diagnostics;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using System.Collections.Generic;

namespace UnitTestsForSaveMail
{
    [TestClass]
    public class TestSaveMail
    {
        static Microsoft.Office.Interop.Outlook.Application outlookApplication = new Microsoft.Office.Interop.Outlook.Application();
        Recipient outlookAddress = outlookApplication.Session.CreateRecipient("test@internal.address");
        readonly Dictionary<object, object> okResult = new Dictionary<object, object> { { "dialogResult", DialogResult.OK }, { "selectedPath", "C:\\TEST" } };
        readonly Dictionary<object, object> negativeResult = new Dictionary<object, object> { { "dialogResult", DialogResult.Cancel }, { "selectedPath", " " } };

        [TestMethod]
        public void TestGetPath()
        {
            Dictionary<object, object> getPath = SaveMail.SaveMailModel.GetPath(DialogResult.OK, "C:\\");

            Assert.AreEqual(getPath["dialogResult"], DialogResult.OK);
            Assert.IsTrue(getPath["selectedPath"].Equals("C:\\"));
        }

        [TestMethod]
        public void TestGetEmailOrigin()
        {
            MailItem internalMailItem = (MailItem)outlookApplication.CreateItem(OlItemType.olMailItem);
            internalMailItem.Sender = outlookAddress.AddressEntry;

            Assert.AreEqual("test@internal.address", SaveMail.SaveMailModel.GetEmailOrigin(internalMailItem));
        }

        [TestMethod]
        public void TestSaveSelected()
        {
            MailItem validMailItem = (MailItem)outlookApplication.CreateItem(OlItemType.olMailItem);
            validMailItem.Sender = outlookAddress.AddressEntry;
            validMailItem.Subject = "valid email subject";
            MailItem[] selectedItems = new MailItem[] { validMailItem };
            
            Assert.IsTrue(SaveMail.SaveMailController.SaveSelected(okResult, selectedItems).Equals("saveSuccess"));
            Assert.IsTrue(SaveMail.SaveMailController.SaveSelected(negativeResult, selectedItems).Equals("saveCancelled"));
            File.Delete("C:\\TEST\\4501-01-01 test@internal.address valid email subject.msg");
        }

        [TestMethod]
        public void TestPathCheck()
        {
            MailItem invalidEmailSubject = (MailItem)outlookApplication.CreateItem(OlItemType.olMailItem);
            invalidEmailSubject.Subject = "\\/:*?\"<>|";

            Assert.IsTrue(SaveMail.SaveMailModel.PathCheck(okResult, invalidEmailSubject).Equals("pathOK"));
            Assert.IsTrue(SaveMail.SaveMailModel.PathCheck(negativeResult, invalidEmailSubject).Equals("saveCancelled"));

        }

        [TestMethod]
        public void TestBlockRootPathCheck()
        {
            MailItem mailItem = (MailItem)outlookApplication.CreateItem(OlItemType.olMailItem);
            mailItem.Subject = "test";
            Dictionary<object, object> okResult = new Dictionary<object, object> { { "dialogResult", DialogResult.OK }, { "selectedPath", "C:\\" } };

            Assert.IsTrue(SaveMail.SaveMailModel.PathCheck(okResult, mailItem).Equals("pathInvalid"));
        }
    }
}
