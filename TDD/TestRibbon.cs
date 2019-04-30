using Microsoft.Office.Interop.Outlook;
using System;
using System.Diagnostics;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;

namespace UnitTests
{
    [TestClass]
    public class TestSaveMail
    {
        static Microsoft.Office.Interop.Outlook.Application outlookApplication = new Microsoft.Office.Interop.Outlook.Application();
        Recipient outlookAddress = outlookApplication.Session.CreateRecipient("test@internal.address");

        [TestMethod]
        public void TestGetPath()
        {
            object[] testGetPath = SaveMail.SaveMail.GetPath(new FolderBrowserDialog());

            Assert.AreEqual(DialogResult.OK, testGetPath[0]);
            Assert.IsFalse(string.IsNullOrWhiteSpace((String)testGetPath[1]));
        }

        [TestMethod]
        public void TestGetEmailOrigin()
        {
            MailItem internalMailItem = (MailItem)outlookApplication.CreateItem(OlItemType.olMailItem);
            internalMailItem.Sender = outlookAddress.AddressEntry;

            Assert.AreEqual("test@internal.address", SaveMail.SaveMail.GetEmailOrigin(internalMailItem));
        }

        [TestMethod]
        public void TestSaveSelected()
        {
            object[] okResult = new object[] { DialogResult.OK, "C:\\" };
            object[] negativeResult = new object[] { DialogResult.Cancel, " " };
            MailItem validMailItem = (MailItem)outlookApplication.CreateItem(OlItemType.olMailItem);
            validMailItem.Sender = outlookAddress.AddressEntry;
            validMailItem.Subject = "valid email subject";
            MailItem[] selectedItems = new MailItem[] { validMailItem };
            
            Assert.IsTrue(SaveMail.SaveMail.SaveSelected(okResult, selectedItems));
            Assert.IsFalse(SaveMail.SaveMail.SaveSelected(negativeResult, selectedItems));
            File.Delete("C:\\01-01-4501 test@internal.address valid email subject.msg");
        }

        [TestMethod]
        public void TestPathCheck()
        {
            object[] okResult = new object[] { DialogResult.OK, "valid path" };
            object[] negativeResult = new object[] { DialogResult.Cancel, " " };
            MailItem invalidEmailSubject = (MailItem)outlookApplication.CreateItem(OlItemType.olMailItem);
            invalidEmailSubject.Subject = "\\/:*?\"<>|";

            Assert.IsTrue(SaveMail.SaveMail.PathCheck(okResult, invalidEmailSubject));
            Assert.IsFalse(SaveMail.SaveMail.PathCheck(negativeResult, invalidEmailSubject));

        }

        [TestMethod]
        public void TestBlockRootPathCheck()
        {
            object[] okResult = new object[] { DialogResult.OK, "C:\\" };
            MailItem mailItem = (MailItem)outlookApplication.CreateItem(OlItemType.olMailItem);
            mailItem.Subject = "test";

            Assert.IsFalse(SaveMail.SaveMail.PathCheck(okResult, mailItem));
        }
    }
}
