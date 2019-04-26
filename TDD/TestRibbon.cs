using Microsoft.Office.Interop.Outlook;
using System;
using System.Diagnostics;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;

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
            object[] okResult = new object[] { DialogResult.OK, "savePath" };
            object[] negativeResult = new object[] { DialogResult.Cancel, " " };
            MailItem validMailItem = (MailItem)outlookApplication.CreateItem(OlItemType.olMailItem);
            validMailItem.Subject = "valid email subject";
            validMailItem.Sender = outlookAddress.AddressEntry;
            MailItem[] selectedItems = new MailItem[] { validMailItem };
            
            Assert.IsTrue(SaveMail.SaveMail.SaveSelected(okResult, selectedItems));
            Assert.IsFalse(SaveMail.SaveMail.SaveSelected(negativeResult, selectedItems));
        }

        [TestMethod]
        public void TestPathCheck()
        {
            object[] mockOkResult = new object[] { DialogResult.OK, "savePath" };
            String invalidEmailSubject = "\\/:*?\"<>|";

            Assert.IsTrue(SaveMail.SaveMail.PathCheck(mockOkResult, invalidEmailSubject));
        }
    }
}
