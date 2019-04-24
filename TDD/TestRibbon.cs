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
            Microsoft.Office.Interop.Outlook.Application outlookApplication = new Microsoft.Office.Interop.Outlook.Application();
            Recipient outlookAddress = outlookApplication.Session.CreateRecipient("test@internal.address");
            MailItem internalMailItem = (MailItem)outlookApplication.CreateItem(OlItemType.olMailItem);
            internalMailItem.Sender = outlookAddress.AddressEntry;

            Assert.AreEqual("test@internal.address", SaveMail.SaveMail.GetEmailOrigin(internalMailItem));
        }

        [TestMethod]
        public void TestSaveSelected()
        {
            object[] mockOkResult = new object[] { DialogResult.OK, "savePath" };
            object[] mockNegativeResult = new object[] { DialogResult.OK, " " };

            Assert.IsTrue(SaveMail.SaveMail.SaveSelected(mockOkResult));
            Assert.IsFalse(SaveMail.SaveMail.SaveSelected(mockNegativeResult));
        }
    }
}
