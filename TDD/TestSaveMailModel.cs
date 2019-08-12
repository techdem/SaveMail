using Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;

namespace UnitTestsForSaveMail
{
    [TestClass]
    public class TestSaveMailModel
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
        public void GetRecentLocation()
        {
            SaveMail.SaveMailLogger.LogAction("Saved to location: testSaveLocation");
            SaveMail.SaveMailLogger.LogAction("Total saved: testOutput");

            Assert.IsTrue(SaveMail.SaveMailModel.GetRecentLocation().Equals("testSaveLocation"));
        }

        [TestMethod]
        public void TestGetEmailAddress()
        {
            MailItem incominglMailItem = (MailItem)outlookApplication.CreateItem(OlItemType.olMailItem);
            incominglMailItem.Sender = outlookAddress.AddressEntry;

            Assert.AreEqual("test@internal.address", SaveMail.SaveMailModel.GetEmailAddress(incominglMailItem, "incoming"));
        }

        [TestMethod]
        public void TestPathCheck()
        {
            MailItem invalidEmailSubject = (MailItem)outlookApplication.CreateItem(OlItemType.olMailItem);
            invalidEmailSubject.Subject = "\\/:*?\"<>|";
            MailItem emptyEmailSubject = (MailItem)outlookApplication.CreateItem(OlItemType.olMailItem);
            emptyEmailSubject.Subject = "";

            Assert.IsTrue(!SaveMail.SaveMailModel.PathCheck(okResult, invalidEmailSubject).Equals("pathInvalid") && !SaveMail.SaveMailModel.PathCheck(okResult, invalidEmailSubject).Equals("saveCancelled"));
            Assert.IsTrue(!SaveMail.SaveMailModel.PathCheck(okResult, emptyEmailSubject).Equals("pathInvalid") && !SaveMail.SaveMailModel.PathCheck(okResult, emptyEmailSubject).Equals("saveCancelled"));
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
