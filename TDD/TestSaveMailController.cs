using Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using System.Collections.Generic;

namespace UnitTestsForSaveMail
{
    [TestClass]
    public class TestSaveMailController
    {
        static Microsoft.Office.Interop.Outlook.Application outlookApplication = new Microsoft.Office.Interop.Outlook.Application();
        Recipient outlookAddress = outlookApplication.Session.CreateRecipient("test@internal.address");
        readonly Dictionary<object, object> okResult = new Dictionary<object, object> { { "dialogResult", DialogResult.OK }, { "selectedPath", "C:\\TEST" } };
        readonly Dictionary<object, object> negativeResult = new Dictionary<object, object> { { "dialogResult", DialogResult.Cancel }, { "selectedPath", " " } };

        [TestMethod]
        public void TestSaveSelected()
        {
            MailItem validMailItem = (MailItem)outlookApplication.CreateItem(OlItemType.olMailItem);
            validMailItem.Sender = outlookAddress.AddressEntry;
            validMailItem.Subject = "valid email subject";
            List<object> selectedItems = new List<object> { validMailItem };
            
            Assert.IsTrue(SaveMail.SaveMailController.SaveSelected(okResult, selectedItems).Equals("saveSuccess"));
            Assert.IsTrue(SaveMail.SaveMailController.SaveSelected(negativeResult, selectedItems).Equals("saveCancelled"));
            File.Delete("C:\\TEST\\4501-01-01 test@internal.address valid email subject.msg");
        }
    }
}
