﻿using Microsoft.Office.Interop.Outlook;
using Moq;
using System;
using System.Diagnostics;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using System.Collections.Generic;

namespace UnitTests
{
    [TestClass]
    public class TestSaveMail
    {
        static Microsoft.Office.Interop.Outlook.Application outlookApplication = new Microsoft.Office.Interop.Outlook.Application();
        Recipient outlookAddress = outlookApplication.Session.CreateRecipient("test@internal.address");
        Dictionary<object, object> okResult = new Dictionary<object, object> { { "dialogResult", DialogResult.OK }, { "selectedPath", "C:\\TEST" } };
        Dictionary<object, object> negativeResult = new Dictionary<object, object> { { "dialogResult", DialogResult.Cancel }, { "selectedPath", " " } };

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
            MailItem validMailItem = (MailItem)outlookApplication.CreateItem(OlItemType.olMailItem);
            validMailItem.Sender = outlookAddress.AddressEntry;
            validMailItem.Subject = "valid email subject";
            MailItem[] selectedItems = new MailItem[] { validMailItem };
            
            Assert.IsTrue(SaveMail.SaveMail.SaveSelected(okResult, selectedItems));
            Assert.IsFalse(SaveMail.SaveMail.SaveSelected(negativeResult, selectedItems));
            File.Delete("C:\\TEST\\01-01-4501 test@internal.address valid email subject.msg");
        }

        [TestMethod]
        public void TestPathCheck()
        {
            MailItem invalidEmailSubject = (MailItem)outlookApplication.CreateItem(OlItemType.olMailItem);
            invalidEmailSubject.Subject = "\\/:*?\"<>|";

            Assert.IsTrue(SaveMail.SaveMail.PathCheck(okResult, invalidEmailSubject));
            Assert.IsFalse(SaveMail.SaveMail.PathCheck(negativeResult, invalidEmailSubject));

        }

        [TestMethod]
        public void TestBlockRootPathCheck()
        {
            MailItem mailItem = (MailItem)outlookApplication.CreateItem(OlItemType.olMailItem);
            mailItem.Subject = "test";
            Dictionary<object, object> okResult = new Dictionary<object, object> { { "dialogResult", DialogResult.OK }, { "selectedPath", "C:\\" } };

            Assert.IsFalse(SaveMail.SaveMail.PathCheck(okResult, mailItem));
        }
    }
}
