﻿using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace SaveMail
{
    public static class SaveMailModel
    {
        readonly static string applicationName = "SaveMail";
        readonly static string logFile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\AppData\\Local\\" + applicationName + "\\Log-" + Environment.MachineName + ".txt";

        // Retrieve a list of selected emails
        public static List<object> GetSelectedEmails()
        {
            Selection selectedEmails = new Microsoft.Office.Interop.Outlook.Application().ActiveExplorer().Selection;
            List<object> emailItems = new List<object>();

            foreach (object item in selectedEmails)
            {
                if (!(item is MailItem))
                {
                    continue;
                }
                emailItems.Add(item);
            }

            Marshal.ReleaseComObject(selectedEmails);

            return emailItems;
        }

        // Parse the result from the folder browser dialog
        public static Dictionary<object, object> GetPath(DialogResult dialogResult, String selectedPath)
        {
            return new Dictionary<object, object>() { { "dialogResult", dialogResult },
                { "selectedPath", selectedPath } };
        }

        // Parse the user config file and determine the most recently used save location

        public static String GetRecentLocation()
        {
            String lastUsedLocation = "";

            using (StreamReader sr = new StreamReader(logFile))
            {
                String configFile = sr.ReadToEnd();

                int lastLocation = configFile.LastIndexOf("Saved to location:") + 19; // string start
                int lastConfirmation = configFile.LastIndexOf("Total saved:") - 24; // string end
                int pathLength = lastConfirmation - lastLocation;

                if (pathLength > 0)
                {
                    lastUsedLocation = configFile.Substring(lastLocation, pathLength);
                }
            }

            return lastUsedLocation;
        }

        // Determine if the email is internal or external
        public static String GetEmailAddress(MailItem email, String direction)
        {
            String address;

            if (email.SenderEmailType == "EX")
            {
                AddressEntry addressEntry = direction.Equals("outgoing") ? email.Recipients[1].AddressEntry : email.Sender;

                if (addressEntry.AddressEntryUserType == OlAddressEntryUserType.olExchangeUserAddressEntry
                    || addressEntry.AddressEntryUserType == OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
                {
                    ExchangeUser internalAddress = addressEntry.GetExchangeUser();

                    address = internalAddress.PrimarySmtpAddress;
                }
                else if (addressEntry.Members != null)
                {
                    address = addressEntry.GetExchangeDistributionList().PrimarySmtpAddress;
                }
                else
                {
                    address = email.Recipients[1].Address;
                }
            }
            else
            {
                address = email.Sender.Address;
            }

            return address;
        }

        // Check if the resulting path for the saved emails is valid
        public static String PathCheck(Dictionary<object, object> savePath, MailItem email)
        {
            if ((DialogResult)savePath["dialogResult"] == DialogResult.OK 
                && !string.IsNullOrWhiteSpace((String)savePath["selectedPath"]))
            {
                if (((String)savePath["selectedPath"]).Equals("C:\\"))
                {
                    return "pathInvalid";
                }

                Regex replaceIllegalCharacters = new Regex("[\\/:*?\"<>|]");

                //email.Subject = String.IsNullOrEmpty(email.Subject) ? "No Subject" : replaceIllegalCharacters.Replace(email.Subject, "");

                //return email.Subject.Length > 15 ? email.Subject.Substring(0, 15) + "(...)" : email.Subject;
                String subject = String.IsNullOrEmpty(email.Subject) ? "No Subject" : replaceIllegalCharacters.Replace(email.Subject, "");

                return subject.Length > 15 ? subject.Substring(0, 15) + "(...)" : subject;
            }

            return "saveCancelled";
        }
    }
}
