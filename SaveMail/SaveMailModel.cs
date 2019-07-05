using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
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
                    return new List<object>();
                }
                emailItems.Add(item);
            }
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

                int lastLocation = configFile.LastIndexOf("Saved to location:") + 19;
                int lastConfirmation = configFile.LastIndexOf("Total saved:") - 24;
                int pathLength = lastConfirmation - lastLocation;

                if (pathLength > 0)
                {
                    lastUsedLocation = configFile.Substring(lastLocation, pathLength);
                }
            }

            return lastUsedLocation;
        }

        // Determine if the email is internal or external
        public static String GetEmailOrigin(MailItem email)
        {
            String emailSender = "";

            if (email.SenderEmailType == "EX")
            {
                AddressEntry address = email.Sender;

                if (address.AddressEntryUserType == OlAddressEntryUserType.olExchangeUserAddressEntry
                    || address.AddressEntryUserType == OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
                {
                    ExchangeUser internalAddress = address.GetExchangeUser();

                    if (internalAddress != null)
                    {
                        emailSender = internalAddress.PrimarySmtpAddress;
                    }
                }
            }
            else
            {
                emailSender = email.Sender.Address;
            }

            return emailSender;
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

                return email.Subject != null ? replaceIllegalCharacters.Replace(email.Subject, "") : "No Subject";
            }
            return "saveCancelled";
        }
    }
}
