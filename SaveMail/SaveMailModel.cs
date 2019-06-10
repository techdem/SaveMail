using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SaveMail
{
    public class SaveMailModel
    {
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
                email.Subject = replaceIllegalCharacters.Replace(email.Subject, "");

                return "pathOK";
            }
            return "saveCancelled";
        }
    }
}
