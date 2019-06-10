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
        public static MailItem[] GetSelectedEmails()
        {
            Selection selectedEmails = new Microsoft.Office.Interop.Outlook.Application().ActiveExplorer().Selection;
            MailItem[] emailItems = new MailItem[selectedEmails.Count];

            for (int i = 0; i < emailItems.Length; i++)
            {
                emailItems[i] = selectedEmails[i + 1];
            }

            return emailItems;
        }

        public static Dictionary<object, object> GetPath(DialogResult dialogResult, String selectedPath)
        {
            return new Dictionary<object, object>() { { "dialogResult", dialogResult },
                { "selectedPath", selectedPath } };
        }

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
