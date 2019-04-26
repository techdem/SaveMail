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
    public class SaveMail
    {
        public static object[] GetPath(FolderBrowserDialog fbd)
        {
            DialogResult dialogResult = fbd.ShowDialog();
            String selectedPath = fbd.SelectedPath;

            return new object[] { dialogResult, selectedPath };
        }

        public static String GetEmailOrigin(MailItem email)
        {
            String emailName = "";

            if (email.SenderEmailType == "EX")
            {
                AddressEntry address = email.Sender;

                if (address.AddressEntryUserType == OlAddressEntryUserType.olExchangeUserAddressEntry || address.AddressEntryUserType == OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
                {
                    ExchangeUser internalAddress = address.GetExchangeUser();

                    if (internalAddress != null)
                    {
                        emailName = internalAddress.PrimarySmtpAddress;
                    }
                }
            }
            else
            {
                emailName = email.Sender.Address;
            }

            return emailName;
        }

        public static bool SaveSelected(object[] savePath, MailItem[] emailItems)
        {
            String emailSender;

            foreach (MailItem email in emailItems)
            {
                if (email != null && PathCheck(savePath, email.Subject))
                {
                    emailSender = GetEmailOrigin(email);
                    email.SaveAs(savePath[1] + email.ReceivedTime.ToString("dd-MM-yyyy") + " " + emailSender + " " + email.Subject + ".msg", OlSaveAsType.olMSG);

                    return true;
                }
            }
            return false;
        }

        public static bool PathCheck(object[] savePath, String emailName)
        {
            if ((DialogResult)savePath[0] == DialogResult.OK && !string.IsNullOrWhiteSpace((String)savePath[1]))
            {
                Regex replaceIllegalCharacters = new Regex("[\\/:*?\"<>|]");

                replaceIllegalCharacters.Replace(emailName, "");

                return true;
            }
            return false;
        }
    }
}
