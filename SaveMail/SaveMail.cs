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
            String emailSender = "";

            if (email.SenderEmailType == "EX")
            {
                AddressEntry address = email.Sender;

                if (address.AddressEntryUserType == OlAddressEntryUserType.olExchangeUserAddressEntry || address.AddressEntryUserType == OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
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

        public static bool SaveSelected(object[] savePath, MailItem[] emailItems)
        {
            String emailSender;

            foreach (MailItem email in emailItems)
            {
                if (email != null && PathCheck(savePath, email))
                {
                    emailSender = GetEmailOrigin(email);
                    email.SaveAs(savePath[1] + "\\" + email.ReceivedTime.ToString("dd-MM-yyyy") + " " + emailSender + " " + email.Subject + ".msg", OlSaveAsType.olMSG);
                }
                else
                {
                    return false;
                }
            }
            return true;
        }

        public static bool PathCheck(object[] savePath, MailItem email)
        {
            if (((String)savePath[1]).Equals("C:\\"))
            {
                return false;
            }

            if ((DialogResult)savePath[0] == DialogResult.OK 
                && !string.IsNullOrWhiteSpace((String)savePath[1]))
            {
                Regex replaceIllegalCharacters = new Regex("[\\/:*?\"<>|]");
                email.Subject = replaceIllegalCharacters.Replace(email.Subject, "");

                return true;
            }

            return false;
        }
    }
}
