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
        public static Dictionary<object, object> GetPath()
        {
            Dictionary<object,object> showDialog = SaveMailUI.ShowDialog();

            return new Dictionary<object, object> { { "dialogResult", showDialog["dialogResult"] },
                { "selectedPath", showDialog["selectedPath"] } };
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

        public static bool SaveSelected(Dictionary<object, object> savePath, MailItem[] emailItems)
        {
            String emailSender;

            foreach (MailItem email in emailItems)
            {
                if (email != null && PathCheck(savePath, email))
                {
                    emailSender = GetEmailOrigin(email);
                    email.SaveAs(savePath["selectedPath"] + "\\" + email.ReceivedTime.ToString("dd-MM-yyyy") + " " + emailSender + " " + email.Subject + ".msg", OlSaveAsType.olMSG);
                }
                else
                {
                    return false;
                }
            }
            return true;
        }

        public static bool PathCheck(Dictionary<object, object> savePath, MailItem email)
        {
            if (((String)savePath["selectedPath"]).Equals("C:\\"))
            {
                return false;
            }

            if ((DialogResult)savePath["dialogResult"] == DialogResult.OK 
                && !string.IsNullOrWhiteSpace((String)savePath["selectedPath"]))
            {
                Regex replaceIllegalCharacters = new Regex("[\\/:*?\"<>|]");
                email.Subject = replaceIllegalCharacters.Replace(email.Subject, "");

                return true;
            }

            return false;
        }
    }
}
