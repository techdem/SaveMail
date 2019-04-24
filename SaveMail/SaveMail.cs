﻿using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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

        public static bool SaveSelected(object[] savePath)
        {
            String emailName;

            foreach (MailItem email in new Microsoft.Office.Interop.Outlook.Application().ActiveExplorer().Selection)
            {
                if (email != null)
                {
                    emailName = GetEmailOrigin(email);

                    if ((DialogResult)savePath[0] == DialogResult.OK && !string.IsNullOrWhiteSpace((String)savePath[1]))
                    {
                        email.SaveAs(savePath[1] + email.ReceivedTime.ToString("dd-MM-yyyy") + " " + emailName + " " + email.Subject + ".msg", OlSaveAsType.olMSG);
                    }
                    else
                    {
                        return false;
                    }
                }
            }

            return true;
        }
    }
}
