﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;

namespace SaveMail
{
    public partial class MainRibbon
    {
        private void MainRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void Button1_Click(object sender, RibbonControlEventArgs e)
        {
            var browsePath = new FolderBrowserDialog();
            String emailName = "";

            foreach(MailItem email in new Microsoft.Office.Interop.Outlook.Application().ActiveExplorer().Selection)
            {
                if(email != null)
                {
                    if(email.SenderEmailType == "EX")
                    {
                        AddressEntry address = email.Sender;

                        if(address.AddressEntryUserType == OlAddressEntryUserType.olExchangeUserAddressEntry || address.AddressEntryUserType == OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
                        {
                            ExchangeUser internalAddress = address.GetExchangeUser();

                            if(internalAddress != null)
                            {
                                emailName = internalAddress.PrimarySmtpAddress;
                            }
                        }
                    }
                    else
                    {
                        emailName = email.SenderEmailAddress;
                    }

                    DialogResult result = browsePath.ShowDialog();

                    if(result == DialogResult.OK && !string.IsNullOrWhiteSpace(browsePath.SelectedPath))
                    {
                        email.SaveAs(browsePath.SelectedPath+emailName+".msg", OlSaveAsType.olMSG);
                        MessageBox.Show("Item saved in: " + browsePath.SelectedPath);
                    }
                }
            }
        }
    }
}
