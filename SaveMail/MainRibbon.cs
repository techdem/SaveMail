using System;
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
        static FolderBrowserDialog fbd;
        String emailName = "";

        private void MainRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            fbd = new FolderBrowserDialog();
        }

        private void Button1_Click(object sender, RibbonControlEventArgs e)
        {
            Object[] getPath = GetPath(fbd);

            foreach (MailItem email in new Microsoft.Office.Interop.Outlook.Application().ActiveExplorer().Selection)
            {
                if (email != null)
                {
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
                        emailName = email.SenderEmailAddress;
                    }

                    if ((DialogResult)getPath[0] == DialogResult.OK && !string.IsNullOrWhiteSpace((String)getPath[1]))
                    {
                        email.SaveAs(getPath[1] + email.ReceivedTime.ToString("dd-MM-yyyy") + " " + emailName + " " + email.Subject + ".msg", OlSaveAsType.olMSG);
                    }
                }
            }
            MessageBox.Show("Saved successfully in:\n\n" + (String)getPath[2]);
        }

        public static object[] GetPath(FolderBrowserDialog fbd)
        {
            DialogResult dialogResult = fbd.ShowDialog();
            String selectedPath = fbd.SelectedPath;

            return new object[] { dialogResult, selectedPath };
        }
    }
}
