using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;

namespace SaveMail
{
    public partial class MailRibbon
    {
        private void MailRibbon_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void SaveSelectedButton_Click(object sender, RibbonControlEventArgs e)
        {
            Selection selectedMail = new Microsoft.Office.Interop.Outlook.Application().ActiveExplorer().Selection;
            MailItem[] emailItems = new MailItem[selectedMail.Count];
            Dictionary<object, object> savePath = SaveMailUI.ShowDialog();

            for (int i = 0; i < emailItems.Length; i++)
            {
                emailItems[i] = selectedMail[i+1];
            }

            SaveMailUI.Confirmation(SaveMail.SaveSelected(savePath, emailItems));
        }
    }
}
