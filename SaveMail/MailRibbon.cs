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
        FolderBrowserDialog fbd;
        private void MainRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            fbd = new FolderBrowserDialog();
        }

        private void SaveSelectedButton_Click(object sender, RibbonControlEventArgs e)
        {
            object[] savePath = SaveMail.GetPath(fbd);

            if (SaveMail.SaveSelected(savePath) == true)
            {
                MessageBox.Show("Saved successfully in:\n\n" + (String)savePath[1]);
            }
            else
            {
                MessageBox.Show("Items not saved");
            }
        }
    }
}
