using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;

namespace SaveMail
{
    // Controller class representing the entry point for the plugin
    public partial class SaveMailController
    {
        private void SaveMailController_Load(object sender, RibbonUIEventArgs e)
        {
        }

        // Main method with at event listener for the plugin button
        private void SaveSelectedButton_Click(object sender, RibbonControlEventArgs e)
        {
            List<object> selectedEmails = SaveMailModel.GetSelectedEmails();

            if (selectedEmails.Count != 0)
            {
                Dictionary<object, object> savePath = SaveMailView.ShowBrowserDialog();
                SaveMailView.Confirmation(SaveSelected(savePath, selectedEmails));
            }
            else
            {
                SaveMailView.Notify("invalidSelection");
            }
        }

        // Method that invokes a sanity check for the path and saves e-mails to drive
        public static String SaveSelected(Dictionary<object, object> savePath, List<object> emailItems)
        {
            String emailSender;

            foreach (MailItem email in emailItems)
            {
                String pathCheckResult = SaveMailModel.PathCheck(savePath, email);

                if (pathCheckResult.Equals("pathOK"))
                {
                    emailSender = SaveMailModel.GetEmailOrigin(email);
                    email.SaveAs(savePath["selectedPath"] + "\\" + email.ReceivedTime.ToString("yyyy-MM-dd") + " " + emailSender + " " + email.Subject + ".msg", OlSaveAsType.olMSG);
                }
                else
                {
                    return pathCheckResult;
                }
            }
            return "saveSuccess";
        }
    }
}
