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
        SaveMailLogging sml = new SaveMailLogging("SaveMail");
        private void SaveMailController_Load(object sender, RibbonUIEventArgs e)
        {
            sml.LogAction("Plugin Loaded!");
        }

        // Main method with at event listener for the plugin button
        private void SaveSelectedButton_Click(object sender, RibbonControlEventArgs e)
        {
            sml.LogAction("Plugin Activated!");
            List<object> selectedEmails = SaveMailModel.GetSelectedEmails();
            sml.LogAction("Selected " + selectedEmails.Count + " emails.");

            if (selectedEmails.Count != 0)
            {
                Dictionary<object, object> savePath = SaveMailView.ShowBrowserDialog();
                SaveMailView.Confirmation(SaveSelected(savePath, selectedEmails));
                
            }
            else
            {
                SaveMailView.Notify("invalidSelection");
                sml.LogAction("Selection contains invalid item...");
            }
        }

        // Method that invokes a sanity check for the path and saves e-mails to drive
        public String SaveSelected(Dictionary<object, object> savePath, List<object> emailItems)
        {
            String emailSender;

            foreach (MailItem email in emailItems)
            {
                String pathCheckResult = SaveMailModel.PathCheck(savePath, email);
                sml.LogAction("Path check result: " + pathCheckResult);

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
