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
            SaveMailLogger.LogAction("Plugin Loaded!");
        }

        // Main method with at event listener for the plugin button
        private void SaveSelectedButton_Click(object sender, RibbonControlEventArgs e)
        {
            SaveMailLogger.LogAction("Plugin Activated!");
            List<object> selectedEmails = SaveMailModel.GetSelectedEmails();
            SaveMailLogger.LogAction("Selected " + selectedEmails.Count + " emails.");

            if (selectedEmails.Count != 0)
            {
                Dictionary<object, object> savePath = SaveMailView.ShowBrowserDialog();
                SaveMailView.Confirmation(SaveSelected(savePath, selectedEmails));
                
            }
            else
            {
                SaveMailView.Confirmation("invalidSelection");
            }
        }

        // Method that invokes a sanity check for the path and saves e-mails to drive
        public static String SaveSelected(Dictionary<object, object> savePath, List<object> emailItems)
        {
            String emailSender;
            int savedNumber = 0;

            foreach (MailItem email in emailItems)
            {
                String pathCheckResult = SaveMailModel.PathCheck(savePath, email);

                if (!pathCheckResult.Equals("pathInvalid") && !pathCheckResult.Equals("saveCancelled"))
                {
                    emailSender = SaveMailModel.GetEmailOrigin(email);
                    email.SaveAs(savePath["selectedPath"] + "\\" + email.ReceivedTime.ToString("yyyy-MM-dd HH-mm") + " " + emailSender + " " + pathCheckResult + ".msg", OlSaveAsType.olMSG);
                    savedNumber++;
                }
                else
                {
                    return pathCheckResult;
                }
            }
            SaveMailLogger.LogAction("Saved to location: " + savePath["selectedPath"]);
            SaveMailLogger.LogAction("Total saved: " + savedNumber);
            return "saveSuccess";
        }
    }
}
