using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace SaveMail
{
    public static class SaveMailView
    {
        // Display the folder browser dialog
        public static Dictionary<object, object> ShowBrowserDialog()
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog() { Description = "Choose the destination folder for the selected emails." };
            fbd.RootFolder = Environment.SpecialFolder.Desktop;
            fbd.SelectedPath = SaveMailModel.GetRecentLocation();
            DialogResult dialogResult = fbd.ShowDialog();

            return SaveMailModel.GetPath(dialogResult, fbd.SelectedPath);
        }

        // Display the confirmation message or invoke the notification method
        public static void Confirmation(String input)
        {
            if (input.Equals("saveSuccess"))
            {
                MessageBox.Show("Saved successfully!", "SaveMail");
            }
            else
            {
                Notify(input);
                
            }

            SaveMailLogger.LogAction("Action result: " + input);
        }

        // Notify the user if something has gone wrong
        public static void Notify(String input)
        {
            if (input.Equals("invalidSelection"))
            {
                MessageBox.Show("Please make sure selected e-mails don't contain calendar appointments!", "SaveMail");
            }

            if (input.Equals("pathInvalid"))
            {
                MessageBox.Show("Invalid path, please choose a different folder.", "SaveMail");
            }

            if (input.Equals("saveCancelled"))
            {
                MessageBox.Show("Operation cancelled.", "SaveMail");
            }
        }

        // Ask the user if they would like to move the selected e-mails out of the inbox
        public static Boolean Question()
        {
            return MessageBox.Show("Would you also like to move the selected e-mails out of the Inbox and into the Saved Mail folder?", "Save Mail", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes ? true : false;
        }
    }
}
