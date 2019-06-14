using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SaveMail
{
    public static class SaveMailView
    {
        // Display the folder browser dialog
        public static Dictionary<object, object> ShowBrowserDialog()
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog() { Description = "Choose the destination folder for the selected emails." };
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
    }
}
