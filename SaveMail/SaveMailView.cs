using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SaveMail
{
    public static class SaveMailView
    {
        readonly static string applicationName = "SaveMail";
        readonly static string logFile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\AppData\\Local\\" + applicationName + "\\Log-" + Environment.MachineName + ".txt";

        // Display the folder browser dialog
        public static Dictionary<object, object> ShowBrowserDialog()
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog() { Description = "Choose the destination folder for the selected emails." };
            fbd.RootFolder = Environment.SpecialFolder.MyComputer;

            using (StreamReader sr = new StreamReader(logFile))
            {
                String configFile = sr.ReadToEnd();

                int lastLocation = configFile.LastIndexOf("Saved to location:") + 19;
                int lastConfirmation = configFile.LastIndexOf("Total saved:") - 24;
                int pathLength = lastConfirmation - lastLocation;

                if(pathLength > 0)
                {
                    fbd.SelectedPath = configFile.Substring(lastLocation, pathLength);
                }
            }
            
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
    }
}
