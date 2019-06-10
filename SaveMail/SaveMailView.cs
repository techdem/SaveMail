using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SaveMail
{
    public class SaveMailView
    {
        public static Dictionary<object, object> ShowBrowserDialog()
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog() { Description = "Choose the destination folder for the selected emails." };
            DialogResult dialogResult = fbd.ShowDialog();

            return SaveMailModel.GetPath(dialogResult, fbd.SelectedPath);
        }

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

        public static void Notify(String input)
        {
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
