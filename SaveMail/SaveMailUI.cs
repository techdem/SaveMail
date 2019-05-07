using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SaveMail
{
    public class SaveMailUI
    {
        public static Dictionary<object, object> ShowDialog()
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog() { Description = "Choose the destination folder for the selected emails." };
            DialogResult dialogResult = fbd.ShowDialog();

            return SaveMail.GetPath(dialogResult, fbd.SelectedPath);
        }

        public static void Confirmation(String input)
        {
            if (input.Equals("success"))
            {
                MessageBox.Show("Saved successfully!", "SaveMail");
            }
            else
            {
                MessageBox.Show("Items not saved.", "SaveMail");
            }
        }

        public static void Notify(String input)
        {
            if (input.Equals("invalidPath"))
            {
                MessageBox.Show("Invalid path, please choose a different folder.", "SaveMail");
            }
        }
    }
}
