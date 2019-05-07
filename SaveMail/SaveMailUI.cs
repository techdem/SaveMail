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
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            DialogResult dialogResult = fbd.ShowDialog();

            return SaveMail.GetPath(dialogResult, fbd.SelectedPath);
        }

        public static void Confirmation(String output)
        {
            if (output.Equals("success"))
            {
                MessageBox.Show("Saved successfully");
            }
            else
            {
                MessageBox.Show("Items not saved.");
            }
        }
    }
}
