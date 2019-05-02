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
            String selectedPath = fbd.SelectedPath;

            return new Dictionary<object, object>() { { "dialogResult", dialogResult },
                { "selectedPath", selectedPath } };
        }
    }
}
