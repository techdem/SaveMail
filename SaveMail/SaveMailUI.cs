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
        public IDictionary<String, String> ShowDialog()
        {
            return new Dictionary<String, String>() { { "dialogResult", "OK" }, { "selectedPath", "C:\\" } };
        }
    }
}
