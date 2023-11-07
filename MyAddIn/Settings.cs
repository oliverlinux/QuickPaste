using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MyAddIn
{
    public partial class Settings : Form

    {
        
        public Settings()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            File.WriteAllText(System.Environment.SpecialFolder.LocalApplicationData + "data.dat",richTextBox1.Text);

        }
    }
}
