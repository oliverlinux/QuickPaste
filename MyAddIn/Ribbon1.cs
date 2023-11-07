using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Windows.Forms;

namespace MyAddIn
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
        
            string usertext = "";
            try
            {
                usertext = File.ReadAllText(System.Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\\data.dat");
                
            }
            catch
            {
                Form TheSettingsForError = new Settings();
                TheSettingsForError.Visible = true;
                TheSettingsForError.Activate();
                
            }
            if (usertext != "")
            {
                if (Globals.Ribbons.Ribbon1.checkBox1.Checked == true)
                {
                    Globals.ThisAddIn.Application.ActiveDocument.Content.Text = Globals.ThisAddIn.Application.ActiveDocument.Content.Text + "\n" + usertext;
                }
                else
                {
                    Globals.ThisAddIn.Application.ActiveDocument.Content.InsertAfter(usertext);
                }
            }
            
            
        }
        


        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Form TheSettings = new Settings();
            TheSettings.Visible = true;
            TheSettings.Activate();
            
        }

    }
}
