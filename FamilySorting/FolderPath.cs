using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FamilySorting
{
    public partial class FolderPath : Form
    {
        public FolderPath()
        {
            InitializeComponent();
        }

        private void button_Click(object sender, EventArgs e)
        {
            string path = Assembly.GetExecutingAssembly().Location;
            string familyFolderPath = path.Replace(@"\FamilySorting.dll", "") + @"\res\folder.txt";
            
            if (File.Exists(familyFolderPath))
            {
                // Create a file to write to.
                using (StreamWriter sw = File.CreateText(familyFolderPath))
                {
                    sw.Write(textBox.Text);
                }
            }
            Main.Folder = textBox.Text;
            Close();
        }
    }
}
