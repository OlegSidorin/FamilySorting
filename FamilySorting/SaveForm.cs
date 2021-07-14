using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;

namespace FamilySorting
{
    public partial class SaveForm : System.Windows.Forms.Form
    {
        public Document Doc { get; set; }
        public SaveForm()
        {
            InitializeComponent();
        }

        private void saveButton_Click(object sender, EventArgs e)
        {
            SaveAsOptions sao = new SaveAsOptions();
            sao.OverwriteExistingFile = true;

            if (!Directory.Exists(labelPath.Text.ToString() + @"\")) 
            {
                Directory.CreateDirectory(labelPath.Text.ToString() + @"\");
            }

            Doc.SaveAs(labelPath.Text.ToString() + @"\" + textFamilyName.Text.ToString() + ".rfa", sao);
            Main.Comment = textComment.Text.ToString();
            Close();
        }
    }
}
