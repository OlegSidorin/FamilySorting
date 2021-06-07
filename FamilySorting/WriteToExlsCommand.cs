namespace FamilySorting
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using System.IO;
    using System.Reflection;
    using System.Windows.Media.Imaging;
    using Autodesk.Revit.UI;
    using Autodesk.Revit.DB;

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    class WriteToExlsCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;
            string path = Assembly.GetExecutingAssembly().Location;
            //string originalFile = app.SharedParametersFilename;
            string fopFilePath = Path.GetDirectoryName(path) + "\\res\\ФОП.txt";
            if (doc.IsFamilyDocument)
            {
                TaskDialog.Show("Warning", "Privet");
                
                //try
                //{
                //    var p = familyManager.get_Parameter("МСК_Версия Revit");
                //    using (Transaction t = new Transaction(doc, "Set"))
                //    {
                //        t.Start();
                //        string build = commandData.Application.Application.VersionNumber.ToString();
                //        familyManager.Set(p, build);
                //        t.Commit();
                //    }
                                
                //}
                //catch (Exception e)
                //{
                //    TaskDialog.Show("Warning", e.ToString());
                //}
            }
            else
            {
                TaskDialog.Show("Warning", "Это не семейство, команда работает только в семействе");
            }
            return Result.Succeeded;
        }
    }
}
