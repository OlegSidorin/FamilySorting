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
    using System.Threading;

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    class SaveCommand : IExternalCommand
    {
        public string User { get; set; } = Environment.UserName.ToString();
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;
            string path = Assembly.GetExecutingAssembly().Location;

            if (doc.IsFamilyDocument)
            {
                FamilyManager familyManager = doc.FamilyManager;
                FamilyType familyType;
                familyType = familyManager.CurrentType;
                if (familyType == null)
                {
                    using (Transaction t = new Transaction(doc, "change"))
                    {
                        t.Start();
                        familyType = familyManager.NewType(doc.Title);
                        familyManager.CurrentType = familyType;
                        t.Commit();
                    }
                }

                #region check is family type is Annotation
                bool isAnnotation = false;
                try
                {

                    var fm = new FilteredElementCollector(doc).OfClass(typeof(Family)).Cast<Family>().ToList().FirstOrDefault();
                    if (fm.FamilyCategory.CategoryType.ToString() == "Annotation")
                    {
                        isAnnotation = true;
                    }
                    else
                    {
                        isAnnotation = false;
                    }

                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Warning2", ex.ToString());
                }
                #endregion

                string pathTo = "";
                try
                {
                    var pathToParameter = familyManager.get_Parameter("КПСП_Путь к семейству");
                    pathTo = familyType.AsString(pathToParameter);
                    SaveForm saveForm = new SaveForm();
                    saveForm.Doc = doc;
                    saveForm.textFamilyName.Text = doc.Title;
                    saveForm.labelPath.Text = pathTo;
                    saveForm.textComment.Text = "";
                    saveForm.Show();
                }
                catch
                {
                    TaskDialog.Show("Warning", "Не заполнены параметры семейства");
                }
                
                

                //FileSaveDialog fileSaveDialog = new FileSaveDialog("Revit Families (*.rfa)|*.rfa"); // "Revit Projects (*.rvt)|*.rvt|Revit Families (*.rfa)|*.rfa"
                //fileSaveDialog.InitialFileName = doc.Title;
                //fileSaveDialog.Show();

                //doc.Save();
                //doc.SaveAs(@"C:\Users\Sidorin_O\Documents\TEST\testsavefamily.rfa", sao);
            }
            else
            {
                TaskDialog.Show("Warning", "Это не семейство, команда работает только в семействе");
            }

            //TaskDialog.Show("Final", "Файл классификатора загружен" + "\n" + assemblyTablePath + "\n - - - - - -" + str);

            return Result.Succeeded;
        }
    }
}
