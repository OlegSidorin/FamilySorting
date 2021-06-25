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
    class ClearCommand : IExternalCommand
    {
        public string User { get; set; } = Environment.UserName.ToString();
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;
            string path = Assembly.GetExecutingAssembly().Location;
            //string originalFile = app.SharedParametersFilename;
            string fopFilePath = Path.GetDirectoryName(path) + "\\res\\ФОП.txt";
            string log = "";
            bool isExist = false;
            string[] paramtersArray =
            {
                "КПСП_GUID семейства", "МСК_Версия семейства"
            };
            
            if (doc.IsFamilyDocument)
            {
                //TaskDialog.Show("Warning", "Privet");
                FamilyManager familyManager = doc.FamilyManager;
                FamilyType familyType;
                familyType = familyManager.CurrentType;
                if (familyType == null)
                {
                    using (Transaction t = new Transaction(doc, "change"))
                    {
                        t.Start();
                        familyType = familyManager.NewType("Тип 1");
                        familyManager.CurrentType = familyType;
                        t.Commit();
                    }
                }

                try
                {
                    commandData.Application.Application.SharedParametersFilename = fopFilePath;
                    using (Transaction t = new Transaction(doc,"Clear GUID"))
                    {
                        t.Start();
                        DefinitionFile sharedParametersFile = commandData.Application.Application.OpenSharedParameterFile();
                        DefinitionGroup sharedParametersGroup = sharedParametersFile.Groups.get_Item("14_Управление семействами");
                        Definition sharedParameterDefinition;
                        ExternalDefinition externalDefinition;

                        FamilyParameterSet parametersList = familyManager.Parameters;

                        isExist = false;
                        foreach (var st in paramtersArray)
                        {
                            foreach (FamilyParameter fp in parametersList)
                            {
                                if (st == fp.Definition.Name)                              
                                    isExist = true;
                            }
                            if (!isExist)
                            {
                                sharedParameterDefinition = sharedParametersGroup.Definitions.get_Item(st);
                                externalDefinition = sharedParameterDefinition as ExternalDefinition;
                                familyManager.AddParameter(externalDefinition, BuiltInParameterGroup.PG_ADSK_MODEL_PROPERTIES, false);
                                log += "\nДобавлен параметр <" + st + ">";
                            }
                            isExist = false;

                        }
                        isExist = false;
                        
                        t.Commit();
                    }

                }
                catch (Exception e)
                {
                    TaskDialog.Show("Warning 1", e.ToString());
                }
                try
                {
                    using (Transaction t = new Transaction(doc, "Apply"))
                    {
                        t.Start();
                        log += "\n---";

                        var p = familyManager.get_Parameter("КПСП_GUID семейства");

                        if (familyType.AsString(p) != "")
                        {
                            familyManager.Set(p, "");
                            log += "\nКПСП_GUID семейства пуст ";
                        }

                        p = familyManager.get_Parameter("МСК_Версия семейства");

                        if (familyType.AsString(p) != "")
                        {
                            familyManager.Set(p, "");
                            log += "\nМСК_Версия семейства пуст ";
                        }

                        t.Commit();
                    }
                }
                catch (Exception e)
                {
                    TaskDialog.Show("Warning 2", e.ToString());
                }

            }
            else
            {
                TaskDialog.Show("Warning main", "Это не семейство, команда работает только в семействе");
            }

            //TaskDialog.Show("Final", log);
            return Result.Succeeded;
        }
    }
}
