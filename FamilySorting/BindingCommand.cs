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
    class BindingCommand : IExternalCommand
    {
        public string User { get; set; } = Environment.UserName.ToString();
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;
            //string originalFile = app.SharedParametersFilename;
            string fopFilePath = Main.FOPPath;
            string assemblyTablePath = Main.ClassificatorPath;
            string log = "";
            bool isExist = false;

            var resourceReference = ExternalResourceReference.CreateLocalResource(doc, ExternalResourceTypes.BuiltInExternalResourceTypes.AssemblyCodeTable,
                ModelPathUtils.ConvertUserVisiblePathToModelPath(assemblyTablePath), PathType.Absolute) as ExternalResourceReference;
            using (Transaction t = new Transaction(doc, "erer"))
            {
                t.Start();
                AssemblyCodeTable.GetAssemblyCodeTable(doc).LoadFrom(resourceReference, new KeyBasedTreeEntriesLoadResults());
                t.Commit();
            }


            string[] paramtersArray =
            {
                "КПСП_GUID семейства", "КПСП_Дисциплина", "КПСП_Категория", "КПСП_Подкатегория", "МСК_Версия Revit", "МСК_Версия семейства", "КПСП_Статус",  
                "КПСП_Библиотека семейств", "КПСП_Инструкция", "КПСП_Путь к семейству", "КПСП_Дата редактирования", "КПСП_Автор", "КПСП_Вложенные семейства"
            };
            string[] paramtersMSKTypeArray =
            {
                "МСК_Марка", "МСК_Наименование", "МСК_Завод-изготовитель", "МСК_Материал", "МСК_Описание", "МСК_Масса", "МСК_Масса_Текст",
                "МСК_Размер_Ширина", "МСК_Размер_Высота", "МСК_Размер_Толщина", "МСК_Размер_Глубина", "МСК_ЕдИзм", "МСК_Примечание", "МСК_Обозначение",
            };
            string[] paramtersMSKTypeAnnotationArray =
            {
                "МСК_Марка", "МСК_Наименование", "МСК_Описание", "МСК_ЕдИзм", "МСК_Примечание", "МСК_Обозначение"
            };
            string[] paramtersMSKInstArray =
            {
                "МСК_Наименование краткое", "МСК_Код изделия",  "МСК_Позиция"
            };
            
            if (doc.IsFamilyDocument)
            {
                //TaskDialog.Show("Warning", "Privet");
                FamilyManager familyManager = doc.FamilyManager;
                FamilyType familyType;
                familyType = familyManager.CurrentType;
                #region check is family type is Annotation
                bool isAnnotation = false;
                try
                {

                    var annotationFamily = new FilteredElementCollector(doc).OfClass(typeof(Family)).Cast<Family>().ToList().FirstOrDefault();
                    if (annotationFamily.FamilyCategory.CategoryType.ToString() == "Annotation")
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

                try
                {
                    commandData.Application.Application.SharedParametersFilename = fopFilePath;
                    using (Transaction t = new Transaction(doc,"Add paramters"))
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
                        if (!isAnnotation)
                        {
                            isExist = false;
                            foreach (var st in paramtersMSKTypeArray)
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
                                    if (sharedParameterDefinition.Name == "МСК_Материал")
                                        familyManager.AddParameter(externalDefinition, BuiltInParameterGroup.PG_MATERIALS, false);
                                    else if (sharedParameterDefinition.Name.Contains("МСК_Размер_"))
                                        familyManager.AddParameter(externalDefinition, BuiltInParameterGroup.PG_GEOMETRY, false);
                                    else
                                        familyManager.AddParameter(externalDefinition, BuiltInParameterGroup.PG_TEXT, false);
                                    log += "\nДобавлен параметр <" + st + ">";
                                }
                                isExist = false;
                            }
                            isExist = false;
                            foreach (var st in paramtersMSKInstArray)
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
                                    familyManager.AddParameter(externalDefinition, BuiltInParameterGroup.PG_TEXT, true);
                                    log += "\nДобавлен параметр <" + st + ">";
                                }
                                isExist = false;
                            }
                        }
                        if (isAnnotation)
                        {
                            isExist = false;
                            foreach (var st in paramtersMSKTypeAnnotationArray)
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
                                    if (sharedParameterDefinition.Name == "МСК_Материал")
                                        familyManager.AddParameter(externalDefinition, BuiltInParameterGroup.PG_MATERIALS, false);
                                    else if (sharedParameterDefinition.Name.Contains("МСК_Размер_"))
                                        familyManager.AddParameter(externalDefinition, BuiltInParameterGroup.PG_GEOMETRY, false);
                                    else
                                        familyManager.AddParameter(externalDefinition, BuiltInParameterGroup.PG_TEXT, false);
                                    log += "\nДобавлен параметр <" + st + ">";
                                }
                                isExist = false;
                            }
                            isExist = false;
                            foreach (var st in paramtersMSKInstArray)
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
                                    familyManager.AddParameter(externalDefinition, BuiltInParameterGroup.PG_TEXT, true);
                                    log += "\nДобавлен параметр <" + st + ">";
                                }
                                isExist = false;
                            }
                        }
                        t.Commit();
                    }
                }
                catch (Exception e)
                {
                    TaskDialog.Show("Warning 1", e.ToString());
                }
                try
                {
                    using (Transaction t = new Transaction(doc, "Set parameters"))
                    {
                        t.Start();
                        log += "\n---";

                        var p = familyManager.get_Parameter("МСК_Версия Revit");
                        string build = "R" + commandData.Application.Application.VersionNumber.ToString().Remove(0,2);
                        familyManager.Set(p, build);
                        log += "\nНовое значение <МСК_Версия Revit>: " + build;

                        p = familyManager.get_Parameter("КПСП_GUID семейства");

                        if (familyType.AsString(p) == "")
                        {
                            Guid famGuid = Guid.NewGuid();
                            familyManager.Set(p, famGuid.ToString());
                            log += "\nПрисвоен новый Guid семейству: " + familyType.AsString(p);
                        }

                        /*
                        p = familyManager.get_Parameter("МСК_Версия семейства");
                        string vs = familyType.AsString(p);
                        int vs_int = 1;
                        if (vs != "")
                        {
                            int.TryParse(vs, out vs_int);
                            vs_int += 1;
                            familyManager.Set(p, vs_int.ToString());
                            log += "\nВерсия семейства: " + familyType.AsString(p);
                        }
                        if (vs == "")
                        {
                            familyManager.Set(p, "1");
                            log += "\nВерсия семейства: " + familyType.AsString(p);
                        }
                        */

                        p = familyManager.get_Parameter("КПСП_Дата редактирования");

                        DateTime today = DateTime.Now;

                        int.TryParse(today.Day.ToString(), out int dayInt);
                        int.TryParse(today.Month.ToString(), out int monthInt);
                        string sDate = String.Format("{0:D2}-{1:D2}-{2}", dayInt, monthInt, today.Year.ToString());
                        familyManager.Set(p, sDate);
                        log += "\nДата редактирования: " + familyType.AsString(p);

                        p = familyManager.get_Parameter("КПСП_Автор");
                        familyManager.Set(p, User);
                        log += "\nАвтор: " + familyType.AsString(p);

                        #region set parameters with Classificator Cod
                        if (!isAnnotation)
                        {
                            var pKey = familyManager.get_Parameter("Код по классификатору");
                            string pKeyValue = familyType.AsString(pKey);

                            p = familyManager.get_Parameter("КПСП_Дисциплина");
                            familyManager.Set(p, TableEntry.GetDiscipline(pKeyValue));
                            log += "\nДисциплина: " + familyType.AsString(p) + ": " + familyType.AsString(pKey);
                            p = familyManager.get_Parameter("КПСП_Категория");
                            familyManager.Set(p, TableEntry.GetCategory(pKeyValue));
                            log += "\nКатегория: " + familyType.AsString(p) + ": " + familyType.AsString(pKey);
                            p = familyManager.get_Parameter("КПСП_Подкатегория");
                            familyManager.Set(p, TableEntry.GetSubCategory(pKeyValue));
                            log += "\nПодкатегория: " + familyType.AsString(p) + ": " + familyType.AsString(pKey);

                            p = familyManager.get_Parameter("КПСП_Путь к семейству");
                            familyManager.Set(p, TableEntry.GetPathToFamily(pKeyValue).Replace("RXX", build));

                            log += "\nПуть к семейству: " + familyType.AsString(p) + ": " + familyType.AsString(pKey);
                            p = familyManager.get_Parameter("КПСП_Инструкция");
                            familyManager.Set(p, TableEntry.GetPathToInstruction(pKeyValue));
                            log += "\nИнструкция: " + familyType.AsString(p) + ": " + familyType.AsString(pKey);
                        }
                        #endregion

                        #region set parameters if Annotation
                        if (isAnnotation)
                        {

                            p = familyManager.get_Parameter("КПСП_Дисциплина");
                            familyManager.Set(p, "Общие");

                            p = familyManager.get_Parameter("КПСП_Категория");
                            familyManager.Set(p, "Аннотации");

                            try
                            {
                                var annotationFamily = new FilteredElementCollector(doc).OfClass(typeof(Family)).Cast<Family>().ToList().FirstOrDefault();
                                if (annotationFamily.FamilyCategory.CategoryType.ToString() == "Annotation")
                                {
                                    if (annotationFamily.FamilyCategory.Name.Contains("Марки "))
                                    {
                                        p = familyManager.get_Parameter("КПСП_Подкатегория");
                                        familyManager.Set(p, "Марки");
                                        p = familyManager.get_Parameter("КПСП_Путь к семейству");
                                        familyManager.Set(p, @"K:\Стандарт\ТИМ Семейства\0_Библиотека семейств RXX\1 Общие\Аннотации\Марки".Replace("RXX", build));
                                    }
                                    else if (annotationFamily.FamilyCategory.Name.Contains("Просмотр заголовков"))
                                    {
                                        p = familyManager.get_Parameter("КПСП_Подкатегория");
                                        familyManager.Set(p, "Название вида");
                                        p = familyManager.get_Parameter("КПСП_Путь к семейству");
                                        familyManager.Set(p, @"K:\Стандарт\ТИМ Семейства\0_Библиотека семейств RXX\1 Общие\Аннотации\Название вида".Replace("RXX", build));
                                    }
                                    else if (annotationFamily.FamilyCategory.Name.Contains("Обозначения высотных отметок"))
                                    {
                                        p = familyManager.get_Parameter("КПСП_Подкатегория");
                                        familyManager.Set(p, "Обозначение высотной отметки");
                                        p = familyManager.get_Parameter("КПСП_Путь к семейству");
                                        familyManager.Set(p, @"K:\Стандарт\ТИМ Семейства\0_Библиотека семейств RXX\1 Общие\Аннотации\Обозначение высотной отметки".Replace("RXX", build));
                                    }
                                    else if (annotationFamily.FamilyCategory.Name.Contains("Обозначения осей"))
                                    {
                                        p = familyManager.get_Parameter("КПСП_Подкатегория");
                                        familyManager.Set(p, "Обозначение оси");
                                        p = familyManager.get_Parameter("КПСП_Путь к семейству");
                                        familyManager.Set(p, @"K:\Стандарт\ТИМ Семейства\0_Библиотека семейств RXX\1 Общие\Аннотации\Обозначение оси".Replace("RXX", build));
                                    }
                                    else
                                    {
                                        p = familyManager.get_Parameter("КПСП_Подкатегория");
                                        familyManager.Set(p, "Типовые аннотации");
                                        p = familyManager.get_Parameter("КПСП_Путь к семейству");
                                        familyManager.Set(p, @"K:\Стандарт\ТИМ Семейства\0_Библиотека семейств RXX\1 Общие\Аннотации\Типовые аннотации".Replace("RXX", build));
                                    }
                                }
                                else
                                {
                                    
                                }

                            }
                            catch (Exception ex)
                            {
                                TaskDialog.Show("Warning2", ex.ToString());
                            }

                            /*
                            p = familyManager.get_Parameter("КПСП_Подкатегория");
                            familyManager.Set(p, TableEntry.GetSubCategory(pKeyValue));
                            log += "\nПодкатегория: " + familyType.AsString(p) + ": " + familyType.AsString(pKey);

                            p = familyManager.get_Parameter("КПСП_Путь к семейству");
                            familyManager.Set(p, TableEntry.GetPathToFamily(pKeyValue).Replace("RXX", build));

                            p = familyManager.get_Parameter("КПСП_Инструкция");
                            familyManager.Set(p, TableEntry.GetPathToInstruction(pKeyValue));
                            */
                        }
                        #endregion

                        IList<FamilyInstance> vlozhennieSemeistva = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).WhereElementIsNotElementType().Cast<FamilyInstance>().ToList();
                        string spisokSemeistv = "";
                        int i = 0;
                        int vsegoSemeistv = vlozhennieSemeistva.Count;
                        string estObschieSemeistva = "Нет";
                        List<string> guids = new List<string>();
                        foreach (FamilyInstance fi in vlozhennieSemeistva)
                        {
                            if (i == 0)
                                spisokSemeistv += "всего " + vsegoSemeistv.ToString() + ": ";
                            if (i != 0)
                                spisokSemeistv += ", ";
                            FamilySymbol fType = fi.Symbol;
                            Family fam = fType.Family;
                            spisokSemeistv += fam.Name; // + "(" + fam.get_Parameter(new Guid("11b18c00-5d82-4226-8b5f-74526a7ec4f8")) + ")";
                            try
                            {
                                string str = fType.get_Parameter(new Guid("11b18c00-5d82-4226-8b5f-74526a7ec4f8")).AsString();
                                spisokSemeistv += "(" + str + ")";
                                estObschieSemeistva = "Да";
                            }
                            catch
                            {
                                // do nothing
                            }
                            i += 1;
                        }
                        
                        //if (spisokSemeistv.Contains("КПСП"))
                        //    estObschieSemeistva = "Да";
                        p = familyManager.get_Parameter("КПСП_Вложенные семейства");
                        familyManager.Set(p, estObschieSemeistva);
                        //log += "\nСемейства: " + familyType.AsString(p) + ": " + familyType.AsString(pKey);

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
                TaskDialog.Show("Warning", "Это не семейство, команда работает только в семействе");
            }

            //TaskDialog.Show("Final", log);
            return Result.Succeeded;
        }
    }
}
