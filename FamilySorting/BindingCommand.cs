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

            string[] guidsArray =
            {
                "11b18c00-5d82-4226-8b5f-74526a7ec4f8", "3c0744d8-3713-4311-b03e-885f7441d360", "f956074c-276f-4f03-bcdf-890dc4a6038a",
                "eb3b5f14-25f5-41eb-942f-d8dd33d766c1",
                "37384649-c3c8-4fc2-a08e-c2206438f528", "85cd0032-c9ee-4cd3-8ffa-b2f1a05328e3", "fd97b929-0274-408b-8299-9981cb982fc5",
                "fdf7bfa4-5294-45c5-b979-c388d3a062da", "18ea3aa8-5275-470f-94da-e35bb4c80e46", "728fe6d4-f0e9-4418-b261-25c67382b379",
                "0b3fd4ed-0256-43e5-a997-5311f4c19091", "A80FE9BB-B06E-46BD-B50D-D32486ED228F", "8F0F22FE-8DA1-4CF3-B94D-3DC33041E5D3",
                "fb30c7d4-3e3c-4fe6-821b-189cf35b7f9f", "647b5bc9-6570-416c-93d3-bd0d159775f2", "a8cdbf7b-d60a-485e-a520-447d2055f351",
                "8b5e61a2-b091-491c-8092-0b01a55d4f45", "9b3dbd60-5be3-4842-9dbe-cd644ef5f9e8", "946c4e27-a56c-422d-999c-778a150b950e",
                "a8832df7-0302-4a63-a6e1-47a01632b987", "8f2e4f93-9472-4941-a65d-0ac468fd6a6d", "da753fe3-ecfa-465b-9a2c-02f55d0c2ff1",
                "ef3ac60d-2cf8-4bd8-bd66-dbcb42e92f4a",
                "f13b35e5-9fb9-4cf8-b330-efe01d3780c4", "e7edd112-da46-46c3-886c-934dad841efb",
                "bfa2f0d2-ccd0-4a02-95c7-573f0a9829c3", "2fd9e8cb-84f3-4297-b8b8-75f444e124ed",  "ae8ff999-1f22-4ed7-ad33-61503d85f0f4"
            };
            string[] parametersForGuidsArray =
            {
                "КПСП_GUID семейства", "КПСП_Дисциплина", "КПСП_Категория", "КПСП_Подкатегория", "МСК_Версия Revit", "МСК_Версия семейства", "КПСП_Статус",
                "КПСП_Библиотека семейств", "КПСП_Инструкция", "КПСП_Путь к семейству", "КПСП_Дата редактирования", "КПСП_Автор", "КПСП_Вложенные семейства",
                "МСК_Марка", "МСК_Наименование", "МСК_Завод-изготовитель", "МСК_Материал", "МСК_Описание", "МСК_Масса", "МСК_Масса_Текст",
                "МСК_Размер_Ширина", "МСК_Размер_Высота", "МСК_ЕдИзм", "МСК_Примечание", "МСК_Обозначение",
                "МСК_Наименование краткое", "МСК_Код изделия",  "МСК_Позиция"
            };
            string[] parametersArray =
            {
                "КПСП_GUID семейства", "КПСП_Дисциплина", "КПСП_Категория", "КПСП_Подкатегория", "МСК_Версия Revit", "МСК_Версия семейства", "КПСП_Статус",  
                "КПСП_Библиотека семейств", "КПСП_Инструкция", "КПСП_Путь к семейству", "КПСП_Дата редактирования", "КПСП_Автор", "КПСП_Вложенные семейства"
            };
            string[] paramtersMSKTypeArray =
            {
                "МСК_Марка", "МСК_Наименование", "МСК_Завод-изготовитель", "МСК_Материал", "МСК_Описание", "МСК_Масса", "МСК_Масса_Текст",
                "МСК_Размер_Ширина", "МСК_Размер_Высота", "МСК_ЕдИзм", "МСК_Примечание", "МСК_Обозначение"
            };
            string[] paramtersMSKTypeAnnotationArray =
            {
                //"МСК_Марка", "МСК_Наименование", "МСК_Описание", "МСК_ЕдИзм", "МСК_Примечание", "МСК_Обозначение", "МСК_Завод-изготовитель"
            };
            string[] paramtersMSKInstArray =
            {
                "МСК_Наименование краткое", "МСК_Код изделия",  "МСК_Позиция"
            };
            string[] paramtersMSKInstAnnotationArray =
{
                //"МСК_Наименование краткое", "МСК_Код изделия",  "МСК_Позиция"
            };

            if (doc.IsFamilyDocument)
            {
                
                FamilyManager familyManager = doc.FamilyManager;
                FamilyType familyType = familyManager.CurrentType;
                FamilyParameterSet parametersList = familyManager.Parameters;

                #region check and delete wrong parameter
                for (var i = 0; i < guidsArray.Length; i++)
                {
                    var p = familyManager.get_Parameter(new Guid(guidsArray[i]));
                    if (p != null)
                    {
                        if (!(p.Definition.Name == parametersForGuidsArray[i]))
                        {
                            using (Transaction tr = new Transaction(doc, "delete parameter"))
                            {
                                tr.Start();
                                string str1 = p.Definition.Name;
                                familyManager.RemoveParameter(p);
                                TaskDialog.Show("Warning", "Был удален общий параметр: " + str1);
                                tr.Commit();
                            }
                        }
                    }
                    
                }
                #endregion

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

                #region check if no types in family
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
                #endregion

                #region add parameters in family
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
                        parametersList = familyManager.Parameters;

                        isExist = false;
                        foreach (var st in parametersArray)
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
                                        familyManager.AddParameter(externalDefinition, BuiltInParameterGroup.INVALID, false);
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
                            foreach (var st in paramtersMSKInstAnnotationArray)
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
                #endregion

                #region set values to parameters in all familytypes
                FamilyTypeSet types = familyManager.Types;
                Guid famGuid = Guid.NewGuid();
                string pKeyValue = "";

                using (Transaction trans = new Transaction(doc, "Get Cod"))
                {
                    trans.Start();
                    if (!isAnnotation)
                    {
                        var pKey = familyManager.get_Parameter("Код по классификатору");
                        pKeyValue = familyType.AsString(pKey);
                    }
                    trans.Commit();
                }
                
                foreach (FamilyType type in types)
                {
                    try
                    {
                        using (Transaction t = new Transaction(doc, "Set parameters to current type"))
                        {
                            t.Start();

                            familyManager.CurrentType = type;

                            var p = familyManager.get_Parameter("МСК_Версия Revit");
                            string build = "R" + commandData.Application.Application.VersionNumber.ToString().Remove(0, 2);
                            familyManager.Set(p, build);

                            p = familyManager.get_Parameter("КПСП_GUID семейства");

                            if (type.AsString(p) == "")
                            {
                                familyManager.Set(p, famGuid.ToString());
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
                            }
                            if (vs == "")
                            {
                                familyManager.Set(p, "1");
                            }
                            */

                            p = familyManager.get_Parameter("КПСП_Дата редактирования");

                            DateTime today = DateTime.Now;

                            int.TryParse(today.Day.ToString(), out int dayInt);
                            int.TryParse(today.Month.ToString(), out int monthInt);
                            string sDate = String.Format("{0:D2}-{1:D2}-{2}", dayInt, monthInt, today.Year.ToString());
                            familyManager.Set(p, sDate);

                            p = familyManager.get_Parameter("КПСП_Автор");
                            familyManager.Set(p, User);

                            p = familyManager.get_Parameter("КПСП_Библиотека семейств");
                            familyManager.Set(p, @"K:\Стандарт\ТИМ Семейства\0_Реестр семейств\Реестр.xlsm");

                            #region set parameters with Classificator Cod
                            if (!isAnnotation)
                            {
                                p = familyManager.get_Parameter("Код по классификатору");
                                familyManager.Set(p, pKeyValue);

                                p = familyManager.get_Parameter("КПСП_Дисциплина");
                                familyManager.Set(p, TableEntry.GetDiscipline(pKeyValue));

                                p = familyManager.get_Parameter("КПСП_Категория");
                                familyManager.Set(p, TableEntry.GetCategory(pKeyValue));

                                p = familyManager.get_Parameter("КПСП_Подкатегория");
                                familyManager.Set(p, TableEntry.GetSubCategory(pKeyValue));

                                p = familyManager.get_Parameter("КПСП_Путь к семейству");
                                string pathToFam = TableEntry.GetPathToFamily(pKeyValue).Replace("RXX", build);
                                familyManager.Set(p, pathToFam);

                                p = familyManager.get_Parameter("КПСП_Инструкция");
                                familyManager.Set(p, TableEntry.GetPathToInstruction(pKeyValue));
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
                                        else if (annotationFamily.FamilyCategory.Name.Contains("Головные части уровней"))
                                        {
                                            p = familyManager.get_Parameter("КПСП_Подкатегория");
                                            familyManager.Set(p, "Обозначение уровня");
                                            p = familyManager.get_Parameter("КПСП_Путь к семейству");
                                            familyManager.Set(p, @"K:\Стандарт\ТИМ Семейства\0_Библиотека семейств RXX\1 Общие\Аннотации\Обозначение уровня".Replace("RXX", build));
                                        }
                                        else if (annotationFamily.FamilyCategory.Name.Contains("Заголовки фрагментов"))
                                        {
                                            p = familyManager.get_Parameter("КПСП_Подкатегория");
                                            familyManager.Set(p, "Обозначение фрагмента");
                                            p = familyManager.get_Parameter("КПСП_Путь к семейству");
                                            familyManager.Set(p, @"K:\Стандарт\ТИМ Семейства\0_Библиотека семейств RXX\1 Общие\Аннотации\Обозначение фрагмента".Replace("RXX", build));
                                        }
                                        else if (annotationFamily.FamilyCategory.Name.Contains("Основные надписи"))
                                        {
                                            p = familyManager.get_Parameter("КПСП_Подкатегория");
                                            familyManager.Set(p, "Основные надписи");
                                            p = familyManager.get_Parameter("КПСП_Путь к семейству");
                                            familyManager.Set(p, @"K:\Стандарт\ТИМ Семейства\0_Библиотека семейств RXX\1 Общие\Аннотации\Основные надписи".Replace("RXX", build));
                                        }
                                        else if (annotationFamily.FamilyCategory.Name.Contains("Ссылка на вид"))
                                        {
                                            p = familyManager.get_Parameter("КПСП_Подкатегория");
                                            familyManager.Set(p, "Ссылки на вид");
                                            p = familyManager.get_Parameter("КПСП_Путь к семейству");
                                            familyManager.Set(p, @"K:\Стандарт\ТИМ Семейства\0_Библиотека семейств RXX\1 Общие\Аннотации\Ссылки на вид".Replace("RXX", build));
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

                            p = familyManager.get_Parameter("КПСП_Вложенные семейства");
                            familyManager.Set(p, estObschieSemeistva);

                            t.Commit();
                        }
                    }
                    catch (Exception e)
                    {
                        TaskDialog.Show("Warning 2", e.ToString());
                    }
                }

                #endregion
            }
            else
            {
                TaskDialog.Show("Warning", "Это не семейство, команда работает только в семействе");
            }

            //TaskDialog.Show("Final", log);
            commandData.Application.Application.SharedParametersFilename = Main.FOP_KSP_Path;
            return Result.Succeeded;
        }
    }
}
