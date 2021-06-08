﻿namespace FamilySorting
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
                "МСК_Версия Revit", "МСК_Версия семейства", "КПСП_GUID семейства", "КПСП_Категория", "КПСП_Подкатегория", "КПСП_Статус",  
                "КПСП_Библиотека семейств", "КПСП_Инструкция", "КПСП_Путь к семейству", "КПСП_Дисциплина", "КПСП_Дата редактирования"
            };
            string[] paramtersMSKArray =
            {
                "МСК_ЕдИзм", "МСК_Завод-изготовитель", "МСК_Код изделия", "МСК_Масса_Текст", "МСК_Наименование", "МСК_Обозначение", "МСК_Позиция", "МСК_Примечание"
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
                        familyType = familyManager.NewType("Обобщенные модели");
                        familyManager.CurrentType = familyType;
                        t.Commit();
                    }
                }

                try
                {
                    commandData.Application.Application.SharedParametersFilename = fopFilePath;
                    using (Transaction t = new Transaction(doc,"Add paramter"))
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
                                familyManager.AddParameter(externalDefinition, BuiltInParameterGroup.PG_DATA, false);
                                log += "\nВнедрен параметр <" + st + ">";
                            }
                            isExist = false;
                        }
                        isExist = false;
                        foreach (var st in paramtersMSKArray)
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
                                familyManager.AddParameter(externalDefinition, BuiltInParameterGroup.PG_TEXT, false);
                                log += "\nВнедрен параметр <" + st + ">";
                            }
                            isExist = false;
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
                    using (Transaction t = new Transaction(doc, "Apply"))
                    {
                        t.Start();
                        log += "\n---";

                        var p = familyManager.get_Parameter("МСК_Версия Revit");
                        string build = commandData.Application.Application.VersionNumber.ToString();
                        familyManager.Set(p, build);
                        log += "\nНовое значение <МСК_Версия Revit>: " + build;

                        p = familyManager.get_Parameter("КПСП_GUID семейства");

                        if (familyType.AsString(p) == "")
                        {
                            Guid famGuid = Guid.NewGuid();
                            familyManager.Set(p, famGuid.ToString());
                            log += "\nПрисвоили новый Guid семейству: " + familyType.AsString(p);
                        }

                        p = familyManager.get_Parameter("КПСП_Дата редактирования");

                        DateTime today = DateTime.Now;

                        int.TryParse(today.Day.ToString(), out int dayInt);
                        int.TryParse(today.Month.ToString(), out int monthInt);
                        string sDate = String.Format("{0:D2}-{1:D2}-{2}", dayInt, monthInt, today.Year.ToString());
                        familyManager.Set(p, sDate);
                        log += "\nДата редактирования: " + familyType.AsString(p);

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

            TaskDialog.Show("Final", log);
            return Result.Succeeded;
        }
    }
}