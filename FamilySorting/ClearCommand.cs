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
            //string originalFile = app.SharedParametersFilename;
            string fopFilePath = Main.FOPPath;
            string log = "";
            bool isExist = false;
            string[] giudKPSPArray =
            {
                "11b18c00-5d82-4226-8b5f-74526a7ec4f8", "3c0744d8-3713-4311-b03e-885f7441d360", "f956074c-276f-4f03-bcdf-890dc4a6038a", 
                "eb3b5f14-25f5-41eb-942f-d8dd33d766c1",
                "37384649-c3c8-4fc2-a08e-c2206438f528", "85cd0032-c9ee-4cd3-8ffa-b2f1a05328e3", "fd97b929-0274-408b-8299-9981cb982fc5",
                "fdf7bfa4-5294-45c5-b979-c388d3a062da", "18ea3aa8-5275-470f-94da-e35bb4c80e46", "728fe6d4-f0e9-4418-b261-25c67382b379",
                "0b3fd4ed-0256-43e5-a997-5311f4c19091", "A80FE9BB-B06E-46BD-B50D-D32486ED228F", "8F0F22FE-8DA1-4CF3-B94D-3DC33041E5D3"
            };
            string[] guidMSKTypeArray =
            {
                "fb30c7d4-3e3c-4fe6-821b-189cf35b7f9f", "647b5bc9-6570-416c-93d3-bd0d159775f2", "a8cdbf7b-d60a-485e-a520-447d2055f351", // последний завод, кот выз ошибку
                "8b5e61a2-b091-491c-8092-0b01a55d4f45", "9b3dbd60-5be3-4842-9dbe-cd644ef5f9e8", "946c4e27-a56c-422d-999c-778a150b950e",
                "a8832df7-0302-4a63-a6e1-47a01632b987", "8f2e4f93-9472-4941-a65d-0ac468fd6a6d", "da753fe3-ecfa-465b-9a2c-02f55d0c2ff1",
                "ef3ac60d-2cf8-4bd8-bd66-dbcb42e92f4a",
                "f13b35e5-9fb9-4cf8-b330-efe01d3780c4", "e7edd112-da46-46c3-886c-934dad841efb"
            };
            string[] guidMSKInstArray =
            {
                "bfa2f0d2-ccd0-4a02-95c7-573f0a9829c3", "2fd9e8cb-84f3-4297-b8b8-75f444e124ed",  "ae8ff999-1f22-4ed7-ad33-61503d85f0f4"
            };
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
                        familyType = familyManager.NewType("Тип 1");
                        familyManager.CurrentType = familyType;
                        t.Commit();
                    }
                }

                #region clear 
                //TaskDialog.Show("Warning", "Privet");
                try
                {
                    commandData.Application.Application.SharedParametersFilename = fopFilePath;
                    using (Transaction t = new Transaction(doc,"Clear"))
                    {
                        t.Start();
                        FamilyParameterSet parametersList = familyManager.Parameters;

                        foreach (var guid in giudKPSPArray)
                        {
                            try
                            {
                                var p = familyManager.get_Parameter(new Guid(guid)); 
                                familyManager.RemoveParameter(p);
                            }
                            catch
                            {
                                
                            }
                        }
                        //foreach (var guid in guidMSKTypeArray)
                        //{
                        //    try
                        //    {
                        //        var p = familyManager.get_Parameter(new Guid(guid));
                        //        familyManager.RemoveParameter(p);
                        //    }
                        //    catch
                        //    {
                                
                        //    }
                        //}
                        //foreach (var guid in guidMSKInstArray)
                        //{
                        //    try
                        //    {
                        //        var p = familyManager.get_Parameter(new Guid(guid)); 
                        //        familyManager.RemoveParameter(p);
                        //    }
                        //    catch
                        //    {
                                
                        //    }
                        //}
                        t.Commit();
                    }
                }
                catch (Exception e)
                {
                    TaskDialog.Show("Warning 1", e.ToString());
                }

                #endregion

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
