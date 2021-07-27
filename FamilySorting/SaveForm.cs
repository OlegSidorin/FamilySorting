using System;
using System.IO;
using System.Collections.Generic;
using System.Threading;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using OfficeOpenXml;

namespace FamilySorting
{
    public partial class SaveForm : System.Windows.Forms.Form
    {
        public Document Doc { get; set; }

        public static string Comment { get; set; } = "";
        public static string SaveFilePath { get; set; } = "";

        public SaveButtonExternalEventNewGuid saveButtonExternalEventNewGuid;
        public ExternalEvent externalEventNewGuid;

        public SaveButtonExternalEventSaveExls saveButtonExternalEventSaveExls;
        public ExternalEvent externalEventSaveExls;

        public SaveButtonExternalEventNewVer saveButtonExternalEventNewVer;
        public ExternalEvent externalEventNewVer;

        public SaveButtonExternalEventSaveFile saveButtonExternalEventSaveFile;
        public ExternalEvent externalEventSaveFile;

        public SaveForm()
        {
            InitializeComponent();
            saveButtonExternalEventNewGuid = new SaveButtonExternalEventNewGuid();
            externalEventNewGuid = ExternalEvent.Create(saveButtonExternalEventNewGuid);
            saveButtonExternalEventSaveExls = new SaveButtonExternalEventSaveExls();
            externalEventSaveExls = ExternalEvent.Create(saveButtonExternalEventSaveExls);
            saveButtonExternalEventNewVer = new SaveButtonExternalEventNewVer();
            externalEventNewVer = ExternalEvent.Create(saveButtonExternalEventNewVer);
            saveButtonExternalEventSaveFile = new SaveButtonExternalEventSaveFile();
            externalEventSaveFile = ExternalEvent.Create(saveButtonExternalEventSaveFile);
        }

        private void saveButton_Click(object sender, EventArgs e)
        {

            #region find problrm with guid
            bool guidProblem = false;
            var fp = new FamilyParameters();
            fp.GetParameters(Doc);

            using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(Main.ReestrPath)))
            {
                foreach (var ws in excelPackage.Workbook.Worksheets)
                {
                    if (ws.ToString() == "Реестр семейств")
                    {
                        int row = 2;
                        while ((ws.Cells[row, 1].Value != null) & (ws.Cells[row, 2].Value != null) & (ws.Cells[row, 6].Value != null))
                        {
                            if (ws.Cells[row, 1].Value.ToString() == fp.Guid)
                            {
                                if (ws.Cells[row, 2].Value.ToString() == fp.Disciplina &&
                                    ws.Cells[row, 3].Value.ToString() == fp.Kategoria &&
                                    ws.Cells[row, 4].Value.ToString() == fp.Podkaterogia &&
                                    ws.Cells[row, 5].Value.ToString().Replace(".rfa", "") == textFamilyName.Text.ToString())
                                {
                                    guidProblem = false;
                                }
                                else
                                {
                                    guidProblem = true;
                                }
                            }
                            row += 1;
                        }
                    }
                }
            }
            #endregion

            #region new guid or simply new ver
            if (guidProblem)
            {
                externalEventNewGuid.Raise();
            }
            else
            {
                externalEventNewVer.Raise();
            }
            #endregion

            SaveForm.Comment = textComment.Text.ToString();
                
            if (!(fp.Ssilka == @"C:\Users\" + Main.User + @"\Documents"))
            {
                externalEventSaveExls.Raise();
            }

            
            SaveFilePath = labelPath.Text.ToString() + @"\" + textFamilyName.Text.ToString() + ".rfa";
            SaveAsOptions sao = new SaveAsOptions();
            sao.OverwriteExistingFile = true;
            if (!Directory.Exists(labelPath.Text.ToString() + @"\"))
            {
                Directory.CreateDirectory(labelPath.Text.ToString() + @"\");
            }

            Doc.SaveAs(SaveFilePath, sao);

            //}
            externalEventSaveFile.Raise();
            Close(); 
        }

        public class SaveButtonExternalEventNewGuid : IExternalEventHandler
        {
            public void Execute(UIApplication app)
            {
                UIDocument uidoc = app.ActiveUIDocument;
                Document doc = uidoc.Document;
                Methods methods = new Methods();
                methods.NewGuid(doc);
                //TaskDialog.Show("Warning", "Iz External Eventa");
                return;
            }

            public string GetName()
            {
                return "External Event On Save Button";
            }
        }
        public class SaveButtonExternalEventNewVer : IExternalEventHandler
        {
            public void Execute(UIApplication app)
            {
                UIDocument uidoc = app.ActiveUIDocument;
                Document doc = uidoc.Document;
                Methods methods = new Methods();
                methods.NewVer(doc);
                //TaskDialog.Show("Warning", "Iz External Eventa");
                return;
            }

            public string GetName()
            {
                return "External Event On Save Button";
            }
        }
        public class SaveButtonExternalEventSaveExls : IExternalEventHandler
        {
            public void Execute(UIApplication app)
            {
                UIDocument uidoc = app.ActiveUIDocument;
                Document doc = uidoc.Document;
                Methods methods = new Methods();

                IList<FamilyInstance> vlozhennieSemeistva = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).WhereElementIsNotElementType().Cast<FamilyInstance>().ToList();
                List<string> guids = new List<string>();
                List<string> guidsDistinct = new List<string>();
                int guidsVsego = 0;

                foreach (FamilyInstance fi in vlozhennieSemeistva)
                {
                    FamilySymbol fType = fi.Symbol;
                    Family fam = fType.Family;
                    try
                    {
                        string str = fType.get_Parameter(new Guid("11b18c00-5d82-4226-8b5f-74526a7ec4f8")).AsString();
                        guids.Add(str);
                        guidsVsego += 1;
                    }
                    catch
                    {
                        // do nothing
                    }
                    guidsDistinct = guids.Distinct().ToList();
                }

                var filePath = "";
                if (File.Exists(Main.ReestrPath))
                {
                    filePath = Main.ReestrPath;
                }
                else
                {
                    string newdir= Main.Folder + @"\0_Реестр семейств\Админ";
                    string newfile = newdir + @"\Реестр_семейств.xlsx";
                    Directory.CreateDirectory(newdir);
                    File.Copy(Main.ReestrLocalPath, newfile, true);
                    filePath = newfile;
                }
                
                var f = new FamilyParameters();
                f.GetParameters(doc);

                using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(filePath)))
                {

                    foreach (var ws in excelPackage.Workbook.Worksheets)
                    {
                        if (ws.ToString() == "Реестр семейств")
                        {
                            int row = 2;
                            int rowIfExsist = 0;
                            while ((ws.Cells[row, 1].Value != null) & (ws.Cells[row, 2].Value != null) & (ws.Cells[row, 6].Value != null))
                            {
                                if (ws.Cells[row, 1].Value.ToString() == f.Guid)
                                    rowIfExsist = row;
                                row += 1;
                            }
                            if (rowIfExsist != 0)
                            {
                                if (ws.Cells[rowIfExsist, 2].Value.ToString() == f.Disciplina &&
                                    ws.Cells[rowIfExsist, 3].Value.ToString() == f.Kategoria &&
                                    ws.Cells[rowIfExsist, 4].Value.ToString() == f.Podkaterogia &&
                                    ws.Cells[rowIfExsist, 5].Value.ToString() == f.ImiaFaila
                                    )
                                {
                                    row = rowIfExsist;
                                    //TaskDialog.Show("Warn", "Sovpalo");
                                }

                            }

                            ws.Cells[row, 1].Value = f.Guid;
                            ws.Cells[row, 2].Value = f.Disciplina;
                            ws.Cells[row, 3].Value = f.Kategoria;
                            ws.Cells[row, 4].Value = f.Podkaterogia;
                            ws.Cells[row, 5].Value = f.ImiaFaila;
                            ws.Cells[row, 6].Value = f.Proizvoditel;
                            ws.Cells[row, 7].Value = f.Marka;
                            ws.Cells[row, 8].Value = f.Vlozhennie;
                            ws.Cells[row, 9].Value = f.DataObnovki;
                            ws.Cells[row, 10].Value = f.Versia;
                            ws.Cells[row, 11].Value = f.Avtor;
                            ws.Cells[row, 12].Value = f.VersiaRevita;
                        }
                    }

                    foreach (var ws in excelPackage.Workbook.Worksheets)
                    {
                        //add data
                        if (ws.ToString() == "История изменений")
                        {
                            int row = 2;
                            while ((ws.Cells[row, 1].Value != null) & (ws.Cells[row, 2].Value != null))
                            {
                                row += 1;
                            }
                            ws.Cells[row, 1].Value = f.Guid;
                            ws.Cells[row, 2].Value = f.ImiaFaila;
                            ws.Cells[row, 3].Value = f.DataObnovki;
                            ws.Cells[row, 4].Value = f.Avtor;
                            ws.Cells[row, 5].Value = SaveForm.Comment;
                        }
                    }

                    foreach (var ws in excelPackage.Workbook.Worksheets)
                    {
                        //add data
                        if (ws.ToString() == "Связь вложенных семейств")
                        {
                            int row = 2;
                            List<int> rows = new List<int>();
                            while ((ws.Cells[row, 1].Value != null) & (ws.Cells[row, 2].Value != null))
                            {
                                foreach (var guid in guidsDistinct)
                                {
                                    if (ws.Cells[row, 2].Value.ToString() == guid)
                                    {
                                        if (ws.Cells[row, 1].Value.ToString() == f.Guid)
                                        {
                                            rows.Add(row);
                                        }
                                    }
                                }

                                row += 1;
                            }

                            for (int i = 0; i < guidsDistinct.Count; i++)
                            {
                                ws.Cells[row + i, 1].Value = f.Guid;
                                ws.Cells[row + i, 2].Value = guidsDistinct[i];
                                ws.Cells[row + i, 3].Value = string.Join(", ", rows);
                            }

                            if (rows.Count != 0)
                            {
                                int shift = 0;
                                foreach (var r in rows)
                                {
                                    ws.DeleteRow(r - shift, 1, true);
                                    shift += 1;
                                }
                            }

                        }
                    }
                    //save the changes
                    try
                    {
                        excelPackage.Save();
                    }
                    catch
                    {
                        TaskDialog.Show("Warning", "Невозможно произвести запись в файл..");
                    }

                }

                

                return;
            }

            public string GetName()
            {
                return "External Event On Save Exls";
            }
        }
        public class SaveButtonExternalEventSaveFile : IExternalEventHandler
        {
            public void Execute(UIApplication app)
            {
                UIDocument uidoc = app.ActiveUIDocument;
                Document doc = uidoc.Document;
                Methods methods = new Methods();
                methods.SaveFile(doc);
                return;
            }

            public string GetName()
            {
                return "External Event On Save Button";
            }
        }
    }
}
