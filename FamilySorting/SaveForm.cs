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
using OfficeOpenXml;

namespace FamilySorting
{
    public partial class SaveForm : System.Windows.Forms.Form
    {
        public Document Doc { get; set; }
        public SaveButtonExternalEvent saveButtonExternalEvent;
        public ExternalEvent externalEvent;
        public SaveForm()
        {
            InitializeComponent();
            saveButtonExternalEvent = new SaveButtonExternalEvent();
            externalEvent = ExternalEvent.Create(saveButtonExternalEvent);
        }

        private void saveButton_Click(object sender, EventArgs e)
        {

            // поиск нарушения полного совпадения guid - дисциплина - категория - подкатегория - имя файла
            // если будет не полное совпадение тогда guid новый назначен будет
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

            TaskDialog.Show("Warn", Doc.Title + " : " + guidProblem.ToString());

            if (guidProblem)
                externalEvent.Raise();

            TaskDialog.Show("Warn", Doc.Title + " : " + "точка 2 после установки нового guid");

            SaveAsOptions sao = new SaveAsOptions();
            sao.OverwriteExistingFile = true;
            if (!Directory.Exists(labelPath.Text.ToString() + @"\"))
            {
                Directory.CreateDirectory(labelPath.Text.ToString() + @"\");
            }

            Doc.SaveAs(labelPath.Text.ToString() + @"\" + textFamilyName.Text.ToString() + ".rfa", sao);

            var filePath = Main.ReestrPath;
            var f = new FamilyParameters();
            f.GetParameters(Doc);

            IList<FamilyInstance> vlozhennieSemeistva = new FilteredElementCollector(Doc).OfClass(typeof(FamilyInstance)).WhereElementIsNotElementType().Cast<FamilyInstance>().ToList();
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
                        ws.Cells[row, 5].Value = textComment.Text.ToString();
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



            Close(); 
        }

        public class SaveButtonExternalEvent : IExternalEventHandler
        {
            public void Execute(UIApplication app)
            {
                UIDocument uidoc = app.ActiveUIDocument;
                Document doc = uidoc.Document;
                Methods.NewGuid(doc);
                //TaskDialog.Show("Warning", "Iz External Eventa");
                return;
            }

            public string GetName()
            {
                return "External Event On Save Button";
            }
        }
    }
}
