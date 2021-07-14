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
    using System.Diagnostics;
    using OfficeOpenXml;
    using System.Windows.Forms;

    //using Microsoft.Win32;

    class WriteToExls
    {

        public static void Write(UIDocument uiDoc)
        {
            
            Document doc = uiDoc.Document;
            string path = Assembly.GetExecutingAssembly().Location;
            string fopFilePath = Path.GetDirectoryName(path) + "\\res\\ФОП.txt";
            bool polnoyeSovpadenie = false;
            bool allOk = false;
            var filePath = Main.ReestrPath;
            var rsf = new RS_Fields();
            
            if (doc.IsFamilyDocument)
            {
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

                using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(filePath)))
                {
                    //create an instance of the the first sheet in the loaded file
                    //ExcelWorksheet worksheet;
                    foreach (var ws in excelPackage.Workbook.Worksheets)
                    {
                        //add data
                        
                        if (ws.ToString() == "Реестр семейств")
                        {
                            int row = 2;
                            int rowIfExsist = 0;
                            while ((ws.Cells[row, 1].Value != null) & (ws.Cells[row, 2].Value != null) & (ws.Cells[row, 6].Value != null))
                            {
                                if (ws.Cells[row, 1].Value.ToString() == RS_Fields.Guid)
                                    rowIfExsist = row;
                                row += 1;
                            }

                            if (rowIfExsist != 0)
                            {
                                if (ws.Cells[rowIfExsist, 2].Value.ToString() == RS_Fields.Disciplina && 
                                    ws.Cells[rowIfExsist, 3].Value.ToString() == RS_Fields.Kategoria && 
                                    ws.Cells[rowIfExsist, 3].Value.ToString() == RS_Fields.Podkaterogia && 
                                    ws.Cells[rowIfExsist, 3].Value.ToString() == RS_Fields.ImiaFaila
                                    )
                                {
                                    polnoyeSovpadenie = true;
                                    row = rowIfExsist;
                                }
                                
                            }


                            //ws.Cells[row, 1].Value = rsFields.Ssilka;

                            if (polnoyeSovpadenie || (!polnoyeSovpadenie && (rowIfExsist == 0)))
                            {
                                ws.Cells[row, 1].Value = RS_Fields.Guid;
                                ws.Cells[row, 2].Value = RS_Fields.Disciplina;
                                ws.Cells[row, 3].Value = RS_Fields.Kategoria;
                                ws.Cells[row, 4].Value = RS_Fields.Podkaterogia;
                                ws.Cells[row, 5].Value = RS_Fields.ImiaFaila;
                                ws.Cells[row, 6].Value = RS_Fields.Proizvoditel;
                                ws.Cells[row, 7].Value = RS_Fields.Marka;
                                ws.Cells[row, 8].Value = RS_Fields.Vlozhennie;
                                ws.Cells[row, 9].Value = RS_Fields.DataObnovki;
                                ws.Cells[row, 10].Value = RS_Fields.Versia;
                                ws.Cells[row, 11].Value = RS_Fields.Avtor;
                                ws.Cells[row, 12].Value = RS_Fields.VersiaRevita;
                                allOk = true;
                            }
                            else
                            {
                                
                            }

                            

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
                            ws.Cells[row, 1].Value = RS_Fields.Guid;
                            ws.Cells[row, 2].Value = RS_Fields.ImiaFaila;
                            ws.Cells[row, 3].Value = RS_Fields.DataObnovki;
                            ws.Cells[row, 4].Value = RS_Fields.Avtor;
                            ws.Cells[row, 5].Value = Main.Comment;

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
                                        if (ws.Cells[row, 1].Value.ToString() == RS_Fields.Guid)
                                        {
                                            rows.Add(row);
                                        }
                                    }
                                }
                                
                                row += 1;
                            }

                            for (int i = 0; i < guidsDistinct.Count; i++)
                            {
                                ws.Cells[row + i, 1].Value = RS_Fields.Guid;
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
                        if (allOk)
                        {
                            excelPackage.Save();
                        }
                        else
                        {
                            TaskDialog.Show("Warning", "Guid семейства принадлежит другому семейству с другим именем или другой категории. Создайте новый Guid");
                        }
                        
                        
                    }
                    catch
                    {
                        TaskDialog.Show("Warning", "Невозможно произвести запись в файл..");
                    }
                    
                }


            }
            else
            {
                TaskDialog.Show("Warning", "Это не семейство, команда работает только в семействе");
            }
        }


        #region Supporting Methods
        private void WriteToExcel(string path, string excelSheet, string str, int columns)
        {
            using (var package = new ExcelPackage())
            {
                var sb = new StringBuilder(str);
                var csb = new StringBuilder();
                var file = new FileInfo(path);
                var sheet = package.Workbook.Worksheets.Add(excelSheet);
                var punkt = new List<int>();
                punkt.Add(0);
                int dRow = 1;
                int dCol = 0;
                string currentString = "";
                double currentDouble = -1;
                List<int> groupMarkers = new List<int>();



                for (int i = 0; i < sb.Length; i++)
                {
                    char ch = sb[i];
                    if (ch.ToString() == "\n")
                    {
                        string currentCell = CurCell(dRow, dCol);
                        currentString = csb.Replace("\n", "").Replace("\t", "").ToString();
                        if (Double.TryParse(currentString, out currentDouble))
                            sheet.Cells[currentCell].Value = currentDouble;
                        else
                        {
                            sheet.Cells[currentCell].Value = currentString;
                            if (currentString.Contains("М_ИОС") || currentString.Contains("М_КР") || currentString.Contains("М_АР"))
                            {
                                sheet.Cells[dRow, 1, dRow, columns].Style.Font.Bold = true;
                                sheet.Cells[dRow, 1, dRow, columns].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                sheet.Cells[dRow, 1, dRow, columns].Style.Fill.BackgroundColor.SetColor(100, 105, 124, 231);
                                sheet.Cells[dRow, 1, dRow, columns].Style.Font.Color.SetColor(System.Drawing.Color.LightYellow);
                                sheet.Cells[dRow + 1, 1, dRow + 1, columns].Style.Font.Bold = true;
                                sheet.Cells[dRow + 1, 1, dRow + 1, columns].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                sheet.Cells[dRow + 1, 1, dRow + 1, columns].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                sheet.Cells[dRow + 1, 1, dRow + 1, columns].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                sheet.Cells[dRow + 1, 1, dRow + 1, columns].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                groupMarkers.Add(dRow);
                            }
                            if (currentString.Contains("Отчет") || currentString.Contains("Условные"))
                            {
                                sheet.Cells[dRow, 1, dRow, columns].Style.Font.Bold = true;
                                sheet.Cells[dRow, 1, dRow, columns].Style.Font.Color.SetColor(System.Drawing.Color.DarkBlue);
                            }
                        }
                        csb.Clear();
                        dRow += 1;
                        dCol = 0;

                    }
                    if (ch.ToString() == "\t")
                    {
                        string currentCell = CurCell(dRow, dCol);
                        currentString = csb.Replace("\n", "").Replace("\t", "").ToString();

                        if (Double.TryParse(currentString, out currentDouble))
                        {
                            sheet.Cells[currentCell].Value = currentDouble;
                            //sheet.Cells[currentCell].Style.Font.Italic = true;
                            //sheet.Cells[currentCell].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            //sheet.Cells[currentCell].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Aqua);
                            //sheet.Cells[currentCell].Style.Font.Color.SetColor(System.Drawing.Color.Green);
                        }
                        else
                        {
                            sheet.Cells[currentCell].Value = currentString;
                            if (currentString == "")
                            {
                                sheet.Cells[currentCell].Style.Font.Bold = true;
                                sheet.Cells[currentCell].Style.Font.Color.SetColor(System.Drawing.Color.Red);
                            }
                            if (currentString == "")
                            {
                                sheet.Cells[currentCell].Style.Font.Bold = true;
                                sheet.Cells[currentCell].Style.Font.Color.SetColor(System.Drawing.Color.DarkRed);
                                //sheet.Cells[currentCell].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.LightGray;
                                //sheet.Cells[currentCell].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightPink);
                            }
                            if (currentString == "")
                            {
                                sheet.Cells[currentCell].Style.Font.Bold = true;
                                sheet.Cells[currentCell].Style.Font.Color.SetColor(System.Drawing.Color.DarkRed);
                                //sheet.Cells[currentCell].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.LightGray;
                                //sheet.Cells[currentCell].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightCyan);
                            }
                            if (currentString.Contains("!!"))
                            {
                                sheet.Cells[currentCell].Style.Font.Bold = true;
                                sheet.Cells[currentCell].Style.Font.Color.SetColor(System.Drawing.Color.DarkRed);
                                sheet.Cells[currentCell].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                sheet.Cells[currentCell].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                            }
                            if (currentString.Contains("??"))
                            {
                                sheet.Cells[currentCell].Style.Font.Bold = true;
                                sheet.Cells[currentCell].Style.Font.Color.SetColor(System.Drawing.Color.DarkRed);
                                sheet.Cells[currentCell].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.LightDown;
                                sheet.Cells[currentCell].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                            }
                            if (currentString.Contains("МСК_Код по классификатору") || currentString.Contains("Значение Кода"))
                            {
                                sheet.Cells[currentCell].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.LightGray;
                                sheet.Cells[currentCell].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);
                            }
                        }
                        csb.Clear();
                        dCol += 1;
                    }
                    else
                    {
                        csb.Append(ch);
                    }

                }


                /* справочно
				 * '\u0065' A (char)
					" " (ASCII 32 (0x20)), обычный пробел.
					"\t" (ASCII 9 (0x09)), символ табуляции.
					"\n" (ASCII 10 (0x0A)), символ перевода строки.
					"\r" (ASCII 13 (0x0D)), символ возврата каретки.
					"\0" (ASCII 0 (0x00)), NUL-байт.
					"\x0B" (ASCII 11 (0x0B)), вертикальная табуляция.
				*/

                //sheet.Cells["A1"].Value = sb.ToString();

                //sheet.Cells["A2"].Value = row.Count;
                //sheet.Cells["A3"].Value = col.Count;
                //sheet.Cells["A4"].Value = col.First();
                //col.RemoveAt(0);
                //sheet.Cells["A5"].Value = col.First();

                // Save to file

                groupMarkers.Add(dRow - 7); // 7 это вычет из-за Условных обозначений

                int lastMarker = groupMarkers.Count;

                for (var n = 0; n < lastMarker - 1; n++)
                {
                    for (var i = groupMarkers[n] + 2; i < groupMarkers[n + 1] - 1; i++)
                    {
                        sheet.Row(i).OutlineLevel = 1;
                        sheet.Row(i).Collapsed = true;
                    }
                }
                package.SaveAs(file);
            }
        }
        private string CurCell(int row, int col)
        {
            int let = 25;
            string curCell;
            if (col <= let)
                curCell = Convert.ToChar(65 + col).ToString() + (row).ToString();
            else
                curCell = Convert.ToChar(65).ToString() + Convert.ToChar(65 + col - let - 1).ToString() + (row).ToString();
            return curCell;
        }
        private void OpenFolder(string folderPath)
        {

            if (Directory.Exists(folderPath))
            {
                ProcessStartInfo startInfo = new ProcessStartInfo
                {
                    Arguments = folderPath,
                    FileName = "explorer.exe",
                };
                var pr = Process.GetProcessesByName("explorer").Where(xxx => xxx.MainWindowTitle.Contains("Реестр семейств")).ToList();
                string ss = "";
                if (pr != null)
                {
                    //var str = pr.Count.ToString();
                    //TaskDialog.Show("Warning", str);
                    foreach (var p in pr)
                    {
                        try
                        {
                            p.Kill();
                        }
                        catch { }

                    }
                }
                //TaskDialog.Show("Process", ss);
                Process.Start(startInfo);
            }
            else
            {
                TaskDialog.Show("Warning", String.Format("{0} не существует", folderPath));
            }
        }
        #endregion
    }

    public class RS_Fields
    {
        public static string Ssilka { get; set; }
        public static string Guid { get; set; }
        public static string Disciplina { get; set; }
        public static string Kategoria { get; set; }
        public static string Podkaterogia { get; set; }
        public static string ImiaFaila { get; set; }
        public static string Proizvoditel { get; set; }
        public static string Marka { get; set; }
        public static string Vlozhennie { get; set; }
        public static string DataObnovki { get; set; }
        public static string Versia { get; set; }
        public static string Avtor { get; set; }
        public static string VersiaRevita { get; set; }

       
        public static void Get_Fields(Document doc)
        {
            FamilyManager familyManager = doc.FamilyManager;
            var familyType = familyManager.CurrentType;
            try
            {
                var p = familyManager.get_Parameter("КПСП_Путь к семейству");
                Ssilka = familyType.AsString(p);
                p = familyManager.get_Parameter("КПСП_GUID семейства");
                Guid = familyType.AsString(p);
                p = familyManager.get_Parameter("КПСП_Дисциплина");
                Disciplina = familyType.AsString(p);
                p = familyManager.get_Parameter("КПСП_Категория");
                Kategoria = familyType.AsString(p);
                p = familyManager.get_Parameter("КПСП_Подкатегория");
                Podkaterogia = familyType.AsString(p);
                p = familyManager.get_Parameter("МСК_Завод-изготовитель");
                Proizvoditel = familyType.AsString(p);
                p = familyManager.get_Parameter("МСК_Обозначение");
                Marka = familyType.AsString(p);
                p = familyManager.get_Parameter("КПСП_Вложенные семейства");
                Vlozhennie = familyType.AsString(p);
                p = familyManager.get_Parameter("КПСП_Дата редактирования");
                DataObnovki = familyType.AsString(p);
                p = familyManager.get_Parameter("МСК_Версия семейства");
                Versia = familyType.AsString(p);
                ImiaFaila = doc.Title.ToString() + ".rfa";
                p = familyManager.get_Parameter("МСК_Версия Revit");
                VersiaRevita = familyType.AsString(p);
                p = familyManager.get_Parameter("КПСП_Автор");
                Avtor = familyType.AsString(p);
            }
            catch
            {

            }
        }

    }
}
