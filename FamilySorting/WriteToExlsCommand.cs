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

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    class WriteToExlsCommand : IExternalCommand
    {
        public string FolderPath { get; set; } = "K:\\Подразделения\\ТИМ\\Обмен\\Проект автоматизация семейств\\Реестр семейств";

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
}
