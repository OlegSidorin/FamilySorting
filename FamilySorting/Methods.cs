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

    class Methods
    {

        public void NewGuid(Document doc)
        {

            string fopFilePath = Main.FOPPath;

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
                try
                {
                    using (Transaction t = new Transaction(doc, "Apply"))
                    {
                        t.Start();

                        var p = familyManager.get_Parameter("КПСП_GUID семейства");
                        var guid = Guid.NewGuid().ToString();
                        familyManager.Set(p, guid);

                        p = familyManager.get_Parameter("МСК_Версия семейства");
                        familyManager.Set(p, "1");

                        p = familyManager.get_Parameter("КПСП_Дата редактирования");
                        DateTime today = DateTime.Now;
                        int.TryParse(today.Day.ToString(), out int dayInt);
                        int.TryParse(today.Month.ToString(), out int monthInt);
                        string sDate = String.Format("{0:D2}-{1:D2}-{2}", dayInt, monthInt, today.Year.ToString());
                        familyManager.Set(p, sDate);

                        t.Commit();
                    }
                }
                catch (Exception e)
                {
                    TaskDialog.Show("!Warning!", e.ToString());
                }

            }
            else
            {
                TaskDialog.Show("Warning main", "Это не семейство, команда работает только в семействе");
            }

        }
        public void NewVer(Document doc)
        {

            string fopFilePath = Main.FOPPath;

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
                try
                {
                    using (Transaction t = new Transaction(doc, "Apply"))
                    {
                        t.Start();

                        var p = familyManager.get_Parameter("МСК_Версия семейства");
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

                        p = familyManager.get_Parameter("КПСП_Дата редактирования");
                        DateTime today = DateTime.Now;
                        int.TryParse(today.Day.ToString(), out int dayInt);
                        int.TryParse(today.Month.ToString(), out int monthInt);
                        string sDate = String.Format("{0:D2}-{1:D2}-{2}", dayInt, monthInt, today.Year.ToString());
                        familyManager.Set(p, sDate);

                        t.Commit();
                    }
                }
                catch (Exception e)
                {
                    TaskDialog.Show("!Warning!", e.ToString());
                }

            }
            else
            {
                TaskDialog.Show("Warning main", "Это не семейство, команда работает только в семействе");
            }

        }
        public string GetParameter(Document doc, string str)
        {
            string result = "";

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
                try
                {
                    using (Transaction t = new Transaction(doc, "GET"))
                    {
                        t.Start();
                        var p = familyManager.get_Parameter(str);
                        result = familyType.AsString(p);
                        t.Commit();
                    }
                }
                catch (Exception e)
                {
                    TaskDialog.Show("!Warning!", e.ToString());
                }

            }
            else
            {
                TaskDialog.Show("Warning main", "Это не семейство, команда работает только в семействе");
            }
            return result;
        }


        #region Supporting Methods
        #endregion
    }

    public class FamilyParameters
    {
        public string Ssilka { get; set; }
        public string Guid { get; set; }
        public string Disciplina { get; set; }
        public string Kategoria { get; set; }
        public string Podkaterogia { get; set; }
        public string ImiaFaila { get; set; }
        public string Proizvoditel { get; set; }
        public string Marka { get; set; }
        public string Vlozhennie { get; set; }
        public string DataObnovki { get; set; }
        public string Versia { get; set; }
        public string Avtor { get; set; }
        public string VersiaRevita { get; set; }

       
        public void GetParameters(Document doc)
        {
            FamilyManager familyManager = doc.FamilyManager;
            var familyType = familyManager.CurrentType;

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

            try
            {
                var p = familyManager.get_Parameter("КПСП_Путь к семейству");
                Ssilka = familyType.AsString(p);
                if ((Ssilka == "") || (Ssilka == " "))
                    Ssilka = @"C:\Users\" + Main.User + @"\Documents";
                p = familyManager.get_Parameter("КПСП_GUID семейства");
                Guid = familyType.AsString(p);
                p = familyManager.get_Parameter("КПСП_Дисциплина");
                Disciplina = familyType.AsString(p);
                p = familyManager.get_Parameter("КПСП_Категория");
                Kategoria = familyType.AsString(p);
                p = familyManager.get_Parameter("КПСП_Подкатегория");
                Podkaterogia = familyType.AsString(p);
                if (!isAnnotation)
                {
                    p = familyManager.get_Parameter("МСК_Завод-изготовитель");
                    Proizvoditel = familyType.AsString(p);
                    p = familyManager.get_Parameter("МСК_Обозначение");
                    Marka = familyType.AsString(p);
                }
                else
                {
                    Proizvoditel = "";
                    Marka = "";
                }
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
                TaskDialog.Show("Warning from FamilyParameters.GetParameters", "No parameters");
            }
        }

    }
}
