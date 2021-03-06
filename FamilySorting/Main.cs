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
    public class Main : IExternalApplication
    {
        public static string Folder { get; set; } = "K:\\Стандарт\\ТИМ Семейства\\";
        public static string TabName { get; set; } = "КПСП";
        public static string PanelSortName { get; set; } = "Семейства";
        public static string Comment { get; set; } = " ";
        public static string FolderReestrPath { get; set; } = Folder + "0_Реестр семейств\\Админ";
        public static string ReestrPath { get; set; } = Folder + "0_Реестр семейств\\Админ\\Реестр_семейств.xlsx";
        public static string ReestrLocalPath { get; set; } = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\res\\Реестр_семейств.xlsx";
        public static string ReestrMLocalPath { get; set; } = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\res\\Реестр.xlsm";
        public static string FOPPath { get; set; } = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\res\\ФОП.txt";
        public static string FOP_KSP_Path { get; set; } = "K:\\Стандарт\\ТИМ Шаблоны\\КПСП_ФОП.txt";
        public static string FOP_KSP_Local_Path { get; set; } = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\res\\КПСП_ФОП.txt";
        public static string ClassificatorPath { get; set; } = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\res\\Классификатор семейств.txt";
        public static string User { get; set; } = Environment.UserName.ToString();
        public Result OnStartup(UIControlledApplication application)
        {
            List<RibbonPanel> panelList = new List<RibbonPanel>();
            try
            {
                panelList = application.GetRibbonPanels(TabName);
            }
            catch
            {

            }

            if (panelList.Count == 0)
            {
                application.CreateRibbonTab(TabName);
            }
            RibbonPanel panelSort = application.CreateRibbonPanel(TabName, PanelSortName);

            string path = Assembly.GetExecutingAssembly().Location;
            Folder = GetLibFolder();

            var ClassBtnData = new PushButtonData("ClassBtnData", "Назначить\nклассификатор", path, "FamilySorting.ClassCommand")
            {
                ToolTipImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\class-32.png", UriKind.Absolute)),
                ToolTip = "Подгружает файл классификатора для назначения кода классификатора на семейство"
            };
            var ClassBtn = panelSort.AddItem(ClassBtnData) as PushButton;
            ClassBtn.LargeImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\class-32.png", UriKind.Absolute));

            var AddParamsBtnData = new PushButtonData("AddParamsBtnData", "Добавить\nпараметры", path, "FamilySorting.BindingCommand")
            {
                ToolTipImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\sh-32.png", UriKind.Absolute)),
                ToolTip = "Добавляет общие параметры в семейство, которые позволят его каталогизировать"
            };
            //var AddParamsBtn = panelSort.AddItem(AddParamsBtnData) as PushButton;
            AddParamsBtnData.LargeImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\sh-32.png", UriKind.Absolute));

            var ClearBtnData = new PushButtonData("ClearBtnData", "Очистить\nсемейство", path, "FamilySorting.ClearCommand")
            {
                ToolTipImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\duster-32.png", UriKind.Absolute)),
                ToolTip = "Удаляются общие параметры из семейства из группы Свойства модели"
            };
            //var ClearBtn = panelSort.AddItem(ClearBtnData) as PushButton;
            ClearBtnData.LargeImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\duster-32.png", UriKind.Absolute));

            var FolderBtnData = new PushButtonData("FolderBtnData", "Расположение\nбиблиотеки", path, "FamilySorting.FolderCommand")
            {
                ToolTipImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\folder-32.png", UriKind.Absolute)),
                ToolTip = "Установить расположение библиотеки семейств"
            };
            //var FolderBtn = panelSort.AddItem(FolderBtnData) as PushButton;
            FolderBtnData.LargeImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\folder-32.png", UriKind.Absolute));

            SplitButtonData sBtnData = new SplitButtonData("splitButton", "Split");
            SplitButton sBtn = panelSort.AddItem(sBtnData) as SplitButton;
            sBtn.AddPushButton(AddParamsBtnData);
            sBtn.AddPushButton(ClearBtnData);
            sBtn.AddPushButton(FolderBtnData);

            var SaveBtnData = new PushButtonData("SaveBtnData", "Сохранить\nв библиотеку", path, "FamilySorting.SaveCommand")
            {
                ToolTipImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\save-32.png", UriKind.Absolute)),
                ToolTip = "Сохраняет семейство в библиотеку"
            };
            var SaveBtn = panelSort.AddItem(SaveBtnData) as PushButton;
            SaveBtn.LargeImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\save-32.png", UriKind.Absolute));

            //var ExlsBtnData = new PushButtonData("ExlsBtnData", "Записать\nв журнал", path, "FamilySorting.WriteToExlsCommand")
            //{
            //    ToolTipImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\exls-32.png", UriKind.Absolute)),
            //    ToolTip = "Записывает параметры в журнал"
            //};
            //var ExlsBtn = panelSort.AddItem(ExlsBtnData) as PushButton;
            //ExlsBtn.LargeImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\exls-32.png", UriKind.Absolute));

            return Result.Succeeded;
        }
        public static string GetLibFolder()
        {
            string path = Assembly.GetExecutingAssembly().Location;
            string folder = "";
            string familyFolderPath = path.Replace(@"\FamilySorting.dll", "") + @"\res\folder.txt";

            //Open the file to read from.
            using (StreamReader sr = File.OpenText(familyFolderPath))
            {
                string s;
                while ((s = sr.ReadLine()) != null)
                {
                    folder += s;
                }
            }
            return folder;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

       
    }

}
