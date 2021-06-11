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
        public static string TabName { get; set; } = "КПСП";
        public static string PanelSortName { get; set; } = "Семейства";
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

            var ClassBtnData = new PushButtonData("ClassBtnData", "Назначить\nклассификатор", path, "FamilySorting.ClassCommand")
            {
                ToolTipImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\class.png", UriKind.Absolute)),
                ToolTip = "Подгружает файл классификатора для назначения кода классификатора на семейство"
            };
            var ClassBtn = panelSort.AddItem(ClassBtnData) as PushButton;
            ClassBtn.LargeImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\class-32.png", UriKind.Absolute));

            var SortBtnData = new PushButtonData("SortBtnData", "Добавить\nпараметры", path, "FamilySorting.BindingCommand")
            {
                ToolTipImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\sh.png", UriKind.Absolute)),
                ToolTip = "Добавляет общие параметры в семейство, которые позволят его каталогизировать"
            };
            var SortBtn = panelSort.AddItem(SortBtnData) as PushButton;
            SortBtn.LargeImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\sh-32.png", UriKind.Absolute));

            var ClearBtnData = new PushButtonData("ClearBtnData", "Очистить\nGUID семейства", path, "FamilySorting.ClearCommand")
            {
                ToolTipImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\duster.png", UriKind.Absolute)),
                ToolTip = "Стирает GUID семейства"
            };
            var ClearBtn = panelSort.AddItem(ClearBtnData) as PushButton;
            ClearBtn.LargeImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\duster-32.png", UriKind.Absolute));

            var ExlsBtnData = new PushButtonData("ExlsBtnData", "Записать\nв журнал", path, "FamilySorting.WriteToExlsCommand")
            {
                ToolTipImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\exls.png", UriKind.Absolute)),
                ToolTip = "Записывает параметры в журнал"
            };
            var ExlsBtn = panelSort.AddItem(ExlsBtnData) as PushButton;
            ExlsBtn.LargeImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\exls-32.png", UriKind.Absolute));

            return Result.Succeeded;
        }


        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

       
    }

}
