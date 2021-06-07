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

            var SortBtnData = new PushButtonData("SortBtnData", "Внедрить\nпараметры", path, "FamilySorting.SortingCommand")
            {
                ToolTipImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\sh.png", UriKind.Absolute)),
                ToolTip = "Внедряет общие параметры в семейство, которые позволят его каталогизировать"
            };
            var SortBtn = panelSort.AddItem(SortBtnData) as PushButton;
            SortBtn.LargeImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\sh-32.png", UriKind.Absolute));

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
