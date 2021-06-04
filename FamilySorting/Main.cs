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
        public static string PanelSortName { get; set; } = "Сортировка";
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

            var SortBtnData = new PushButtonData("SortBtnData", "Параметры\nдля сортировки", path, "FamilySorting.SortingCommand")
            {
                ToolTipImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\sort-32.png", UriKind.Absolute)),
                ToolTip = "Внедряет общие параметры в семейство, которые позволят его каталогизировать"
            };
            var SortBtn = panelSort.AddItem(SortBtnData) as PushButton;
            SortBtn.LargeImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\sort-32.png", UriKind.Absolute));

            return Result.Succeeded;
        }


        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

       
    }

}
