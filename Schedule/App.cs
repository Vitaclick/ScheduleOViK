using System;
using System.IO;
using System.Reflection;

using Autodesk.Revit.UI;
using System.Windows.Media.Imaging;

namespace Schedule
{
    class App : IExternalApplication
    {
        public static int Security() // Метод по защите программы в зависимости от контрольной даты
        {
            DateTime ControlDate = new DateTime(2020, 12, 30, 12, 0, 0); // Формируем контрольную дату в формате: год, месяц, день, часы, минуты, секунды
            DateTime dt = DateTime.Today; // получаем сегодняшнюю дату в таком же формате
            TimeSpan deltaDate = ControlDate - dt; // вычисляем разницу между контрольной и текущей датой
            string deltaDateDays = deltaDate.Days.ToString();
            int deltaDateDaysInt = Convert.ToInt32(deltaDateDays);
            //string curDate = dt.ToShortDateString(); //Результат: 06.03.2014
            if (deltaDateDaysInt >= 1 && deltaDateDaysInt <= 7)
            {
                TaskDialog.Show("Истекает срок действия программы Спека", String.Concat("До окончания срока действия программы Спека осталось дней: ", deltaDateDays, ".\r\nЧтобы избежать появления данного предупреждения обратитесь к разработчику\r\nили удалите Спеку в панели управления - программы и компоненты"));
            }
            if (deltaDateDaysInt < 1)
            {
                TaskDialog.Show("Истёк срок действия программы Спека", String.Concat("Срок действия программы Спека истёк", ".\r\nЧтобы избежать появления данного предупреждения обратитесь к разработчику\r\nили удалите программу Спека в панели управления - программы и компоненты.\r\nСпасибо за использование программы."));
            }
            return deltaDateDaysInt;
        }

        // define a method that will create our tab and button
        static void AddRibbonPanel(UIControlledApplication application)
        {
            // Create a custom ribbon tab
            String tabName = "Py.Schedule";
            application.CreateRibbonTab(tabName);

            // Add a new ribbon panel
            RibbonPanel ribbonPanel = application.CreateRibbonPanel(tabName, "OV.Schedule");
            RibbonPanel ribbonPanel1 = application.CreateRibbonPanel(tabName, "HVAC Equipment List");
            RibbonPanel ribbonPanel2 = application.CreateRibbonPanel(tabName, "HVAC Parts List");
            RibbonPanel ribbonPanel3 = application.CreateRibbonPanel(tabName, "Electric");


            // Get dll assembly path
            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;

            // create push button for CleanSchedule
            PushButtonData b1Data = new PushButtonData(
                "cmdCleanSchedule",
                "Очистить" + System.Environment.NewLine + "  параметры  ",
                thisAssemblyPath,
                "Schedule.CleanSchedule");

            PushButton pb1 = ribbonPanel.AddItem(b1Data) as PushButton;
            pb1.ToolTip = "Очистка параметров спецификации";
            BitmapImage pb1Image = new BitmapImage(new Uri("pack://application:,,,/Schedule;component/Resources/cleanTableIcon.png"));
            pb1.LargeImage = pb1Image;

            // create push button for GenerateNames
            PushButtonData b2Data = new PushButtonData(
                "cmdGenerateNames",
                "Сгенерировать" + System.Environment.NewLine + "  наименования  ",
                thisAssemblyPath,
                "Schedule.GenerateNames");

            PushButton pb2 = ribbonPanel.AddItem(b2Data) as PushButton;
            pb2.ToolTip = "Генерация наименований воздуховодов, труб, соединительных деталей";
            BitmapImage pb2Image = new BitmapImage(new Uri("pack://application:,,,/Schedule;component/Resources/genNamesIcon.png"));
            pb2.LargeImage = pb2Image;

            // create push button for UpdateSchedule
            PushButtonData b3Data = new PushButtonData(
                "cmdUpdateSchedule",
                "Обновить" + System.Environment.NewLine + "  параметры  ",
                thisAssemblyPath,
                "Schedule.UpdateSchedule");

            PushButton pb3 = ribbonPanel.AddItem(b3Data) as PushButton;
            pb3.ToolTip = "Обновление системных параметров спецификации";
            BitmapImage pb3Image = new BitmapImage(new Uri("pack://application:,,,/Schedule;component/Resources/updateTableIcon.png"));
            pb3.LargeImage = pb3Image;

            // create push button for ExportToExcel
            PushButtonData b4Data = new PushButtonData(
                "cmdExportToExcel",
                "Экспорт" + System.Environment.NewLine + "  в Excel ",
                thisAssemblyPath,
                "Schedule.ExportToExcel");

            PushButton pb4 = ribbonPanel.AddItem(b4Data) as PushButton;
            pb4.ToolTip = "Экспортировать спецификацию в Excel";
            BitmapImage pb4Image = new BitmapImage(new Uri("pack://application:,,,/Schedule;component/Resources/exportIcon.png"));
            pb4.LargeImage = pb4Image;

            // create push button for ImportFromExcel
            PushButtonData b5Data = new PushButtonData(
                "cmdImportFromExcel",
                "Импорт" + System.Environment.NewLine + "  из Excel  ",
                thisAssemblyPath,
                "Schedule.ImportFromExcel");

            PushButton pb5 = ribbonPanel.AddItem(b5Data) as PushButton;
            pb5.ToolTip = "Импортировать данные из Excel";
            BitmapImage pb5Image = new BitmapImage(new Uri("pack://application:,,,/Schedule;component/Resources/importIcon.png"));
            pb5.LargeImage = pb5Image;


            // create push button for ToExcelEqL
            PushButtonData b6Data = new PushButtonData(
                "cmdToExcelEqL",
                "Equip List" + System.Environment.NewLine + "to Excel",
                thisAssemblyPath,
                "Schedule.ToExcelEqL");

            PushButton pb6 = ribbonPanel1.AddItem(b6Data) as PushButton;
            pb6.ToolTip = "Экспортировать данные в базу Equipment List";
            BitmapImage pb6Image = new BitmapImage(new Uri("pack://application:,,,/Schedule;component/Resources/exportIcon.png"));
            pb6.LargeImage = pb6Image;


            // create push button for FromExcelEqL
            PushButtonData b7Data = new PushButtonData(
                "cmdFromExcelEqL",
                "Import" + System.Environment.NewLine + "Equip List",
                thisAssemblyPath,
                "Schedule.FromExcelEqL");

            PushButton pb7 = ribbonPanel1.AddItem(b7Data) as PushButton;
            pb7.ToolTip = "Импортировать данные из Equipment List";
            BitmapImage pb7Image = new BitmapImage(new Uri("pack://application:,,,/Schedule;component/Resources/importIcon.png"));
            pb7.LargeImage = pb7Image;


            // create push button for UpdatePartsList
            PushButtonData b8Data = new PushButtonData(
                "cmdUpdatePartsList",
                "Update" + System.Environment.NewLine + "Parts List",
                thisAssemblyPath,
                "Schedule.UpdatePartsList");

            PushButton pb8 = ribbonPanel2.AddItem(b8Data) as PushButton;
            pb8.ToolTip = "Обновить Parts List";
            BitmapImage pb8Image = new BitmapImage(new Uri("pack://application:,,,/Schedule;component/Resources/updateTableIcon.png"));
            pb8.LargeImage = pb8Image;

          // create push button for AvtToExcel
          PushButtonData b9Data = new PushButtonData(
            "cmdAvtToExcel",
            "Export AVT" + System.Environment.NewLine + "to Excel",
            thisAssemblyPath,
            "Schedule.AvtToExcel");

          PushButton pb9 = ribbonPanel3.AddItem(b9Data) as PushButton;
          pb9.ToolTip = "Экспортировать TAG автоматики в эксель";
          BitmapImage pb9Image = new BitmapImage(new Uri("pack://application:,,,/Schedule;component/Resources/exportIcon.png"));
          pb9.LargeImage = pb9Image;


          // create push button for AvtFromExcel
          PushButtonData b10Data = new PushButtonData(
            "cmdAvtFromExcel",
            "Import AVT" + System.Environment.NewLine + "from Excel",
            thisAssemblyPath,
            "Schedule.AvtFromExcel");

          PushButton pb10 = ribbonPanel3.AddItem(b10Data) as PushButton;
          pb10.ToolTip = "Импортировать TAG автоматики из эксель";
          BitmapImage pb10Image = new BitmapImage(new Uri("pack://application:,,,/Schedule;component/Resources/importIcon.png"));
          pb10.LargeImage = pb10Image;

    }

        public Result OnShutdown(UIControlledApplication application)
        {
            // do nothing
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            // call our method that will load up our toolbar
            AddRibbonPanel(application);
            return Result.Succeeded;
        }
    }
}
