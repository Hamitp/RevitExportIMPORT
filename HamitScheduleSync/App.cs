using System;
using System.Reflection;
using System.Windows;
using System.Windows.Media;
using Autodesk.Revit.UI;

namespace HamitScheduleSync
{
    public class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            // Create a custom Ribbon tab
            string tabName = "Hamit";
            try
            {
                application.CreateRibbonTab(tabName);
            }
            catch (Exception)
            {
                // Tab might already exist
            }

            // Create a Ribbon panel
            RibbonPanel panel = application.CreateRibbonPanel(tabName, "Schedule Sync");

            // Get the assembly path
            string assemblyPath = Assembly.GetExecutingAssembly().Location;

            // Create push buttons
            PushButtonData btnExportData = new PushButtonData(
                "cmdExportSchedule",
                "Excel'e Aktar",
                assemblyPath,
                "HamitScheduleSync.ExportScheduleCommand");

            btnExportData.ToolTip = "Aktif Metraj (Schedule) tablosunu yüksek kaliteli bir Excel (*.xlsx) dosyasına aktarır.";
            btnExportData.LargeImage = CreateExcelIconWPF(true);

            PushButtonData btnImportData = new PushButtonData(
                "cmdImportSchedule",
                "Excel'den Al",
                assemblyPath,
                "HamitScheduleSync.ImportScheduleCommand");
            
            btnImportData.ToolTip = "Önceden aktarılmış ve güncellenmiş Excel dosyasındaki verileri modele geri yazar.";
            btnImportData.LargeImage = CreateExcelIconWPF(false);

            // Add buttons to panel
            panel.AddItem(btnExportData);
            panel.AddItem(btnImportData);

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        private ImageSource CreateExcelIconWPF(bool isExport)
        {
            var drawingGroup = new DrawingGroup();
            using (var dc = drawingGroup.Open())
            {
                // Background transparent container to lock the size to exactly 32x32
                dc.DrawRectangle(Brushes.Transparent, null, new Rect(0, 0, 32, 32));

                // Excel Green Rounded Rectangle
                var baseColor = Color.FromRgb(33, 115, 70);
                var gradColor = Color.FromRgb(50, 180, 100);
                var bgBrush = new LinearGradientBrush(gradColor, baseColor, new Point(0, 0), new Point(1, 1));
                
                dc.DrawRoundedRectangle(bgBrush, null, new Rect(2, 2, 28, 28), 4, 4);

                // Draw "X"
                var xPen = new Pen(Brushes.White, 3.5);
                xPen.StartLineCap = PenLineCap.Round;
                xPen.EndLineCap = PenLineCap.Round;

                double xOffset = isExport ? 0 : 8; // shift X to make room for arrow
                
                dc.DrawLine(xPen, new Point(8 + xOffset, 10), new Point(16 + xOffset, 22));
                dc.DrawLine(xPen, new Point(16 + xOffset, 10), new Point(8 + xOffset, 22));

                // Draw Arrow
                var arrPen = new Pen(Brushes.White, 3);
                arrPen.StartLineCap = PenLineCap.Round;
                arrPen.EndLineCap = PenLineCap.Triangle;
                
                if (isExport)
                {
                    // Outward arrow
                    dc.DrawLine(arrPen, new Point(16, 16), new Point(27, 16));
                }
                else 
                {
                    // Inward arrow
                    dc.DrawLine(arrPen, new Point(27, 16), new Point(16, 16));
                }
            }
            
            return new DrawingImage(drawingGroup);
        }
    }
}
