using System;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ClosedXML.Excel;
using System.IO;
using System.Windows.Forms;

namespace HamitScheduleSync
{
    [Transaction(TransactionMode.Manual)]
    public class ImportScheduleCommand : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Prompt user for the Excel file
            string filePath = "";
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Dosyası (*.xlsx)|*.xlsx";
                openFileDialog.Title = "İçe Aktarılacak Excel Dosyasını Seçin";

                if (openFileDialog.ShowDialog() != DialogResult.OK)
                {
                    return Result.Cancelled;
                }
                filePath = openFileDialog.FileName;
            }

            try
            {
                using (var workbook = new XLWorkbook(filePath))
                {
                    var worksheet = workbook.Worksheet(1); // Assuming first sheet
                    var rows = worksheet.RangeUsed().RowsUsed().Skip(1); // Skip header row

                    int lastColIdx = worksheet.LastColumnUsed().ColumnNumber();

                    using (Transaction t = new Transaction(doc, "Excel'den Veri Al (Hamit)"))
                    {
                        t.Start();

                        int successCount = 0;
                        int failCount = 0;

                        foreach (var row in rows)
                        {
                            // In Export, we put ElementId in the last column
                            string elementIdStr = row.Cell(lastColIdx).GetString();
                            if (string.IsNullOrEmpty(elementIdStr)) continue;

#if REVIT2024 || REVIT2025 || REVIT2026 // Adjust based on your API Version, we use long for newer API ElementIds
                            if (long.TryParse(elementIdStr, out long elementIdLong))
                            {
                                ElementId eid = new ElementId(elementIdLong);
#else
                            if (int.TryParse(elementIdStr, out int elementId))
                            {
                                ElementId eid = new ElementId(elementId);
#endif
                            
                                Element elem = doc.GetElement(eid);

                                if (elem != null)
                                {
                                    // Normally, here we would map standard headers back to Revit parameters.
                                    // For now, this is a placeholder implementation that tries to match
                                    // column headers to parameter names and updates string parameters.
                                    
                                    bool elementUpdated = false;
                                    for(int c = 1; c < lastColIdx; c++) 
                                    {
                                        string headerName = worksheet.Cell(1, c).GetString();
                                        string cellValue = row.Cell(c).GetString();
                                        
                                        // Attempt to find a parameter with this name
                                        Parameter param = elem.LookupParameter(headerName);

                                        // Sometimes it's a type parameter
                                        if (param == null)
                                        {
                                            ElementId typeId = elem.GetTypeId();
                                            if (typeId != ElementId.InvalidElementId)
                                            {
                                                Element typeElem = doc.GetElement(typeId);
                                                if (typeElem != null)
                                                {
                                                    param = typeElem.LookupParameter(headerName);
                                                }
                                            }
                                        }

                                        if (param != null && !param.IsReadOnly)
                                        {
                                            try 
                                            {
                                                switch(param.StorageType)
                                                {
                                                    case StorageType.String:
                                                        if (param.AsString() != cellValue) 
                                                        {
                                                            param.Set(cellValue);
                                                            elementUpdated = true;
                                                        }
                                                        break;
                                                    case StorageType.Integer:
                                                        if (int.TryParse(cellValue, out int intVal) && param.AsInteger() != intVal)
                                                        {
                                                            param.Set(intVal);
                                                            elementUpdated = true;
                                                        }
                                                        break;
                                                    case StorageType.Double:
                                                        if (double.TryParse(cellValue, out double dblVal))
                                                        {
                                                            // Usually Need to convert display units to internal units, but skipping for simplicity in basic string maps
                                                            // param.Set(dblVal); 
                                                        }
                                                        break;
                                                }
                                            }
                                            catch { /* Ignore individual param failures */ }
                                        }
                                    }

                                    if(elementUpdated)
                                        successCount++;
                                }
                                else
                                {
                                    failCount++;
                                }
                            }
                        }

                        t.Commit();
                        Autodesk.Revit.UI.TaskDialog.Show("Başarılı", $"{successCount} eleman güncellendi. {failCount} eleman bulunamadı.");
                    }
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                Autodesk.Revit.UI.TaskDialog.Show("Hata", "İçe aktarma sırasında bir sorun oluştu: " + ex.Message);
                return Result.Failed;
            }
        }
    }
}
