using System;
using System.IO;
using System.Reflection;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;

using IronPython.Hosting;
using Microsoft.Scripting.Hosting;

namespace Schedule
{
  [Transaction(TransactionMode.Manual)]
  public class ExportToExcel : IExternalCommand
  {
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
      // Create conneciton to DB (закрывать соединение не нужно)
      var dbTransfer = new TransferDB();

      // Get application and document objects
      UIApplication ui_app = commandData.Application;
      UIDocument ui_doc = ui_app?.ActiveUIDocument;
      Document doc = ui_doc?.Document;

      // Select sheet and range
      string sheetName = "Общий";
      var range = $"{sheetName}!A:A";

      try
      {

        ScriptEngine engine = Python.CreateEngine();

        ScriptScope scope = engine.CreateScope();
        scope.SetVariable("doc", doc);
        scope.SetVariable("uidoc", ui_doc);
        scope.SetVariable("uiapp", ui_app);

        string scriptName = Assembly.GetExecutingAssembly().GetName().Name + ".Resources." + "ToExcel.py";
        Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(scriptName);
        if (stream != null)
        {
          string script = new StreamReader(stream).ReadToEnd();
          engine.Execute(script, scope);
        }

        // Import schedule data from IPython
        var revitData = new List<IList<object>>() { };

        var dynamicDataFromPy = scope.GetVariable("revData");

        foreach (var i in dynamicDataFromPy)
        {
          revitData.Add((IList<object>)i);
        }

        // Forming request from spreadsheet
        var sheetBatchValues = dbTransfer.ReadBatchSheetData(new[] { "Общий!A:A", "Общий!C:C", "Общий!D:D" });

        // Compose unique keys for matching with Revit data
        var uniqueSheetKeys = new List<string>() { };
        if (sheetBatchValues.Count > 0)
        {
          for (int i = 0; i < sheetBatchValues.Max(x => x.Count); i++)
          {
            var key = sheetBatchValues[0].ElementAtOrDefault(i) as string + sheetBatchValues[1].ElementAtOrDefault(i) +
                      sheetBatchValues[2].ElementAtOrDefault(i);
            uniqueSheetKeys.Add(key);
          }
        }
        // match revit data values with spreadsheet
        var filteredNewValues = new List<IList<object>> { };

        foreach (var dataRow in revitData)
        {
          // form unique key for revit schedule data
          var uniqueRevitDataKey = dataRow[0] as string + dataRow[2] + dataRow[3];
          if (!uniqueSheetKeys.Contains(uniqueRevitDataKey))
          {
            filteredNewValues.Add(dataRow);
          }
        }

        dbTransfer.WriteData(range, filteredNewValues);

        return Result.Succeeded;
      }
      // This is where we "catch" potential errors and define how to deal with them
      catch (Autodesk.Revit.Exceptions.OperationCanceledException)
      {
        // If user decided to cancel the operation return Result.Canceled
        return Result.Cancelled;
      }
      catch (Exception ex)
      {
        // If something went wrong return Result.Failed
        message = ex.Message;
        return Result.Failed;
      }

    }
  }
}
