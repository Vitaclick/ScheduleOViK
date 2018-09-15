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
      var range = $"{sheetName}!A:M";

      try
      {
        Dictionary<string, object> options = new Dictionary<string, object> { ["Debug"] = true };

        ScriptEngine engine = Python.CreateEngine(options);

        ScriptSource script = engine.CreateScriptSourceFromFile(@"D:\#Projects\#REPOS\ScheduleOViK\Schedule\Scripts\OVToExcel.py");
        ScriptScope scope = engine.CreateScope();
        scope.SetVariable("doc", doc);
        scope.SetVariable("uidoc", ui_doc);
        scope.SetVariable("uiapp", ui_app);

        dynamic res = script.Execute(scope);


        var data = new List<IList<object>>() { };

        var dynamicDataFromPy = scope.GetVariable("revData");

        foreach (var i in dynamicDataFromPy)
        {
          data.Add((IList<object>)i);
        }


        //        data.Add(dynamicDataFromPy);


        // dbTransfer.WriteData(range, data);
        //                string scriptName = Assembly.GetExecutingAssembly().GetName().Name + ".Resources." + "ToExcel.py";
        //                Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(scriptName);
        //                if (stream != null)
        //                {
        //                    string script = new StreamReader(stream).ReadToEnd();
        //                    engine.Execute(script, scope);
        //                }

        var sheetValues = dbTransfer.ReadData(new[] { "Общий!A:A", "Общий!C:C", "Общий!D:D" });
        // flatten array
        foreach (var row in sheetValues)
        {
          
        }
        var keySheetValues = sheetValues.SelectMany(x => x).Distinct();
        // match data values with spreadsheet
        var filteredNewValues = new List<IList<object>> { };

        foreach (var dataRow in data)
        {
          if (!keySheetValues.Contains(dataRow[0]))
          {
            filteredNewValues.Add(dataRow);
          }
        }


        // transfer.UpdateSpreadsheet(range, data);
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
