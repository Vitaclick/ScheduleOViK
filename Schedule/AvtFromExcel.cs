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
  public class AvtFromExcel : IExternalCommand
  {
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
      // Get application and document objects
      UIApplication ui_app = commandData.Application;
      UIDocument ui_doc = ui_app?.ActiveUIDocument;
      Document doc = ui_doc?.Document;

      // Select sheet and range
      var projectInfo = doc.ProjectInformation;
      var sheetName = projectInfo.LookupParameter("AG_Scp_Лист спецификации")?.AsString();
      if (string.IsNullOrEmpty(sheetName))
      {
        TaskDialog.Show("Ошибка параметра", "Параметр информации о проекте\n\"AG_Scp_Лист спецификации\"\nне заполнен, либо отсутствует");
        return Result.Failed;
      }

      var range = $"{sheetName}!A:AX";

      // Create conneciton between user and Spreadsheet
      var spreadsheetId = "10ep8Rjo0Sp0brza3eE66doqYRAd4AegziTJqjTlgW8Y";

      var dbTransferAvt = new TransferDB(spreadsheetId);
      var dataFromSpreadsheet = dbTransferAvt.ReadSheetData(range);

      try
      {
        var opts = new Dictionary<string, object>();
        if (System.Diagnostics.Debugger.IsAttached)
          opts["Debug"] = true;
        ScriptEngine engine = Python.CreateEngine(opts);
        ScriptScope scope = engine.CreateScope();

        engine.GetSysModule().SetVariable("dataFromSpreadsheet", dataFromSpreadsheet.ToArray());
        scope.SetVariable("doc", doc);
        scope.SetVariable("uidoc", ui_doc);
        //engine.ExecuteFile(@"C:\Drive\ARMOPlug\ScheduleGoogle\ScheduleOViK\Schedule\Resources\AvtFromExcel.py", scope);

        string scriptName = Assembly.GetExecutingAssembly().GetName().Name + ".Resources." + "AvtFromExcel.py";
        Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(scriptName);
        if (stream != null)
        {
          string script = new StreamReader(stream).ReadToEnd();
          engine.Execute(script, scope);
        }

        TaskDialog.Show("Всё хорошо", "ОК");

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
