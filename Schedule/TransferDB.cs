
using System;
using System.Collections;
using System.IO;
using System.Threading;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using Google.Apis.Services;
using Google.Apis.Auth.OAuth2;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;

namespace Schedule
{
  public class TransferDB
  {
    string[] scopes = { SheetsService.Scope.Spreadsheets };
    public string applicationName = "Current transfer";
    string spreadsheetId = "16OOlFqawtSqN3xgl-Kn4VkdyTFKba53nXrFPvubDjR0";
    private SheetsService service;
    GoogleCredential credential;

    public TransferDB()
    {
      // !!!!!!!! CHANGE TO RESOURCES !!!!!!!!!!!!!!!
      using (var stream = new FileStream(@"D:\#Projects\#REPOS\ScheduleOViK\Schedule\client_secret.json", FileMode.Open, FileAccess.Read))
      {
        credential = GoogleCredential.FromStream(stream).CreateScoped(scopes);
      }
      // Create Google Sheets API service
      service = new SheetsService(new BaseClientService.Initializer()
      {
        HttpClientInitializer = credential,
        ApplicationName = applicationName
      });
    }
    public void UpdateSpreadsheet(string range, List<IList<object>> data)
    {
      range = "Общий!B2";
      var valueRange = new ValueRange();
      valueRange.MajorDimension = "ROWS";

      valueRange.Values = new List<IList<object>> { new List<object> { "updated", "keke" } };

      var updateRequest = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, range);
      updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
      var updateResponse = updateRequest.Execute();
    }

    public void WriteData(string range, List<IList<object>> data)
    {
      var valueRange = new ValueRange();

      valueRange.Values = data;

      var appendRequest = service.Spreadsheets.Values.Append(valueRange, spreadsheetId, range);

      appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
      var appendResponse = appendRequest.Execute();
    }

    public List<List<object>> ReadData(string[] ranges)
    {
      // Get values of given range
      SpreadsheetsResource.ValuesResource.BatchGetRequest request = service.Spreadsheets.Values.BatchGet(spreadsheetId);
      request.Ranges = ranges;


      //        SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);

      BatchGetValuesResponse response = request.Execute();


      IList<IList<object>> values = response.ValueRanges.SelectMany(x => x.Values).ToList();
      var outValues = new List<List<object>> { };

      if (values != null && values.Count > 0)
      {
        foreach (var row in values)
        {
          outValues.Add(row as List<object>);
        }
        return outValues;
      }
      else
      {
        new TaskDialog("Нет данных для импорта");
        return new List<List<object>>();
      }
    }
  }
}
