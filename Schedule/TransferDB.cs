
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
using System.Reflection;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;

namespace Schedule
{
  public class TransferDB
  {
    string[] scopes = { SheetsService.Scope.Spreadsheets };
    public string applicationName = "Current transfer";
//#if DEBUG
//    string spreadsheetId = "1LTUA8C-tIBgieNw7PjDGjkSBvV0_i5vpwnJc2C_sKLk";

    string _spreadsheetId;
    private SheetsService service;
    GoogleCredential credential;

    public TransferDB(string spreadsheetId)
    {
// DEBUG MODE:
//      string connectionSecret = @"D:\#Projects\#REPOS\ScheduleOViK\Schedule\LibTreeSecret_debug.json";
//      using (var stream = new FileStream(connectionSecret, FileMode.Open, FileAccess.Read))

      // get connection credentials from file
      _spreadsheetId = spreadsheetId;
      string connectionSecret = Assembly.GetExecutingAssembly().GetName().Name + ".Resources." + "client_secret.json";

      using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(connectionSecret))
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

      var updateRequest = service.Spreadsheets.Values.Update(valueRange, _spreadsheetId, range);
      updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
      var updateResponse = updateRequest.Execute();
    }

    public void WriteData(string range, List<IList<object>> data)
    {
      var valueRange = new ValueRange();

      valueRange.Values = data;

      var appendRequest = service.Spreadsheets.Values.Append(valueRange, _spreadsheetId, range);

      appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
      var appendResponse = appendRequest.Execute();
    }

    public void WriteColumn(string range, List<IList<object>> data)
    {
      var valueRange = new ValueRange();

      valueRange.Values = data;

      var updateRequest = service.Spreadsheets.Values.Update(valueRange, _spreadsheetId, range);

      updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
      var updateResponse = updateRequest.Execute();

    }

    public List<List<object>> ReadSheetData(string range)
    {
      var request = service.Spreadsheets.Values.Get(_spreadsheetId, range);
      ValueRange response = request.Execute();

      IList<IList<object>> values = response.Values;

      var resultValues = new List<List<object>> { };

      if (values != null && values.Count > 0)
      {
        foreach (var rowValue in values)
        {
          if (rowValue.Count == 0) continue;
          resultValues.Add(rowValue as List<object>);
        }
      }
      return resultValues;
    }

    public List<List<object>> ReadBatchSheetData(string[] ranges)
    {
      // Get values of given range
      SpreadsheetsResource.ValuesResource.BatchGetRequest request = service.Spreadsheets.Values.BatchGet(_spreadsheetId);
      request.Ranges = ranges;

      BatchGetValuesResponse response = request.Execute();

      IList<ValueRange> responses = response.ValueRanges;

      var resultValues = new List<List<object>> { };

      foreach (var res in responses)
      {
        if (res.Values != null)
        {

          var values = res.Values.SelectMany(x => x.DefaultIfEmpty("")).ToList();
          resultValues.Add(values);
        }
      }

      return resultValues;
    }
  }
}
