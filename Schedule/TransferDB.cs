
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

    public void WriteData(string range, List<IList<object>> data)
    {
      var valueRange = new ValueRange();

      valueRange.Values = data;

      var appendRequest = service.Spreadsheets.Values.Append(valueRange, spreadsheetId, range);

      appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
      var appendResponse = appendRequest.Execute();
    }

    public IList<IList<object>> ReadData(string range)
    {
      // Get values of given range
      SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);
      var response = request.Execute();
      IList<IList<object>> values = response.Values;
      if (values != null && values.Count > 0)
      {
        foreach (var row in values)
        {
          Debug.WriteLine($@"{row[0]} | {row[1]} | {row[2]}");
        }
        return values;
      }
      else
      {
        new TaskDialog("Нет данных для импорта");
        return new List<IList<object>>();
      }
    }
  }
}
