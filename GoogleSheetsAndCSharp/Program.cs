using System; 
using System.IO; 
using Google.Apis.Auth.OAuth2; 
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data; 
using System.Collections.Generic;   //Need this to use lists<>.


namespace GoogleSheetsAndCSharp
{
    class Program
    {
        static readonly string[] Scopes =  { SheetsService.Scope.Spreadsheets }; 
        static readonly string ApplicationName = "Legislators"; 
        static readonly string SpreadsheetID = "1rH7YVFlcDtJ53Bc2F_ZSrKcX1mimPR61pknw7Gvtzf4"; 
        static readonly string sheet = "congress"; 
        static SheetsService service;

        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!"); 

            GoogleCredential credential; 
            using (var stream = new FileStream("client_secrets.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream).CreateScoped(Scopes); 
            }
            
            service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            }); 

            //ReadEntries();
            //CreateEntry();
            //UpdateEntry();
            DeleteEntry(); 
        }

        static void ReadEntries()
        {
            var range = $"{sheet}!A1:F10"; 
            var request = service.Spreadsheets.Values.Get(SpreadsheetID, range); 

            var response = request.Execute(); 
            var values = response.Values; 
            if (values != null && values.Count > 0)
            {
                foreach (var row in values)
                {
                    Console.WriteLine("{0} {1} | {2} | {3}", row[5], row[4], row[3], row[1]); 
                }
            }
            else
            {
                Console.WriteLine("No data found."); 
            }
        }

        static void CreateEntry()
        {
            var range = $"{sheet}!A:F";
            var valueRange = new ValueRange(;

            var objectList = new List<object>() { "Hello!", "This", "was", "inserted", "via", "C#" };
            valueRange.Values = new List<IList<object>> { objectList };

            var appendRequest = service.Spreadsheets.Values.Append(valueRange, SpreadsheetID, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var appendResponse = appendRequest.Execute();    
        }

        static void UpdateEntry()
        {
            var range = $"{sheet}!D543";
            var valueRange = new ValueRange();

            var objectList = new List<object>() { "updated" };
            valueRange.Values = new List<IList<object>> { objectList };
            
            var updateRequest = service.Spreadsheets.Values.Update(valueRange, SpreadsheetID, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;   //UpdateRequest in this line, is different from the variable updateRequest
            var updateResponse = updateRequest.Execute();
        }

        static void DeleteEntry()
        {
            var range = $"{sheet}!A543:F";
            var requestBody = new ClearValuesRequest();

            var deleteRequest = service.Spreadsheets.Values.Clear(requestBody, SpreadsheetID, range);
            var deleteResponse = deleteRequest.Execute();
        }
    }
}
