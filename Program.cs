using System;
using System.IO;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Services;
using Google.Apis.Sheets.v4.Data;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace GoogleSheetsAndCsharp
{
    class Program
    {
        static readonly string ApplicationName = "DataSheet1";
        static readonly string SpreadsheetId = "1qrJ1e1o8PbUNZLFXLw6TuTUOm5shgwMecJi0Q5PsUxs";
        static readonly string sheet = "FirstPage";
        static SheetsService? service;

        static void Main(string[] args)
        {
            string googleAPIKey ="AIzaSyA79XVffL0n5m2pvJ1_0Zx8jZ3HNRyZXkI"
            string googleClientId = "360690838705-ndsllj69cv7jbbk8u1cnlfm2g34o39it.apps.googleusercontent.com";
            string googleClientSecret = "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";
            string[] scopes = new[] { SheetsService.Scope.Spreadsheets };

            UserCredential credential = LoginAsync(googleClientId, googleClientSecret, scopes).Result;

            service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName
            });

            string line;

            for(;;){
                    line = "";
                    Console.WriteLine("Select Comand C R U D , E to exit");
                    line = Console.ReadLine();
                    if (line == "C" ){
                        CreateEntry();  
                    }else if (line == "R"){
                        ReadEntries();  
                    }else if (line == "U"){
                        UpdateEntry();
                    }else if (line == "D"){
                        DeleteEntry();  
                    }else if (line == "E"){
                        break;  
                    }else  Console.WriteLine("Comando invalido");
           }

         
        }

        public static async Task<UserCredential> LoginAsync(string googleClientId, string googleClientSecret, string[] scopes)
        {
            ClientSecrets secrets = new ClientSecrets()
            {
                ClientId = googleClientId,
                ClientSecret = googleClientSecret
            };

            return await GoogleWebAuthorizationBroker.AuthorizeAsync(secrets, scopes, "user", System.Threading.CancellationToken.None);
        }

        static void ReadEntries()
        {
            if (service == null)
            {
                Console.WriteLine("Google Sheets API service is not initialized.");
                return;
            }

            var range = $"{sheet}!A:C";
            var request = service.Spreadsheets.Values.Get(SpreadsheetId, range);

            try
            {
                var response = request.Execute();
                var values = response.Values;
                if (values != null && values.Count > 0)
                {
                    foreach (var row in values)
                    {
                        // Print columns A to F, which correspond to indices 0 and 4.
                        Console.WriteLine("{0} | {1} | {2}", row[0], row[1], row[2]);
                    }
                }
                else
                {
                    Console.WriteLine("No data found.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }

         static void CreateEntry()
        {
            var range = $"{sheet}!A:F";
            var valueRange = new ValueRange();

            var oblist = new List<object>() { "Hello!", "This", "was", "inserted", "via", "C#" };
            valueRange.Values = new List<IList<object>> { oblist };

            var appendRequest = service.Spreadsheets.Values.Append(valueRange, SpreadsheetId, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var appendReponse = appendRequest.Execute();
        }

        static void UpdateEntry()
        {
            var range = $"{sheet}!D543";
            var valueRange = new ValueRange();

            var oblist = new List<object>() { "updated" };
            valueRange.Values = new List<IList<object>> { oblist };

            var updateRequest = service.Spreadsheets.Values.Update(valueRange, SpreadsheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var appendReponse = updateRequest.Execute();
        }

        static void DeleteEntry()
        {
            var range = $"{sheet}!A34:F";
            var requestBody = new ClearValuesRequest();

            var deleteRequest = service.Spreadsheets.Values.Clear(requestBody, SpreadsheetId, range);
            var deleteReponse = deleteRequest.Execute();
        }
    }
}
