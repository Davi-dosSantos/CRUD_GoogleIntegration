using System;
using System.IO;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Services;
using Google.Apis.Sheets.v4.Data;
using System.Collections.Generic;
using System.Security.Cryptography;
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
            //string googleClientId = "360690838705-ndsllj69cv7jbbk8u1cnlfm2g34o39it.apps.googleusercontent.com";
            //string googleClientSecret = "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";
            //string[] scopes = new[] { SheetsService.Scope.Spreadsheets };

            //UserCredential credential = LoginAsync(googleClientId, googleClientSecret, scopes).Result;

            string googleServiceAccountEmail = "ecos12pfteste@datasheet1-392420.iam.gserviceaccount.com";
            string googleServiceAccountPrivateKey = "-----BEGIN PRIVATE KEY-----\nMIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQCJEsFTwvncQzZh\np5dddmdA40mHjRShzrWNSKEOUaEHZLdmuCZoXknTz492C9QFvb/2CLmiCQYOUs7v\nOQ6NMtY5OFyTav6jV40OToPx7X/3QvwJCDRkiqC04WG6ZiOAnc8EZ5IUw71E6AIn\nsXuHSEFUKQcBTCP2M5qahPNjJkO7x3JA4sFcWJjzQtbNm04R83zoM9rqW3GXxSou\nvDLmBOMzJNGeUP2BRk46Mg5GAfwc7SIdEnz4vX+lETTUDaGEAn98e7gDntiRAdFI\nJTEzU8oLw0oPPPtrk6alacgTWIfIfIiF50S8C8Galj6IBYZMkcAr87RgR6nBQ7zr\n/TpJY3UlAgMBAAECggEANRKk9iiVE9qWUMNSEScKHY6jZq+SYIAnvXd0nJWwkqtF\nc6kzfc+cKD3CX0N/KWXp0HpaXcm+pYcchnWCE9uuJGOVPKL9ywLYI8T0w5Rgqr0t\n1tVta8xdIwvtCf4IGwF/KUZswktzmh120CWhHaU1Xj+wbaksd2RNpSx7DFXBfg/P\nKIJHlRSL8mGDiRD7x4B1pdrEBadOgwfk0PfGVbd2laTIbknIs15v4JWKj3+Zhv8K\nFwCE83r5Wr1N7p+2Ou2xr9FoE2jwS556w+3bdhcDZaOia/HBV4DzKmAyVVdiiUGq\nqMtrq4bOmitOHy71eJLpue/bphYH0iI6HCSdamyKtQKBgQDAgnt/BB6Z/6YP7R9J\nVz2bY1fUgGRjjz/IrHaWmwjW3YPNCvDgHDbCydp9EObKd1siyTnqE48QqVnInnS8\nnZMhFMJ24NaHviv5B0Wwfmm8i3sYIKw6jJDVIbNCBn5pkWrS/x3CkWODJwPr5Pbs\n2I7cJvIC4l9K78gUZvBS/FHeOwKBgQC2R8vps7L/x5SvxJqt86iaTMW693SQ+IGW\nzwfYOYh/5iz8wZl9PE+qcOxIaD573PONzzRAAbkVcv+PW+hS3sngy8rPTqhwQRYq\njXrqLyyP+zt4/ZsZ+1Oz3uab8oDVpA4Wjf/uoWVikbKAsmatT/h/EFH5Gg2jjP2W\nejpgqtfkHwKBgQCpzy9CRg78Rm5kJATx+5tjQskJsEtdKtHXoJFmndC5P2JwbpM1\nDI4dWlJ4+Xyq4Yepcpi8ao5K4ydIeMV+TvymNJqopAF4cX52Rzzox0lbwClPihqB\n9tYWuohV2EaPtm7lOZY1t2txF+w0m55YI1o4xb26X5YxEruJi5e3i8xnWwKBgCjK\nhF12M1Z+CU4URzEqV86/43flrJZMpmNjTTQcG+nTTrn5cSnPd1yDDL1fZqw9U9um\nROEWAZ9FLt+cB6+T38WIlYgy6ArG5fj71EfX6rcF19dJmY4E6kRUW3MGn8Ivhl+R\nw3ZZc+DNDg8y3TtnrApzUoTWSbsR8CXekHXVhZ6tAoGBAI0sFm6aBNUJJvKknBaa\nxNc5E46oH1ym1fz4WBsuu1mtMK4A5NFKmHEhdMS4QlI6Gp1AEV3Qy1w1T8VH3334\nKmpGoufdXU42dwk3TuQsH3xjAScUv5jcD61+UjaFHJWvMWmF/dCDDOiQacejxvyJ\nH97KLbEW1CLreVXKReJKyG8l\n-----END PRIVATE KEY-----\n";
            string[] scopes = new[] { SheetsService.Scope.Spreadsheets };

            ServiceAccountCredential credential = AuthenticateWithServiceAccount(googleServiceAccountEmail, googleServiceAccountPrivateKey, scopes);

            service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName
            });

            DeleteEntry();
            CreateEntry();
            ReadEntries();
            UpdateEntry();

        }

        public static ServiceAccountCredential AuthenticateWithServiceAccount(string serviceAccountEmail, string serviceAccountPrivateKey, string[] scopes)
        {
            var credential = new ServiceAccountCredential(new ServiceAccountCredential.Initializer(serviceAccountEmail)
            {
                Scopes = scopes
            }.FromPrivateKey(serviceAccountPrivateKey));

            return credential;
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
