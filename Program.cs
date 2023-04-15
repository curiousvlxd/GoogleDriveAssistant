using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Util.Store;
using Google.Apis.Sheets.v4.Data;

namespace GoogleDriveAssistant
{
    class Program
    {
        private static readonly string[] Scopes = { DriveService.Scope.Drive, SheetsService.Scope.Spreadsheets };
        private static readonly string ApplicationName = "Google Drive Assistant";
        private static readonly string SpreadsheetName = "AllFilesSpreadsheet";
        private static readonly string CredentialsFilePath = "credentials.json";
        private static readonly string UserTokenFilePath = "user-token";

        static async Task Main(string[] args)
        {
            UserCredential credential = await GetUserCredentialAsync();

            using (var driveService = new DriveService(new BaseClientService.Initializer()
                   {
                       HttpClientInitializer = credential,
                       ApplicationName = ApplicationName,
                   }))
            using (var sheetsService = new SheetsService(new BaseClientService.Initializer()
                   {
                       HttpClientInitializer = credential,
                       ApplicationName = ApplicationName,
                   }))
            {
                await UpdateSpreadsheetEvery15Minutes(sheetsService, driveService);
            }
        }

        private static async Task<UserCredential> GetUserCredentialAsync()
        {
            using (var stream = new FileStream(CredentialsFilePath, FileMode.Open, FileAccess.Read))
            {
                string credPath = UserTokenFilePath;
                return await GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)
                );
            }
        }

        private static async Task<IList<Google.Apis.Drive.v3.Data.File>> GetAllDriveFilesAsync(
            DriveService driveService)
        {
            FilesResource.ListRequest listRequest = driveService.Files.List();
            listRequest.PageSize = 10;
            listRequest.Fields = "nextPageToken, files(name, createdTime)";
            listRequest.Q = "trashed = false";
            var files = new List<Google.Apis.Drive.v3.Data.File>();
            string nextPageToken = null;

            do
            {
                listRequest.PageToken = nextPageToken;
                var response = await listRequest.ExecuteAsync();
                files.AddRange(response.Files);
                nextPageToken = response.NextPageToken;
            } while (nextPageToken != null);

            return files;
        }

        private static async Task<Google.Apis.Drive.v3.Data.File> GetOrCreateSpreadsheetAsync(
            SheetsService sheetsService, DriveService driveService)
        {
            FilesResource.ListRequest spreadsheetRequest = driveService.Files.List();
            spreadsheetRequest.Q = $"name = '{SpreadsheetName}' and trashed = false";

            var spreadsheetResponse = await spreadsheetRequest.ExecuteAsync();

            if (spreadsheetResponse.Files.Count == 0)
            {
                var createRequest = new Google.Apis.Sheets.v4.Data.Spreadsheet
                {
                    Properties = new SpreadsheetProperties
                    {
                        Title = SpreadsheetName
                    }
                };
                var createResponse = await sheetsService.Spreadsheets.Create(createRequest).ExecuteAsync();
                var spreadsheet = await driveService.Files.Get(createResponse.SpreadsheetId).ExecuteAsync();
                Console.WriteLine($"Spreadsheet created: {spreadsheet.Name}");
                return spreadsheet;
            }
            else
            {
                var spreadsheet = spreadsheetResponse.Files[0];
                Console.WriteLine($"Spreadsheet found: {spreadsheet.Name}");
                return spreadsheet;
            }
        }

        private static async Task UpdateSpreadsheetWithDataAsync(
            SheetsService sheetsService, Google.Apis.Drive.v3.Data.File spreadsheet,
            IList<Google.Apis.Drive.v3.Data.File> files)
        {
            var values = new List<IList<object>>();
            foreach (var file in files)
            {
                values.Add(new List<object> { file.Name, file.CreatedTime.Value.ToString("yyyy-MM-dd") });
            }

            var body = new ValueRange
            {
                Values = values
            };

            var updateRequest = sheetsService.Spreadsheets.Values.Update(body, spreadsheet.Id, "A1");
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            var updateResponse = await updateRequest.ExecuteAsync();
            Console.WriteLine($"Spreadsheet updated: {updateResponse.UpdatedCells} cells updated.");

            var batchUpdateRequest = new BatchUpdateSpreadsheetRequest();
            var autoResizeDimensionsRequest = new AutoResizeDimensionsRequest
            {
                Dimensions = new DimensionRange
                {
                    SheetId = 0,
                    Dimension = "COLUMNS",
                    StartIndex = 0,
                    EndIndex = 1
                }
            };
            batchUpdateRequest.Requests = new List<Request> {
                new Request {
                    AutoResizeDimensions = autoResizeDimensionsRequest
                }
            };
            await sheetsService.Spreadsheets.BatchUpdate(batchUpdateRequest, spreadsheet.Id).ExecuteAsync();
        }
        
        private static async Task UpdateSpreadsheetEvery15Minutes(
            SheetsService sheetsService, DriveService driveService)
        {
            while (true)
            {
                IList<Google.Apis.Drive.v3.Data.File> files = await GetAllDriveFilesAsync(driveService);

                Google.Apis.Drive.v3.Data.File spreadsheet =
                    await GetOrCreateSpreadsheetAsync(sheetsService, driveService);

                await UpdateSpreadsheetWithDataAsync(sheetsService, spreadsheet, files);

                Console.WriteLine($"Spreadsheet updated at {DateTime.Now.ToString("hh:mm:ss tt")}");

                await Task.Delay(TimeSpan.FromMinutes(15));
            }
        }
    }
}
