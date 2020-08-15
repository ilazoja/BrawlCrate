using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace BrawlCrate.Spreadsheet
{
    // TODO: Use EPPlus 5 instead of Interop Excel?
    public class SheetBuilder
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "ProjectPlusBuildReporter";

        SheetsService service;
        String spreadsheetId = "14OZlDwHTwwyc3Gsq78hCKE8ruJkPOJeV_207c6bJ4s4";

        int numRequests = 0;

        Application oXL;
        _Workbook oWB;
        _Worksheet oSheet;
        Range oRng;

        int currentLineXl = 2;

        public SheetBuilder()
        {
            UserCredential credential;

            using (var stream =
                new FileStream("Spreadsheet\\credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Sheets API service.
            service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });


            // Clear column
            String range = "List of Songs!F2:F99999"; // update cell F5
            ClearValuesRequest body = new ClearValuesRequest();
            SpreadsheetsResource.ValuesResource.ClearRequest clear = service.Spreadsheets.Values.Clear(body, spreadsheetId, range);
            ClearValuesResponse result = clear.Execute();

            //Start Excel and get Application object.
            oXL = new Application();
            oXL.Visible = true;

            //Get a new workbook.
            oWB = (_Workbook)(oXL.Workbooks.Add(Missing.Value));
            oSheet = (_Worksheet)oWB.ActiveSheet;

            //Add table headers going cell by cell.
            oSheet.Cells[1, 1] = "Song Name";
            oSheet.Cells[1, 2] = "Game Origin";
            oSheet.Cells[1, 3] = "Song ID";
            oSheet.Cells[1, 4] = "Song Filename";
            oSheet.Cells[1, 5] = "Remix Name";
            oSheet.Cells[1, 6] = "Remixer";

            //Format A1:F1 as bold, vertical alignment = center.
            oSheet.get_Range("A1", "F1").Font.Bold = true;
            oSheet.get_Range("A1", "F1").VerticalAlignment = XlVAlign.xlVAlignCenter;
            oSheet.get_Range("A1", "F1").HorizontalAlignment = XlHAlign.xlHAlignCenter;

        }

        public void addTlstToSongList(String tlstName, String fileName, String songID, bool isPinch)
        {

            // Step 1: Write Equation to Find Row

            String range = "List of Songs!A2"; // update cell F5
            ValueRange valueRange = new ValueRange();
            valueRange.MajorDimension = "COLUMNS"; //"ROWS" ;//COLUMNS

            var oblist = new List<object>() { "=ROW(INDIRECT(ADDRESS(MATCH(\"" + fileName + "\",C:C,0),1))) " };
            valueRange.Values = new List<IList<object>> { oblist };

            SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, range);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;

            UpdateValuesResponse result = update.Execute();

            // Step 2: Get Row Location

            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, range);

            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            if (values != null && values.Count > 0)
            {
                if (int.TryParse(values[0][0].ToString(), out int rowLocation))
                {
                    // Step 3: Get current value
                    range = "List of Songs!A" + rowLocation.ToString() + ":F" + rowLocation.ToString();
                    request = service.Spreadsheets.Values.Get(spreadsheetId, range);
                    response = request.Execute();
                    values = response.Values;

                    // Step 4: Update value
                    String updatedValue = String.Empty;
                    if (values != null && values[0].Count > 5)
                    {
                        updatedValue = values[0][5].ToString() + ", " + tlstName;
                    }
                    else updatedValue = tlstName;

                    range = "List of Songs!F" + rowLocation.ToString();
                    valueRange = new ValueRange();
                    valueRange.MajorDimension = "COLUMNS"; //"ROWS" ;//COLUMNS
                    oblist = new List<object>() { updatedValue };
                    valueRange.Values = new List<IList<object>> { oblist };

                    update = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, range);
                    update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                    result = update.Execute();

                    switch(values[0].Count)
                    {
                        case 1:
                            addToTracklist(values[0][0].ToString(), "", songID, "", "", "", isPinch);
                            break;
                        case 2:
                            addToTracklist(values[0][0].ToString(), values[0][1].ToString(), songID, "", "", "", isPinch);
                            break;
                        case 3:
                            addToTracklist(values[0][0].ToString(), values[0][1].ToString(), songID, values[0][2].ToString(), "", "", isPinch);
                            break;
                        case 4:
                            addToTracklist(values[0][0].ToString(), values[0][1].ToString(), songID, values[0][2].ToString(), values[0][3].ToString(), "", isPinch);
                            break;
                        case 5:
                        case 6:
                            addToTracklist(values[0][0].ToString(), values[0][1].ToString(), songID, values[0][2].ToString(), values[0][3].ToString(), values[0][4].ToString(), isPinch);
                            break;
                        default:
                            addToTracklist("", "", songID, "", "", "", isPinch);
                            break;
                    }
                }
            }

            numRequests += 2;

            if (numRequests == 98) // Can't exceed more than 100 requests per 100 seconds per user, avoid error 429 (Too Many Requests) can maybe increase quota
            {
                Thread.Sleep(100000);
                numRequests = 0;

            }
        }

        public void addTracklistHeader(String tlstName)
        {
            oSheet.Cells[currentLineXl, 1] = tlstName;
            oSheet.Range[oSheet.Cells[currentLineXl, 1], oSheet.Cells[currentLineXl, 6]].Merge();
            oSheet.Range[oSheet.Cells[currentLineXl, 1], oSheet.Cells[currentLineXl, 6]].Interior.Color = System.Drawing.Color.LightGray;
            oSheet.Range[oSheet.Cells[currentLineXl, 1], oSheet.Cells[currentLineXl, 6]].Font.Bold = true;
            currentLineXl++;
            
        }

        private void addToTracklist(String songName, String gameOrigin, String songID, String songFileName, String remixName, String remixer, bool isPinch)
        {
            oSheet.Cells[currentLineXl, 1] = songName;
            oSheet.Cells[currentLineXl, 2] = gameOrigin;
            oSheet.Cells[currentLineXl, 3] = songID;
            oSheet.Cells[currentLineXl, 4] = songFileName;
            oSheet.Cells[currentLineXl, 5] = remixName;
            oSheet.Cells[currentLineXl, 6] = remixer;
            if (isPinch) oSheet.Range[oSheet.Cells[currentLineXl, 1], oSheet.Cells[currentLineXl, 1]].HorizontalAlignment = XlHAlign.xlHAlignRight;
            currentLineXl++;
        }

        // TODO: More formatting
    }
}
