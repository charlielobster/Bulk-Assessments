
using ClosedXML.Excel;
using Google.GenAI;
using Google.GenAI.Types;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;
using File = System.IO.File;
using Schema = Google.GenAI.Types.Schema;
using Type = Google.GenAI.Types.Type;
using Path = System.IO.Path;

namespace BulkAssessments
{
    public class GeminiAlias 
    {
        public GeminiAlias() { Name = "NAME"; ApiKey = "API KEY"; }
        public GeminiAlias(string name, string apiKey) { Name = name; ApiKey = apiKey; } 
        public string Name { get; set; } // prefix for user file uploads (to avoid naming conflicts)
        public string ApiKey { get; set; } 
    }

    public class AssessmentsProcessor
    {
        public static async Task Main(string[] args)
        {
            //  Configuration:
            //  Values found in appsettings.json file
            string geminiModel = "GEMINI MODEL"; // Check available models
            string rubricsPath = "RUBRICS FOLDER PATH";
            string reportsParentPath = "REPORTS PARENT FOLDER PATH";
            string processedParentPath = "PROCESSED PARENT FOLDER PATH";
            string workbookTemplateFullPath = "WORKBOOK TEMPLATE PATH";
            string scoresParentPath = "SCORES PARENT FOLDER PATH";
            string assessmentPrompt = "ASSESSMENT PROMPT";

            string? rubricFileUri;
            string? reportFileUri;

            List<GeminiAlias> aliases = [];
            var aliasEnumerator = aliases.GetEnumerator();

            //  Rubrics Folder
            //      Rubric is formatted using the convention: "Lab x Rubrics.pdf"
            //      (e.g., "Lab 5 Rubrics.pdf")
            //  For each lab,
            //      Student Reports Folder
            //  Student reports use the convention:
            //      "(2 letters upper case) Lab x.pdf"
            //      (e.g., "AF Lab 6.pdf", "JP Lab 3.pdf")
            IConfiguration config = new ConfigurationBuilder()
                .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
                .AddJsonFile("appsettings.json")
                .Build();

            if (config != null)
            {
                geminiModel = config.GetValue<string>("Gemini:Model", geminiModel);
                rubricsPath = config.GetValue<string>("Paths:Rubrics", rubricsPath);
                reportsParentPath = config.GetValue<string>("Paths:ReportsParent", reportsParentPath);
                processedParentPath = config.GetValue<string>("Paths:ProcessedParent", processedParentPath);
                config.GetSection("Gemini:Aliases").Bind(aliases);
                aliasEnumerator = aliases.GetEnumerator();
                aliasEnumerator.MoveNext();
                scoresParentPath = config.GetValue<string>("Paths:ScoresParent", scoresParentPath);
                workbookTemplateFullPath = config.GetValue<string>("Paths:WorkbookTemplate", workbookTemplateFullPath);               
                assessmentPrompt = config.GetValue<string>("Prompts:Assessment", assessmentPrompt);
            }

        start:

            var client = new Client(apiKey: aliasEnumerator.Current.ApiKey);

            // Iterate through and display file details
            /*var filesResponse = await client.Files.ListAsync();
            if (filesResponse != null)
            {
                await foreach (var file in filesResponse)
                {
                    Console.WriteLine($"File Name: {file.DisplayName}");
                    Console.WriteLine($"File URI: {file.Uri}");
                    Console.WriteLine($"Mime Type: {file.MimeType}");
                    Console.WriteLine("----------------------------");
                }
            }*/

            // Check available models:
            /*var models = await client.Models.ListAsync();
            await foreach (var m in models)
            {
                System.Console.WriteLine(m.Name); // look for strings starting with "models/gemini-3"
            }*/
            /*
                models/gemini-2.5-flash
                models/gemini-2.5-pro
                models/gemini-2.0-flash
                models/gemini-2.0-flash-001
                models/gemini-2.0-flash-lite-001
                models/gemini-2.0-flash-lite
                models/gemini-2.5-flash-preview-tts
                models/gemini-2.5-pro-preview-tts
                models/gemma-3-1b-it
                models/gemma-3-4b-it
                models/gemma-3-12b-it
                models/gemma-3-27b-it
                models/gemma-3n-e4b-it
                models/gemma-3n-e2b-it
                models/gemma-4-26b-a4b-it
                models/gemma-4-31b-it
                models/gemini-flash-latest
                models/gemini-flash-lite-latest
                models/gemini-pro-latest
                models/gemini-2.5-flash-lite
                models/gemini-2.5-flash-image
                models/gemini-3-pro-preview
                models/gemini-3-flash-preview
                models/gemini-3.1-pro-preview
                models/gemini-3.1-pro-preview-customtools
                models/gemini-3.1-flash-lite-preview
                models/gemini-3-pro-image-preview
                models/nano-banana-pro-preview
                models/gemini-3.1-flash-image-preview
                models/lyria-3-clip-preview
                models/lyria-3-pro-preview
                models/gemini-3.1-flash-tts-preview
                models/gemini-robotics-er-1.5-preview
                models/gemini-robotics-er-1.6-preview
                models/gemini-2.5-computer-use-preview-10-2025
                models/deep-research-max-preview-04-2026
                models/deep-research-preview-04-2026
                models/deep-research-pro-preview-12-2025
                models/gemini-embedding-001
                models/gemini-embedding-2-preview
                models/gemini-embedding-2
                models/aqa
                models/imagen-4.0-generate-001
                models/imagen-4.0-ultra-generate-001
                models/imagen-4.0-fast-generate-001
                models/veo-2.0-generate-001
                models/veo-3.0-generate-001
                models/veo-3.0-fast-generate-001
                models/veo-3.1-generate-preview
                models/veo-3.1-fast-generate-preview
                models/veo-3.1-lite-generate-preview
                models/gemini-2.5-flash-native-audio-latest
                models/gemini-2.5-flash-native-audio-preview-09-2025
                models/gemini-2.5-flash-native-audio-preview-12-2025
                models/gemini-3.1-flash-live-preview
            */

            // Algorithm:
            // For each of the Lab Rubrics in the Rubrics directory
            var labRubrics = Directory.GetFiles(rubricsPath);
            foreach (var labRubricFile in labRubrics)
            {
                var labPrefix = Path.GetFileName(labRubricFile);
                labPrefix = labPrefix.Substring(0, 5);

                string geminiRubricName = ToGeminiName(labRubricFile, aliasEnumerator.Current.Name);
                try
                {
                    var foundRubricFile = await client.Files.GetAsync(geminiRubricName);
                    rubricFileUri = foundRubricFile.Uri;
                }
                catch (ClientError e) when (e.Status == "PERMISSION_DENIED")
                {
                    // create a new version of the file in gemini's cloud
                    var uploadedRubricFile = await client.Files.UploadAsync(
                        labRubricFile,
                        new UploadFileConfig { Name = geminiRubricName, MimeType = "application/pdf" }
                    );
                    rubricFileUri = uploadedRubricFile.Uri;
                }

                // Setup the GenerateContentConfig
                var contentConfig = new GenerateContentConfig
                {
                    // System Instructions include detailed assessment AI Rules
                    SystemInstruction = new Content
                    {
                        Parts = [
                            new () { Text = assessmentPrompt }
                        ]
                    },

                    // Determinism & Formatting
                    Temperature = 0.0f,
                    ResponseMimeType = "application/json",

                    // Forces the JSON to have these specific keys
                    ResponseSchema = new Schema
                    {
                        Type = Type.Array,
                        Items = new Schema
                        {
                            Type = Type.Object,
                            Properties = new Dictionary<string, Schema>
                            {
                                { "RuleID", new Schema { Type = Type.String } },
                                { "RuleName", new Schema { Type = Type.String } },
                                { "Evidence", new Schema { Type = Type.String } },
                                { "Score", new Schema { Type = Type.Integer } }
                            },
                            Required = ["RuleID", "RuleName", "Evidence", "Score" ]
                       } 
                    }
                };

                // get the Reports Folder for the Lab
                var labReports = Directory.GetFiles(reportsParentPath + "\\" + labPrefix);

                //  For each student report for that Lab's Reports Folder,
                foreach (var labReport in labReports)
                {
                    XLWorkbook workbook;

                    var geminiReportName = ToGeminiName(labReport, aliasEnumerator.Current.Name);

                    string scoresFileName = Path.GetFileName(labReport);
                    string reportScoresFullPath = scoresParentPath + "\\" + labPrefix + "\\" +
                        scoresFileName.Substring(0, scoresFileName.IndexOf(".")) + " Scores.xlsx";

                    // Create the Report Scores Workbook now, if it doesn't already exist
                    if (File.Exists(reportScoresFullPath))
                    {
                        workbook = new XLWorkbook(reportScoresFullPath);
                    }
                    else
                    {
                        workbook = new XLWorkbook(workbookTemplateFullPath);
                        workbook.SaveAs(reportScoresFullPath);
                    }

                    // If the report is not already uploaded, upload it now
                    try
                    {
                        var foundReportFile = await client.Files.GetAsync(geminiReportName);
                        reportFileUri = foundReportFile.Uri;
                    }
                    catch (ClientError e) when (e.Status == "PERMISSION_DENIED")
                    {
                        // or else create a new version of the report in gemini's cloud.
                        var uploadedReportFile = await client.Files.UploadAsync(
                            labReport,
                            new UploadFileConfig { Name = geminiReportName, MimeType = "application/pdf" }
                        );
                        reportFileUri = uploadedReportFile.Uri;
                    }

                    //  Run the report assessment three times.
                    //  todo: config settings for run count?
                    for (int i = 1; i < 4; i++)
                    {
                        // If Worksheet already exists, we already processed this run.
                        if (workbook.Worksheets.FirstOrDefault(ws => ws.Name == "Run " + i) != default) continue;

                        try
                        {
                            // Attempt an assessment
                            var assessmentResponse = await client.Models.GenerateContentAsync(
                                model: geminiModel,
                                contents: [
                                    new ()
                            {
                                Parts = [
                                    new () { Text = "REFERENCE DOCUMENT (RUBRICS):"},
                                    new () { FileData = new FileData { FileUri = rubricFileUri, MimeType = "application/pdf" } },
                                    new () { Text = "TARGET DOCUMENT (STUDENT REPORT):"},
                                    new () { FileData = new FileData { FileUri = reportFileUri, MimeType = "application/pdf" } },
                                    //   reinforce primary objectives and formatting conventions
                                    new () { Text = "Compare the TARGET against the REFERENCE. " + assessmentPrompt}
                                ]
                            }
                                ]
                            );

                            // Collect each assessment in a separate spreadsheet in this Report's Scores Workbook
                            Candidate? candidate = assessmentResponse.Candidates?.FirstOrDefault();
                            if (candidate != null && candidate.Content != null && candidate.Content.Parts != null &&
                                candidate.Content.Parts.Count > 0)
                            {
                                Part part = candidate.Content.Parts[0];
                                string jsonResponse = part.Text ?? "";

                                // Add the worksheet.
                                var worksheet = workbook.Worksheets.Add("Run " + i);
                                ConvertToWorksheet(jsonResponse, worksheet);

                                // Save the changes immediately.
                                workbook.Save();
                            }

                        }
                        catch (ClientError ex) when (ex.StatusCode == 429)
                        {
                            // Specific handling for 429 / Resource Exhausted
                            Console.WriteLine("Resources exhausted for alias: " + aliasEnumerator.Current.Name);

                            // Switch to the next alias/api key
                            if (!aliasEnumerator.MoveNext())
                            {
                                aliasEnumerator = aliases.GetEnumerator();
                                aliasEnumerator.MoveNext();
                            }
                            // Take a break between aliases
                            Thread.Sleep(60000);

                            // take it from the top
                            goto start;
                        }

                        // Sleep for a minute to keep TPM down
                        Thread.Sleep(60000);
                    }

                    // Delete the report from gemini's cloud
                    await client.Files.DeleteAsync(geminiReportName);

                    // Move the Report to the Processed Folder
                    if (!Directory.Exists(processedParentPath + "\\Reports\\" + labPrefix)) {
                        Directory.CreateDirectory(processedParentPath + "\\Reports\\" + labPrefix);
                    }
                    File.Move(labReport, processedParentPath + "\\Reports\\" + labPrefix + "\\" + Path.GetFileName(labReport));

                    // Sleep for a minute between Reports.
                    Thread.Sleep(60000);
                }

                // No more use for Rubric cloud file, so delete it.
                // todo: Move Rubrics file to Processed folder
                await client.Files.DeleteAsync(geminiRubricName);

                // Move the Rubrics to the Processed folder.
                File.Move(labRubricFile, processedParentPath + "\\Rubrics\\" + Path.GetFileName(labRubricFile));

                // Also take a breather between Rubrics
                Thread.Sleep(300000);
            }
        }

        public static void ConvertToWorksheet(string jsonResponse, IXLWorksheet worksheet)
        {
            jsonResponse = jsonResponse.Substring(jsonResponse.IndexOf("["));
            jsonResponse = jsonResponse.Substring(0, jsonResponse.LastIndexOf("]") + 1);

            var results = JArray.Parse(jsonResponse);

            worksheet.Cell(1, 1).Value = "Rule ID";
            worksheet.Cell(1, 2).Value = "Score";
            worksheet.Cell(1, 3).Value = "Rule Name";
            worksheet.Cell(1, 4).Value = "Evidence";

            var currentRow = 2;
            foreach (var item in results)
            {
                int score = 0;
                JToken? scoreObject = item["Score"];
                if (scoreObject != null)
                {
                    switch (scoreObject.Type)
                    {
                        case JTokenType.Integer:
                            score = scoreObject?.ToObject<int>() ?? 0;
                            break;
                        case JTokenType.String:
                            if ((scoreObject?.ToString() ?? "").ToLower().Equals("pass")) 
                            {
                                score = 1;
                            }
                            break;
                        default:
                            break;
                    }
                }

                worksheet.Cell(currentRow, 1).Value = item["RuleID"]?.ToString() ?? "";
                worksheet.Cell(currentRow, 2).Value = score;
                worksheet.Cell(currentRow, 3).Value = item["RuleName"]?.ToString() ?? "";
                worksheet.Cell(currentRow, 4).Value = item["Evidence"]?.ToString() ?? "";

                currentRow++;
            }
        }

        // todo: simplify this for conventions above
        public static string ToGeminiName(string path, string alias = "")
        {
            // Get name without extension
            string rawName = Path.GetFileNameWithoutExtension(path);

            // Keep only alphanumeric and hyphens, convert to lowercase
            string sanitized = Regex.Replace(rawName.ToLower(), @"[^a-z0-9]", "-")
                                    .Trim('-'); // Clean up trailing hyphens

            // Must start with "files/"  
            return $"files/{alias}{sanitized}";
        }

    }
}
