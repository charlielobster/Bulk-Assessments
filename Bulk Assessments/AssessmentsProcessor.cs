
using ClosedXML.Excel;
using Google.GenAI;
using Google.GenAI.Types;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json.Linq;
using System.Text;
using System.Text.RegularExpressions;
using File = System.IO.File;
using Schema = Google.GenAI.Types.Schema;
using Type = Google.GenAI.Types.Type;

namespace BulkAssessments
{
    public class AssessmentsProcessor
    {
        public static async Task Main(string[] args)
        {
            //  Configuration:
            //  Values found in appsettings.json file
            const string GEMINI_MODEL = "gemini-1.5-flash";

            string apiKey = "API KEY";
            string rubricsPath = "RUBRICS FOLDER PATH";
            string reportsParentPath = "REPORTS PARENT FOLDER PATH";
            string workbookTemplateFullPath = "WORKBOOK TEMPLATE PATH";
            string scoresParentPath = "SCORES PARENT FOLDER PATH";
            string assessPrompt = "ASSESSMENT PROMPT";

            string? rubricFileUri;
            string? reportFileUri;

            //  Rubrics Folder
            //      Rubric is formatted using the convention: "Lab x Rubrics.pdf"
            //      (e.g., "Lab 5 Rubrics.pdf")
            //  For each lab,
            //      Student Reports Folder
            //  Student reports use the convention:
            //      "(2 or 3 letters upper case) Lab x.pdf"
            //      (e.g., "AF Lab 6.pdf", "JPL Lab 3.pdf")
            IConfiguration config = new ConfigurationBuilder()
                .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
                .AddJsonFile("appsettings.json")
                .Build();

            if (config != null)
            {
                apiKey = config.GetValue<string>("Gemini:ApiKey", apiKey);
                rubricsPath = config.GetValue<string>("Paths:Rubrics", rubricsPath);
                reportsParentPath = config.GetValue<string>("Paths:ReportsParent", reportsParentPath);
                scoresParentPath = config.GetValue<string>("Paths:ScoresParent", scoresParentPath);
                workbookTemplateFullPath = config.GetValue<string>("Paths:WorkbookTemplate", workbookTemplateFullPath);               
                assessPrompt = config.GetValue<string>("Prompts:Assess", assessPrompt);
            }

            var client = new Client(apiKey: apiKey);

            // Check available models:
            //var models = await client.Models.ListAsync();
            //await foreach (var m in models)
            //{
            //    Console.WriteLine(m.Name); // Look for strings starting with "models/gemini-3"
            //}

            // Algorithm:
            // For each of the Lab Rubrics in the Rubrics directory
            var labRubrics = Directory.GetFiles(rubricsPath);
            foreach (var labRubricFile in labRubrics)
            {
                var labPrefix = Path.GetFileName(labRubricFile);
                labPrefix = labPrefix.Substring(0, 5);

                string geminiRubricName = ToGeminiName(labRubricFile);
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
                            new () { Text = assessPrompt }
                        ]
                    },

                    // Determinism & Formatting
                    Temperature = 0.0f,
                    ResponseMimeType = "application/json",

                    // Forces the JSON to have these specific keys
                    ResponseSchema = new Schema
                    {
                        Type = Type.Object,
                        Properties = new Dictionary<string, Schema>
                        {
                            {
                                "audit_results", new Schema
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
                            }
                        },
                        Required = [ "audit_results" ]
                    }
                };

                // get the Reports Folder for the Lab
                var labReports = Directory.GetFiles(reportsParentPath + "\\" + labPrefix);

                //  For each student report for that Lab's Reports Folder,
                foreach (var labReport in labReports)
                {
                    using var workbook = new XLWorkbook(workbookTemplateFullPath);
                    var scoresFileName = Path.GetFileName(labReport);
                    var geminiReportName = ToGeminiName(labReport);

                    try
                    {
                        var foundReportFile = await client.Files.GetAsync(geminiReportName);
                        rreportFileUri = foundReportFile.Uri;
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
                    for (int i = 1; i < 4; i++)
                    {
                        var reportFileBytes = await File.ReadAllBytesAsync(labReport);
                        var assessmentResponse = await client.Models.GenerateContentAsync(
                            model: GEMINI_MODEL,
                            contents: [
                                new ()
                                {
                                    Parts = [
                                        new () { Text = "REFERENCE DOCUMENT (RUBRICS):"},
                                        new () { FileData = new FileData { FileUri = rubricFileUri, MimeType = "application/pdf" } },
                                        new () { Text = "TARGET DOCUMENT (STUDENT REPORT):"},
                                        new () { FileData = new FileData { FileUri = reportFileUri, MimeType = "application/pdf" } },
                                        //   reinforce primary objectives and formatting conventions
                                        new () { Text = "Compare the TARGET against the REFERENCE. " + assessPrompt}                                
                                    ]
                                }
                            ]
                        );

                        // Collect each in a separate spreadsheet in a workbook named after the student's report
                        Candidate? candidate = assessmentResponse.Candidates?.FirstOrDefault();
                        if (candidate != null && candidate.Content != null && candidate.Content.Parts != null &&
                            candidate.Content.Parts.Count > 0)
                        {
                            Part part = candidate.Content.Parts[0];
                            string jsonResponse = part.Text ?? "";

                            // Add the worksheet
                            var worksheet = workbook.Worksheets.Add("Run " + i);
                            ConvertToWorksheet(jsonResponse, worksheet);
                        }

                        // Sleep for three minutes to keep TPM down
                        Thread.Sleep(180000);
                    }

                    // Delete the report from gemini's cloud
                    await client.Files.DeleteAsync(geminiReportName);

                    // Copy the workbook into the Lab scores folder
                    workbook.SaveAs(scoresParentPath + "\\" + labPrefix + "\\" +
                        scoresFileName.Substring(0, scoresFileName.IndexOf(".")) + " Scores.xlsx");

                    // Sleep for 5 minutes between students
                    Thread.Sleep(300000);
                }

                // No more use for Rubric cloud file, so delete it.
                await client.Files.DeleteAsync(geminiRubricName);

                // Also take a breather between Rubrics
                Thread.Sleep(300000);
            }
        }

        public static void ConvertToWorksheet(string jsonResponse, IXLWorksheet worksheet)
        {
            jsonResponse = jsonResponse.Substring(jsonResponse.IndexOf("["));
            jsonResponse = jsonResponse.Substring(0, jsonResponse.LastIndexOf("]") + 1);

            var csvBuilder = new StringBuilder();

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
        public static string ToGeminiName(string path)
        {
            // Get name without extension
            string rawName = Path.GetFileNameWithoutExtension(path);

            // Keep only alphanumeric and hyphens, convert to lowercase
            string sanitized = Regex.Replace(rawName.ToLower(), @"[^a-z0-9]", "-")
                                    .Trim('-'); // Clean up trailing hyphens

            // Must start with "files/"  
            return $"files/{sanitized}";
        }

    }
}
