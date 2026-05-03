
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

// MAJOR TODO: integrate with Brightspace API
// https://community.d2l.com/brightspace/kb/articles/1102-getting-started-with-brightspace-api-automation

namespace BulkAssessments
{
    // Alias DTO
    public class GeminiAlias 
    {
        // Two constructors for such a tiny class seems a little overkill...
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
            string rubricsPath = "RUBRICS FOLDER PATH";
            string reportsParentPath = "REPORTS PARENT FOLDER PATH";
            string processedParentPath = "PROCESSED PARENT FOLDER PATH";
            string workbookTemplateFullPath = "WORKBOOK TEMPLATE PATH";
            string scoresParentPath = "SCORES PARENT FOLDER PATH";
            string assessmentPrompt = "ASSESSMENT PROMPT";
            int sleepInterval = 60000;

            string? rubricFileUri;
            string? reportFileUri;

            // Handle Quota Limits:
            // Try changing account alias and api keys first
            List<GeminiAlias> aliases = [];
            var aliasEnumerator = aliases.GetEnumerator();

            bool anyApiKeyWorks = false; // if true, any api key for the current enumeration of GeminiAliases worked.
                                         // i.e., we completed a request without hitting a quota exception   

            // If those run out, round-robin through the Gemini Models
            // (i.e., Prefer using the model as long as possible before switching.)
            List<string> models = [];
            var modelsEnumerator = models.GetEnumerator();

            //  Rubrics Folder:
            //      Rubric is formatted using the convention: "Lab x Rubrics.pdf"
            //      (e.g., "Lab 5 Rubrics.pdf")
            //  For each lab,
            //      Student Reports Folder:
            //          <ReportsParent>\Lab x\
            //          (e.g., "..\Data\Reports\Lab 3\")
            //  Reports use the convention:
            //      "(2 letters upper case) Lab x.pdf"
            //      (e.g., "AF Lab 6.pdf", "JP Lab 3.pdf")
            IConfiguration config = new ConfigurationBuilder()
                .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
                .AddJsonFile("appsettings.json")
                .Build();

            if (config != null)
            {
                config.GetSection("Gemini:Models").Bind(models);
                modelsEnumerator = models.GetEnumerator();
                modelsEnumerator.MoveNext();

                config.GetSection("Gemini:Aliases").Bind(aliases);
                aliasEnumerator = aliases.GetEnumerator();
                aliasEnumerator.MoveNext();

                rubricsPath = config.GetValue<string>("Paths:Rubrics", rubricsPath);
                reportsParentPath = config.GetValue<string>("Paths:ReportsParent", reportsParentPath);
                processedParentPath = config.GetValue<string>("Paths:ProcessedParent", processedParentPath);
                scoresParentPath = config.GetValue<string>("Paths:ScoresParent", scoresParentPath);
                workbookTemplateFullPath = config.GetValue<string>("Paths:WorkbookTemplate", workbookTemplateFullPath);               

                assessmentPrompt = config.GetValue<string>("Prompts:Assessment", assessmentPrompt);
                sleepInterval = config.GetValue<int>("SleepInterval", sleepInterval);
            }

            #region ConfigurationDump
            Console.WriteLine();
            Console.WriteLine("--- CONFIGURATION ---");
            Console.WriteLine();
            Console.WriteLine("Gemini Aliases");
            Console.WriteLine("Name\t\tApiKey");
            foreach (var alias in aliases)
            {
                Console.WriteLine(alias.Name + "\t\t" + alias.ApiKey);
            }
            Console.WriteLine("Current Alias Name: " + aliasEnumerator.Current.Name + " ApiKey: " + aliasEnumerator.Current.ApiKey);
            Console.WriteLine();
            Console.WriteLine("Gemini AI Models");
            foreach (var model in models)
            {
                Console.WriteLine(model);
            }
            Console.WriteLine("Current Model: " + modelsEnumerator.Current);
            Console.WriteLine();
            Console.WriteLine("Paths");
            Console.WriteLine("Rubrics Folder: " + rubricsPath);
            Console.WriteLine("Reports Parent: " + reportsParentPath);
            Console.WriteLine("Processed Files Parent: " + processedParentPath);
            Console.WriteLine("Scores Parent: " + scoresParentPath);
            Console.WriteLine();
            Console.WriteLine("Workbook Template Full Path: " + workbookTemplateFullPath);
            Console.WriteLine("Sleep Interval: " + sleepInterval);
            Console.WriteLine();
            Console.WriteLine("Assessment Prompt: ");
            Console.WriteLine(assessmentPrompt);
            Console.WriteLine();
            Console.WriteLine();
            #endregion

        restart:

            Console.WriteLine();
            Console.WriteLine("--- MAIN LOOP ---");
            Console.WriteLine();

            var client = new Client(apiKey: aliasEnumerator.Current.ApiKey);

            /* // Iterate through and display file details
            var filesResponse = await client.Files.ListAsync();
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
            /*  // Models example output

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
            Console.WriteLine("Found " + labRubrics.Length + " Rubrics files");

            #region RubricsLoop

            foreach (var labRubricsFile in labRubrics)
            {
                Console.WriteLine();
                Console.WriteLine("--- RUBRICS PROCESSING ---");
                Console.WriteLine();
                Console.WriteLine("Checking Rubric: " + labRubricsFile);

                var labPrefix = Path.GetFileName(labRubricsFile);
                labPrefix = labPrefix.Substring(0, 5);

                string geminiRubricsName = ToGeminiName(labRubricsFile, aliasEnumerator.Current.Name);
                try
                {
                    // If the Rubrics file is not already uploaded, upload it now.
                    Console.WriteLine("Checking for Rubrics file existence on cloud: " + geminiRubricsName);

                    var foundRubricFile = await client.Files.GetAsync(geminiRubricsName);
                    rubricFileUri = foundRubricFile.Uri;
                }
                catch (ClientError e) when (e.Status == "PERMISSION_DENIED")
                {
                    // Otherwise, create a new version of the Rubrics file in gemini's cloud.
                    Console.WriteLine("File not found, uploading file: " + geminiRubricsName);

                    var uploadedRubricFile = await client.Files.UploadAsync(
                        labRubricsFile,
                        new UploadFileConfig { Name = geminiRubricsName, MimeType = "application/pdf" }
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

                // Get the Reports Folder for the Lab.
                var labReports = Directory.GetFiles(reportsParentPath + labPrefix);

                Console.WriteLine("Found " + labReports.Length + " Lab Reports to assess.");

                #region ReportsLoop

                //  For each student report for that Lab's Reports Folder,
                foreach (var labReport in labReports)
                {
                    Console.WriteLine();
                    Console.WriteLine("--- REPORT PROCESSING ---");
                    Console.WriteLine();
                    Console.WriteLine("Assessing Report " + labReport);

                    XLWorkbook workbook;

                    var geminiReportName = ToGeminiName(labReport, aliasEnumerator.Current.Name);

                    string scoresFileName = Path.GetFileName(labReport);
                    string reportScoresFullPath = scoresParentPath + labPrefix + "\\" +
                        scoresFileName.Substring(0, scoresFileName.IndexOf(".")) + " Scores.xlsx";

                    // Create the Report Scores Workbook now, if it doesn't already exist.
                    if (File.Exists(reportScoresFullPath))
                    {
                        Console.WriteLine("Found existing Scores Workbook: " + reportScoresFullPath);
                        workbook = new XLWorkbook(reportScoresFullPath);
                    }
                    else
                    {
                        Console.WriteLine("Creating new Scores Workbook: " + reportScoresFullPath);
                        workbook = new XLWorkbook(workbookTemplateFullPath);
                        workbook.SaveAs(reportScoresFullPath);
                    }

                    // If the Report is not already uploaded, upload it now.
                    try
                    {
                        Console.WriteLine("Checking for Report file existence on cloud: " + geminiReportName);
                        var foundReportFile = await client.Files.GetAsync(geminiReportName);
                        reportFileUri = foundReportFile.Uri;
                    }
                    catch (ClientError e) when (e.Status == "PERMISSION_DENIED")
                    {
                        // Otherwise, create a new version of the Report in gemini's cloud.
                        Console.WriteLine("File not found, uploading file: " + geminiReportName);

                        var uploadedReportFile = await client.Files.UploadAsync(
                            labReport,
                            new UploadFileConfig { Name = geminiReportName, MimeType = "application/pdf" }
                        );
                        reportFileUri = uploadedReportFile.Uri;
                    }

                    #region AssessmentsLoop

                    //  Run the report assessment three times.
                    //  todo: config settings for run count?
                    for (int i = 1; i < 4; i++)
                    {
                        Console.WriteLine();
                        Console.WriteLine("--- REPORT ASSESSMENT ---");
                        Console.WriteLine();

                        // If Worksheet already exists, we already processed this run.
                        if (workbook.Worksheets.FirstOrDefault(ws => ws.Name == "Run " + i) != default)
                        {
                            Console.WriteLine("Found existing sheet for this index: " + i + " -- Continuing...");
                            continue;
                        }
                        try
                        {
                            // Attempt an assessment
                            Console.WriteLine("Requesting Assessment, index: " + i + " ...");

                            var assessmentResponse = await client.Models.GenerateContentAsync(
                                model: modelsEnumerator.Current,
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

                            Console.WriteLine("Assessment Complete");

                            // Collect each assessment in a separate spreadsheet in this Report's Scores Workbook
                            Candidate? candidate = assessmentResponse.Candidates?.FirstOrDefault();
                            if (candidate != null && candidate.Content != null && candidate.Content.Parts != null &&
                                candidate.Content.Parts.Count > 0)
                            {
                                Part part = candidate.Content.Parts[0];
                                string jsonResponse = part.Text ?? "";

                                try
                                {
                                    // Can get various exceptions here, checking manually.
                                    // (could be PII, bad JSON formatting, etc)
                                    if (jsonResponse.IndexOf("[") > -1 && jsonResponse.LastIndexOf("]") > -1)
                                    {
                                        jsonResponse = jsonResponse.Substring(jsonResponse.IndexOf("["));
                                        jsonResponse = jsonResponse.Substring(0, jsonResponse.LastIndexOf("]") + 1);

                                        var results = JArray.Parse(jsonResponse);

                                        // Add the worksheet.
                                        var worksheet = workbook.Worksheets.Add("Run " + i);
                                        ConvertToWorksheet(results, worksheet);

                                        // Save the changes immediately.
                                        workbook.Save();

                                        Console.WriteLine("Successfully added Worksheet: " + "Run " + i);
                                        Console.WriteLine("Saved Workbook: " + reportScoresFullPath);

                                        // We did something! Flag it.
                                        if (!anyApiKeyWorks)
                                        {
                                            anyApiKeyWorks = true;
                                            Console.WriteLine("AnyApiKeyWorks new value: " + anyApiKeyWorks);
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine("Json format issue, no []s found: " + jsonResponse);

                                        // take it from the top
                                        goto restart;
                                    }
                                }
                                catch (Exception ex) // set breakpoint here
                                {
                                    Console.WriteLine("Exception: " + ex.Message);

                                    // take it from the top
                                    goto restart;
                                }
                            }
                        }
                        // Catch quota ClientErrors
                        catch (ClientError ex) when (ex.StatusCode == 429)
                        {
                            // Specific handling for 429 / Resource Exhausted
                            Console.WriteLine();
                            Console.WriteLine("--- RESOURCE EXHAUSTION --");
                            Console.WriteLine();
                            Console.WriteLine("Resources exhausted for alias: " + aliasEnumerator.Current.Name);

                            // Switch to the next alias and api key
                            if (!aliasEnumerator.MoveNext())
                            {
                                Console.WriteLine();
                                Console.WriteLine("--- ALIAS EXHAUSTION --");
                                Console.WriteLine();

                                // If we reached here, we spent all our quotas, for all our aliases
                                Console.WriteLine("Exhausted all aliases for the current model: " + modelsEnumerator.Current);

                                // Reset the alias queue back to the first item
                                aliasEnumerator = aliases.GetEnumerator();
                                aliasEnumerator.MoveNext();

                                // Check to see if we did anything with any alias last round
                                if (!anyApiKeyWorks)
                                {
                                    Console.WriteLine();
                                    Console.WriteLine("--- MODEL EXHAUSTION --");
                                    Console.WriteLine();

                                    // If all our aliases are exhausted and we didn't do anything last round, change the model.
                                    if (!modelsEnumerator.MoveNext())
                                    {
                                        Console.WriteLine("Exhausted all models, resetting back to first.");

                                        // Round-robin back to the first model in the list
                                        modelsEnumerator = models.GetEnumerator();
                                        modelsEnumerator.MoveNext();
                                    }
                                    Console.WriteLine("Changed to new model: " + modelsEnumerator.Current);
                                }
                                else
                                {
                                    // Swapped in a new alias, so haven't done anything yet.
                                    // Reset the anyApiKeyWorks flag.
                                    anyApiKeyWorks = false;
                                    Console.WriteLine("AnyApiKeyWorks new value: " + anyApiKeyWorks);
                                }
                            }
                            else
                            {
                                Console.WriteLine("Trying with next alias: " + aliasEnumerator.Current.Name);
                            }

                            // Take a break between swapping aliases and/or models
                            Console.WriteLine("Sleeping now: " + sleepInterval);
                            Thread.Sleep(sleepInterval);
                            Console.WriteLine();

                            // Take it from the top.
                            goto restart;
                        }

                        // Sleep for a minute to keep TPM down
                        Console.WriteLine("Sleeping now: " + sleepInterval);
                        Thread.Sleep(sleepInterval);
                        Console.WriteLine();
                    }

                    #endregion // AssessmentsLoop

                    Console.WriteLine();
                    Console.WriteLine("--- REPORT FILE CLEAN-UP ---");
                    Console.WriteLine();

                    // Delete the report from gemini's cloud
                    Console.WriteLine("Deleting Report file from cloud: " + geminiReportName);
                    await client.Files.DeleteAsync(geminiReportName);

                    // Move the Report to the Processed Folder
                    if (!Directory.Exists(processedParentPath + "Reports\\" + labPrefix)) {

                        Console.WriteLine("Creating new folder: " + processedParentPath + "Reports\\" + labPrefix);
                        Directory.CreateDirectory(processedParentPath + "Reports\\" + labPrefix);
                    }

                    Console.WriteLine("Moving to Processed folder - Report file: " + labReport);
                    Console.WriteLine("New Location: " + processedParentPath + "Reports\\" + labPrefix + "\\" + Path.GetFileName(labReport));
                    File.Move(labReport, processedParentPath + "Reports\\" + labPrefix + "\\" + Path.GetFileName(labReport));

                    // Sleep for a minute between Reports.
                    Console.WriteLine("Sleeping now: " + sleepInterval);
                    Thread.Sleep(sleepInterval);
                    Console.WriteLine();
                }

                #endregion // ReportsLoop

                Console.WriteLine();
                Console.WriteLine("--- RUBRICS FILE CLEAN-UP ---");
                Console.WriteLine();

                // No more use for Rubric cloud file, so delete it.
                // todo: Move Rubrics file to Processed folder
                Console.WriteLine("Deleting Rubrics file from cloud: " + geminiRubricsName);
                await client.Files.DeleteAsync(geminiRubricsName);

                // Move the Rubrics to the Processed folder.
                Console.WriteLine("Moving to Processed folder - Report file: " + labRubricsFile);
                Console.WriteLine("New Location: " + processedParentPath + "\\Rubrics\\" + Path.GetFileName(labRubricsFile));
                File.Move(labRubricsFile, processedParentPath + "\\Rubrics\\" + Path.GetFileName(labRubricsFile));

                // Also take a breather between Rubrics
                Console.WriteLine("Sleeping now: " + (3 * sleepInterval));
                Thread.Sleep(3 * sleepInterval);
                Console.WriteLine();
            }

            #endregion // RubricsLoop

        }

        public static void ConvertToWorksheet(JArray assessmentResults, IXLWorksheet worksheet)
        {
            worksheet.Cell(1, 1).Value = "Rule ID";
            worksheet.Cell(1, 2).Value = "Score";
            worksheet.Cell(1, 3).Value = "Rule Name";
            worksheet.Cell(1, 4).Value = "Evidence";

            var currentRow = 2;
            foreach (var rule in assessmentResults)
            {
                int score = 0;
                JToken? scoreObject = rule["Score"];
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

                worksheet.Cell(currentRow, 1).Value = rule["RuleID"]?.ToString() ?? "";
                worksheet.Cell(currentRow, 2).Value = score;
                worksheet.Cell(currentRow, 3).Value = rule["RuleName"]?.ToString() ?? "";
                worksheet.Cell(currentRow, 4).Value = rule["Evidence"]?.ToString() ?? "";

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
            return $"files/{alias}-{sanitized}";
        }

    }
}
