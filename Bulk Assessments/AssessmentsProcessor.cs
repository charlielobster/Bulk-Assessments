
using ClosedXML.Excel;
using CsvHelper;
using Google.GenAI;
using Google.GenAI.Types;
using Microsoft.Extensions.Configuration;
using System.Globalization;
using File = System.IO.File;

namespace BulkAssessments
{
    public class AssessmentsProcessor
    {
        public static async Task Main(string[] args)
        {
            //  Configuration:
            //  Values found in appsettings.json file
            const string GEMINI_MODEL = "gemini-2.5-flash";
            string apiKey = "API KEY";
            string rubricsPath = "RUBRICS FOLDER PATH";
            string reportsParentPath = "REPORTS PARENT FOLDER PATH";
            string workbookTemplateFullPath = "WORKBOOK TEMPLATE PATH";
            string scoresParentPath = "SCORES PARENT FOLDER PATH";
            string assessPrompt = "ASSESSMENT PROMPT";

            //  Rubrics Folder
            //      Rubric is formatted using the convention: "Lab x Rubrics.pdf"
            //      (e.g., "Lab 5 Rubrics.pdf")
            //  For each lab,
            //      Verification Prompt
            //          
            //      Student Reports Folder
            //  Student reports use the convention:
            //      "(2 or 3 letters upper case) Lab x.pdf"
            //      (e.g., "AF Lab 6.pdf", "JPL Lab 3.pdf")
            IConfiguration config = new ConfigurationBuilder()
                .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
                .AddJsonFile("appsettings.json")
                .Build();

            var labPrompts = new Dictionary<string, string>();

            if (config != null)
            {
                apiKey = config.GetValue<string>("Gemini:ApiKey", apiKey);
                rubricsPath = config.GetValue<string>("Paths:Rubrics", rubricsPath);
                reportsParentPath = config.GetValue<string>("Paths:ReportsParent", reportsParentPath);
                scoresParentPath = config.GetValue<string>("Paths:ScoresParent", scoresParentPath);
                workbookTemplateFullPath = config.GetValue<string>("Paths:WorkbookTemplate", workbookTemplateFullPath);

                for (int i = 1; i < 8; i++)
                {
                    string key = "Lab " + i;
                    string value = "";
                    value = config.GetValue<string>("Prompts:" + key, value);
                    labPrompts.Add(key, value);
                }
                assessPrompt = config.GetValue<string>("Prompts:Assess", assessPrompt);
            }

            var client = new Client(apiKey: apiKey);

            // Algorithm:
            // For each of the Lab Rubrics in the Rubrics directory
            var labRubrics = Directory.GetFiles(rubricsPath);
            foreach (var labRubricFile in labRubrics)
            {
                var labPrefix = Path.GetFileName(labRubricFile);
                labPrefix = labPrefix.Substring(0, 5);
                var labPrompt = labPrompts[labPrefix];

                var rubricFileBytes = await File.ReadAllBytesAsync(labRubricFile);

                // upload the rubric file if there's no uri for it (is this necessary?)
                // here we upload the rubrics file to gemini
                /*
                var uploadedFile = await client.Files.UploadAsync(
                    labRubricFile,
                    new UploadFileConfig { MimeType = "application/pdf" }
                );
                */
                var verifyResponse = await client.Models.GenerateContentAsync(
                    model: GEMINI_MODEL,
                    contents: [
                        new ()
                        {
                            Parts = [
                                new () { InlineData = new Blob () { MimeType = "application/pdf", Data = rubricFileBytes } }, 
//                                new () { FileData = new FileData { FileUri = uploadedFile.Uri, MimeType = uploadedFile.MimeType } },
                                new () { Text = labPrompt }
                            ]
                        }
                    ]
                );

                // todo: check the verifyResponse for matching grand totals

                // get the Reports Folder for the Lab
                var labReports = Directory.GetFiles(reportsParentPath + "\\" + labPrefix);

                //  For each student report for that Lab's Reports Folder,
                foreach (var labReport in labReports)
                {
                    using var workbook = new XLWorkbook(workbookTemplateFullPath);
                    var scoresFileName = Path.GetFileName(labReport);

                    //  run the report assessment three times.
                    for (int i = 1; i < 3; i++)
                    {
                        var reportFileBytes = await File.ReadAllBytesAsync(labReport);
                        var assessmentResponse = await client.Models.GenerateContentAsync(
                            model: GEMINI_MODEL,
                            contents: [
                                new ()
                                {
                                    Role = "user",
                                    Parts = [
                                        new () { InlineData = new Blob () { MimeType = "application/pdf", Data = reportFileBytes } },
                                        new () { Text = assessPrompt}
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
                            string? csvScores = part.Text;

                            // debug
                            Console.WriteLine(csvScores);

                            if (csvScores != null)
                            {
                                using var reader = new StringReader(csvScores);
                                using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);

                                // 1. Get the data as a list of dynamic objects or strings
                                var scores = csv.GetRecords<dynamic>().ToList();

                                // 2. Add your tab
                                var ws = workbook.Worksheets.Add("Run " + i);

                                // 3. This method takes the collection and maps it to cells automatically
                                ws.Cell(1, 1).InsertTable(scores);
                            }
                        }
                    }

                    // Copy the workbook into the Lab scores folder
                    workbook.SaveAs(scoresParentPath + "\\" + labPrefix + "\\" +
                        scoresFileName.Substring(0, scoresFileName.IndexOf(".")) + " Scores.xlsx");
                }
            }
        }
    }
}
