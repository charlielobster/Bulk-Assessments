using ClosedXML.Excel;
using Microsoft.Data.SqlClient;

const string scoresParent = @"..\Scores\";
const string legacyParent = @"..\Labs\";
const string latexParent = @"..\LaTeX\";
const string reportsParent = @"..\Reports\";

const string connectionString = @"Server=(localdb)\MSSQLLocalDB;Initial Catalog=MyDatabase;Integrated Security=True;";

SqlConnection conn = new (connectionString);
SqlCommand cmd;
int rowsAffected = 0;

try
{
    Console.WriteLine("Opening connection: " + connectionString);
    conn.Open();

    Console.WriteLine("For each file in the legacy folder");
    string[] legacyFiles = Directory.GetFiles(legacyParent);
    foreach (var legacyFilePath in legacyFiles) 
    {
        string legacyFileName = Path.GetFileName(legacyFilePath);
        string labName = legacyFileName.Substring(6);
        labName = labName.Substring(0, labName.IndexOf(".pdf"));

        cmd = new("insert into Lab (LabName) values (@LabName); select scope_identity();", conn);
        cmd.Parameters.Add(new SqlParameter("@LabName", labName));
        int labPK = Convert.ToInt32(cmd.ExecuteScalar());
        Console.WriteLine("Created LabPK: " + labPK + " LabName: " + labName);

        cmd = new("insert into LegacyFile (FKLabPK, LegacyFileName, LegacyFileBytes) values (@FKLabPK, @LegacyFileName, @LegacyFileBytes)", conn);
        cmd.Parameters.Add(new SqlParameter("@FKLabPK", labPK));
        cmd.Parameters.Add(new SqlParameter("@LegacyFileName", legacyFileName));
        cmd.Parameters.Add(new SqlParameter("@LegacyFileBytes", File.ReadAllBytes(legacyFilePath)));
        cmd.ExecuteNonQuery();
        Console.WriteLine("Created LegacyFile: " + legacyFileName);

    }

    Console.WriteLine("For each file in the LaTeX folder");
    string[] latexFiles = Directory.GetFiles(latexParent);
    foreach (var latexFilePath in latexFiles)
    {
        Console.WriteLine("Processing File: " + latexFilePath);

        string contents = File.ReadAllText(latexFilePath);
        string duplicateContents = contents;
        string fileName = Path.GetFileName(latexFilePath);

        int labPK = Convert.ToInt32(fileName.Substring(4, 1));
        Console.WriteLine("labPK: " + labPK);

        Console.WriteLine("Creating LaTeXFile");

        string query = "insert into LaTeXFile (LaTeXFileName, LaTeX) values (@LaTeXFileName, @LaTeX)";
        cmd = new SqlCommand(query, conn);
        cmd.Parameters.Add(new SqlParameter("@LaTeXFileName", fileName));
        cmd.Parameters.Add(new SqlParameter("@LaTeX", duplicateContents));

        rowsAffected = cmd.ExecuteNonQuery();
        Console.WriteLine("Added " + rowsAffected + " rows.");

        Console.WriteLine("fileName: " + fileName);

        while (contents.IndexOf("\\section*{") != -1)
        {
            contents = contents.Substring(contents.IndexOf("\\section*{") + 10);
            int nextSection = contents.IndexOf("\\section*{");
            if (nextSection == -1) 
            {
                nextSection = contents.IndexOf("\\end{document}");
            }
            string sectionContents = contents.Substring(0, nextSection);

            contents = contents.Substring(nextSection);

            string rubricGroup = sectionContents.Substring(0, sectionContents.IndexOf(" "));

            Console.WriteLine("Found Rubric Group: " + rubricGroup);
            cmd = new SqlCommand("select PK from RubricGroup where RubricGroupName = @RubricGroupName", conn);
            cmd.Parameters.Add(new SqlParameter("@RubricGroupName", rubricGroup));
            int rubricGroupPK = 0;
            try
            {
                object obj = cmd.ExecuteScalar();
                rubricGroupPK = (int)(obj ?? 0);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            if (rubricGroupPK == 0) { continue; }

            while (sectionContents.IndexOf("\\subsection*{") != -1)
            {
                string rubricLaTeX = sectionContents.Substring(sectionContents.IndexOf("\\subsection*{"));
                rubricLaTeX = rubricLaTeX.Substring(0, rubricLaTeX.IndexOf("\\end{itemize}") + 13);

                Console.WriteLine("Rubric LaTeX: ");
                Console.WriteLine(rubricLaTeX);

                string rubricName = rubricLaTeX.Substring(rubricLaTeX.IndexOf("\\subsection*{") + 13);
                rubricName = rubricName.Substring(0, rubricName.IndexOf(" Rubric"));

                Console.WriteLine("Found Rubric: " + rubricName);

                string rubricRules = sectionContents.Substring(0, sectionContents.IndexOf("\\end{itemize}"));
                rubricRules = rubricRules.Substring(rubricRules.IndexOf("\\item"));

                string prefix = rubricRules.Substring(rubricRules.IndexOf("[") + 1, 2);
                Console.WriteLine("Rubric Prefix: " + prefix);

                Console.WriteLine("Rubric Rules: ");
                Console.WriteLine(rubricRules);

                Console.WriteLine("Checking for Rubric existence: " + prefix + " Name: " + rubricName);

                cmd = new SqlCommand("select PK from Rubric where RubricName = @RubricName", conn);
                cmd.Parameters.Add(new SqlParameter("@RubricName", rubricName));
                int rubricPK = (int)(cmd.ExecuteScalar() ?? 0);
                if (rubricPK == 0)
                {
                    Console.WriteLine("Rubric not found, creating...");

                    cmd = new SqlCommand("insert into Rubric (FKRubricGroupPK, RubricName, Prefix, RubricLaTeX) values (@FKRubricGroupPK, @RubricName, @Prefix, @RubricLaTeX); select scope_identity();", conn);
                    cmd.Parameters.Add(new SqlParameter("@FKRubricGroupPK", rubricGroupPK));
                    cmd.Parameters.Add(new SqlParameter("@RubricName", rubricName.Trim()));
                    cmd.Parameters.Add(new SqlParameter("@Prefix", prefix));
                    cmd.Parameters.Add(new SqlParameter("@RubricLaTeX", rubricLaTeX));
                    rubricPK = Convert.ToInt32(cmd.ExecuteScalar());
                }

                Console.WriteLine("Creating LabRubric");
                cmd = new("insert into LabRubric (FKLabPK, FKRubricPK) values (@FKLabPK, @FKRubricPK)", conn);
                cmd.Parameters.Add(new SqlParameter("@FKLabPK", labPK));
                cmd.Parameters.Add(new SqlParameter("@FKRubricPK", rubricPK));
                cmd.ExecuteNonQuery();

                while (rubricRules.IndexOf("\\item \\textbf{") != -1)
                {
                    string rule = rubricRules.Substring(rubricRules.IndexOf("\\item \\textbf{ ") + 15);
                    rule = rule.Substring(0, rule.IndexOf("\\item") == -1 ? rule.Length : rule.IndexOf("\\item"));
                    string ruleID = rule.Substring(1, 5);
                    string ruleName = rule.Substring(8);
                    ruleName = ruleName.Substring(0, ruleName.IndexOf(" }") > -1 ? ruleName.IndexOf(" }") : 0);

                    string ruleLaTeX = string.Empty;
                    if (rule.Length > rule.IndexOf(" }") + 5)
                    {
                        ruleLaTeX = rule.Substring(rule.IndexOf(" }") + 5);
                    }

                    Console.WriteLine("Checking for Rubric Rule existence: " + ruleID);

                    cmd = new SqlCommand("select PK from RubricRule where RuleID = @RuleID", conn);
                    cmd.Parameters.Add(new SqlParameter("@RuleID", ruleID));
                    int rubricRulePK = (int)(cmd.ExecuteScalar() ?? 0);
                    if (rubricRulePK == 0)
                    {
                        Console.WriteLine("Rule not found, creating...");

                        cmd = new SqlCommand("insert into RubricRule (FKRubricPK, RuleID, RuleName, RuleLaTeX) values (@FKRubricPK, @RuleID, @RuleName, @RuleLaTeX); select scope_identity();", conn);
                        cmd.Parameters.Add(new SqlParameter("@FKRubricPK", rubricPK));
                        cmd.Parameters.Add(new SqlParameter("@RuleID", ruleID));
                        cmd.Parameters.Add(new SqlParameter("@RuleName", ruleName));
                        cmd.Parameters.Add(new SqlParameter("@RuleLaTeX", ruleLaTeX));
                        rubricRulePK = Convert.ToInt32(cmd.ExecuteScalar());
                    }

                    rubricRules = rubricRules.Substring(rubricRules.IndexOf("\\item \\textbf{ ") + 14);
                }

                sectionContents = sectionContents.Substring(sectionContents.IndexOf("\\end{itemize}") + 13);
            }
        }
    }

    Console.WriteLine("For each report contained in reportsParent: " + reportsParent);
    var reportsFolders = Directory.GetDirectories(reportsParent);
    foreach (var reportsFolder in reportsFolders)
    {
        Console.WriteLine("reportsFolder: " + reportsFolder);
        int labPK = Convert.ToInt32(reportsFolder.Substring(reportsFolder.IndexOf("\\Lab ") + 5, 1));
        Console.WriteLine("Found labPK: " + labPK);

        var reportFiles = Directory.GetFiles(reportsFolder);
        foreach (var reportFilePath in reportFiles)
        {
            Console.WriteLine("reportFilePath: " + reportFilePath);
            string reportFileName = Path.GetFileName(reportFilePath);
            string anonymousID = reportFileName.Substring(0, 2);

            cmd = new("select PK from Student where AnonymousID = @AnonymousID", conn);
            cmd.Parameters.Add(new SqlParameter("@AnonymousID", anonymousID));
            int studentPK = Convert.ToInt32(cmd.ExecuteScalar());

            Console.WriteLine("Found studentPK: " + studentPK);

            cmd = new("insert StudentReport (FKStudentPK, FKLabPK, ReportFileName, FileBytes) values (@FKStudentPK, @FKLabPK, @ReportFileName, @FileBytes)", conn);
            cmd.Parameters.Add(new SqlParameter("@FKStudentPK", studentPK));
            cmd.Parameters.Add(new SqlParameter("@FKLabPK", labPK));
            cmd.Parameters.Add(new SqlParameter("@ReportFileName", reportFileName));
            cmd.Parameters.Add(new SqlParameter("@FileBytes", File.ReadAllBytes(reportFilePath)));
            cmd.ExecuteNonQuery();
        }
    }

    Console.WriteLine("For each folder containing Lab Scores in scoresParent: " + scoresParent);
    var scoresFolders = Directory.GetDirectories(scoresParent);
    
    foreach (var scoresFolder in scoresFolders)
    {
        Console.WriteLine("Processing Files in subfolder: " + scoresFolder);
        var reportScoresFiles = Directory.GetFiles(scoresFolder);

        Console.WriteLine("scoresFolder: " + scoresFolder);
        int labPK = Convert.ToInt32(scoresFolder.Substring(scoresFolder.IndexOf("\\Lab ") + 5, 1));
        Console.WriteLine("Found labPK: " + labPK);

        foreach (var scoresFilePath in reportScoresFiles)
        {
            Console.WriteLine("Processing File: " + scoresFilePath);
            string scoresFileName = Path.GetFileName(scoresFilePath);
            string anonymousID = scoresFileName.Substring(0, 2);

            cmd = new("select PK from Student where AnonymousID = @AnonymousID", conn);
            cmd.Parameters.Add(new SqlParameter("@AnonymousID", anonymousID));
            int studentPK = Convert.ToInt32(cmd.ExecuteScalar());

            cmd = new("select max(PK) as PK from StudentReport where FKStudentPK = @FKStudentPK and FKLabPK = @FKLabPK", conn);
            cmd.Parameters.Add(new SqlParameter("@FKStudentPK", studentPK));
            cmd.Parameters.Add(new SqlParameter("@FKLabPK", labPK));
            int studentReportPK = Convert.ToInt32(cmd.ExecuteScalar());

            cmd = new("insert ScoresFile (FKStudentReportPK, ScoresFileName, ScoresFileBytes) values (@FKStudentReportPK, @ScoresFileName, @ScoresFileBytes); select scope_identity();", conn);
            cmd.Parameters.Add(new SqlParameter("@FKStudentReportPK", studentReportPK));
            cmd.Parameters.Add(new SqlParameter("@ScoresFileName", scoresFileName));
            cmd.Parameters.Add(new SqlParameter("@ScoresFileBytes", File.ReadAllBytes(scoresFilePath)));
            int scoresFilePK = Convert.ToInt32(cmd.ExecuteScalar());

            XLWorkbook scoresWorkbook = new(scoresFilePath);
            for (int r = 1; r < 4; r++)
            {
                int PECount = 0;
                int EPCount = 0;
                int CPCount = 0;

                IXLWorksheet runSheet = scoresWorkbook.Worksheet("Run " + r);
                cmd = new("insert Assessment (FKScoresFilePK) values (@FKScoresFilePK); select scope_identity();", conn);
                cmd.Parameters.Add(new SqlParameter("@FKScoresFilePK", scoresFilePK));
                int assessmentPK = Convert.ToInt32(cmd.ExecuteScalar());

                int i = 2;
                while (runSheet.Cell(i, 1).Value.ToString() != "")
                {
                    string ruleID = runSheet.Cell(i, 1).Value.ToString();
                    if (ruleID == "RuleID")
                    {
                        i++;
                        continue;
                    }
                    if (ruleID.StartsWith("["))
                    {
                        ruleID = ruleID.Substring(1);
                    }
                    if (ruleID.EndsWith("]"))
                    {
                        ruleID = ruleID.Substring(0, ruleID.IndexOf("]"));
                    }
                    string rulePrefix = ruleID.Substring(0, 2);
                    int score = Convert.ToInt32(runSheet.Cell(i, 2).Value.ToString());
                    string ruleName = runSheet.Cell(i, 3).Value.ToString().Trim();

                    if (rulePrefix == "PE") // && ruleName == "Pendulum Experiment") 
                    {
                        if (PECount < 3)
                        {
                            PECount++;
                        }
                        else
                        {
                            rulePrefix = "PF";
                        }
                    }
                    else if (rulePrefix == "EP")
                    {
                        if (EPCount < 4)
                        {
                            EPCount++;
                        }
                        else
                        {
                            rulePrefix = "EQ";
                        }
                    }
                    else if (rulePrefix == "SA" && labPK == 5)
                    {
                        rulePrefix = "SB";
                    }
                    else if (rulePrefix == "CP")
                    {
                        if (CPCount < 3)
                        {
                            CPCount++;
                        }
                        else
                        {
                            rulePrefix = "CQ";
                        }
                    }
                    else if (rulePrefix == "RG")
                    {
                        rulePrefix = "RQ";
                    }

                    string evidence = runSheet.Cell(i, 4).Value.ToString();

                    cmd = new("select PK from RubricRule where RuleID = @RuleID", conn);
                    cmd.Parameters.Add(new SqlParameter("@RuleID", rulePrefix + ruleID.Substring(2)));
                    int rulePK = Convert.ToInt32(cmd.ExecuteScalar());

                    if (rulePK == 0)
                    {
                        Console.WriteLine("Rubric not found: " + rulePrefix + ruleID.Substring(2));
                    }

                    Console.WriteLine("Found RubricRule, PK: " + rulePK);

                    cmd = new("insert RuleScore(FKRulePK, FKAssessmentPK, Evidence, Score) values (@FKRulePK, @FKAssessmentPK, @Evidence, @Score); select scope_identity();", conn);
                    cmd.Parameters.Add(new SqlParameter("@FKRulePK", rulePK));
                    cmd.Parameters.Add(new SqlParameter("@FKAssessmentPK", assessmentPK));
                    cmd.Parameters.Add(new SqlParameter("@Evidence", evidence));
                    cmd.Parameters.Add(new SqlParameter("@Score", score));

                    int ruleScorePK = Convert.ToInt32(cmd.ExecuteScalar());

                    Console.WriteLine("Create RuleScore PK: " + ruleScorePK);
                    i++;
                }
            }
        }
    }    
}
catch (Exception ex)
{
    Console.WriteLine("Caught exception: " + ex.Message);
}


