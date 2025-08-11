using System;
using System.Globalization;
using Spectre.Console;
using System.IO;
using System.Text;
using Microsoft.Data.SqlClient;
using ClosedXML.Excel;
using System.Data;
using ExcelDataReader;
using System.Diagnostics;
using System.Collections;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Data.Common;
using System.Xml.Linq;

namespace KrakenExport
{
    class Program
    {
        // db connection string same for each meathod
        public static string connectionString = Environment.GetEnvironmentVariable("KRAKEN_DB_CONNECTION") ?? "";

        private record FormatDefinition(string Name, int StartRow, int SampleCol, int PlateCol, int WellCol);

        private static readonly Dictionary<int, FormatDefinition> FormatMap = new()
        {
            //this is a list of format definitions for the excel form the name, startrow, coloumns etc.
            { 1, new FormatDefinition("Sample ID Starts in column B",1, 1, 2, 3) }, // Intertek
            { 2, new FormatDefinition("Sample ID starts in column A",1, 0, 1, 2) }   // EIB
        };

        public static void Main(string[] args)
        {
            if (args.Length == 0 || args.Contains("--interactive"))
            {
                while (true)
                {
                    AnsiConsole.Write(
                        new Panel("[bold cyan]Kraken CLI Tool[/]")
                        .Expand()
                        .BorderColor(Spectre.Console.Color.Blue));
                    var capitalizeFirst = (string s) => string.IsNullOrWhiteSpace(s) ? s : char.ToUpper(s[0]) + s.Substring(1).ToLower();

                    string username = Environment.UserName;
                    string firstName = username.Contains('.') ? username.Split('.')[0] : username;
                    string capitalized = capitalizeFirst(firstName);

                    var mainChoice = AnsiConsole.Prompt(
                        new SelectionPrompt<string>()
                            .Title("[yellow]Select an action:[/]")
                            .AddChoices(
                                "Generate XML from single seed order in Lims with wells.",
                                "Generate XML from Excel file",
                                "Generate XML from empty plates(only when no wells are defined in lims)",
                                "Help",
                                "Exit"
                            )
                    );
                    switch (mainChoice)
                    {
                        case "Generate XML from single seed order in Lims with wells.":
                            XmlFromDb();
                            break;
                        case "Generate XML from Excel file":
                            XmlFromExcel();
                            break;
                        case "Generate XML from empty plates(only when no wells are defined in lims)":
                            XmlFromDbWithoutWells();
                            break;
                        case "Exit":
                            AnsiConsole.MarkupLine($"[green]See you next time {capitalized}![/]");
                            return;
                        case "Help":
                            AnsiConsole.Write(
                                new Panel(
                                    "[bold yellow]Help & Instructions[/]\n\n" +
                                    "[bold]Generate XML from single seed order in Lims with wells[/]:\n" +
                                    "  - Use this for orders where:\n" +
                                    "    1. Each individual well is registered on pre-sampled plates (well defined by the customer),\n" +
                                    "    2. Or for sampled orders from bags, where you must specify the number of wells sampled per plate to convert sample IDs to well positions.\n\n" +
                                    "[bold]Generate XML from Excel file[/]:\n" +
                                    "  - Use this for SNP line orders where sample information is defined in the order form.\n" +
                                    "  - The Intertek plate ID will be matched with the order form's plate ID.\n" +
                                    "       [bold] Chose type[/]:\n"+
                                    "           - Specify if routine or verification order, verification will create copies of the plate and well with d1 and d2 added to the names.\n"+
                                    "           - Routine \n" +
                                    "                 -Snp: will create ntc in h11 h22 snp+dart will create ntc in g12 h12\n"+
                                    "           - Verification will create copies of plate and well with d1 and d2.\n"+
                                    "       [bold]Excel Formats[/]:\n" +
                                    "           - Specify where in the Sample List sheet the customers well name is located(col a or b).\n"+
                                    "[bold]Generate XML from empty plates (only when no wells are defined in lims)[/]:\n" +
                                    "  - Use this only for orders where samples exist but no wells are defined in the system.\n" +
                                    "  - You will be prompted for the number of wells, which will be generated for each plate in the system.\n"
                                    
                                )
                                .BorderColor(Spectre.Console.Color.Gold1)
                                .Header("[bold blue]KrakenImport Help[/]", Justify.Center)
                                .Expand()
                            );
                            break;
                    }
                }
            }
        }

        static void XmlFromDb()
        {
            string orderId = AnsiConsole.Ask<string>("Enter [green]Order ID[/] (e.g. SE-25-0130):");
            string outpath = PromptForOutputDirectory();

            int wellsPerPlate = 92;

            string query = @"
                SELECT 
                    ID_TEXT,
                    ITK_PLATE_ID,
                    ENTITY_TEMPLATE_ID,
                    SAMPLE_NAME,
                    CUSTOMER_PLATE_WELL,
                    ID_NUMERIC
                FROM sample
                WHERE 
                    ENTITY_TEMPLATE_ID IN ('PCR_SINGLE_SAMPLE', 'PCR_SINGLE_SUB_SAMPLE_88_WELL') 
                    AND JOB_NAME = @orderId
                ORDER BY ID_NUMERIC;
                ";

            try
            {
                using var conn = new SqlConnection(connectionString);
                conn.Open();

                using var cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@orderId", orderId!);

                var sampleList = new List<(string SampleWellName, string CustomerPlateId, string Well, string PlateIdText)>();
                AnsiConsole.Status()
    .Start("Reading from database...", ctx =>
    {
        using var reader = cmd.ExecuteReader();
        while (reader.Read())
        {
            string idText = reader.GetString(0);
            string plateIdText = reader.GetString(1);
            string template = reader.GetString(2);
            string sampleName = reader.GetString(3);
            string well = reader.GetString(4);

            string finalWell = template == "PCR_SINGLE_SUB_SAMPLE_88_WELL"
                ? MapWellFromId(idText, wellsPerPlate)
                : well;
            sampleList.Add((idText, sampleName, finalWell, plateIdText));
        }
    });

                if(sampleList == null || sampleList.Count == 0){

                AnsiConsole.MarkupLine($"[red] Export aborted:[/] No samples found for {orderId}");
                return;
                }
                // XML Export
                string xmlOutputFile = Path.Combine(outpath, $"{orderId}-Kraken_masterplates.xml");
                GenerateKrakenMasterPlatesXml(
                    projectType: orderId,
                    sampleList: sampleList,
                    outputPath: xmlOutputFile
                );
                AnsiConsole.MarkupLine($"[green] Export complete:[/] [underline]{xmlOutputFile}[/]");

            }
            catch (Exception ex)
            {
                AnsiConsole.WriteException(ex,
                    ExceptionFormats.ShortenPaths | ExceptionFormats.ShortenTypes);

            }
        }

        static string MapWellFromId(string idText, int wellsPerPlate)
        {
            if (idText.Length < 3 || !int.TryParse(idText[^3..], out int num))
                return "BAD_ID";

            int pos = (num % wellsPerPlate == 0) ? wellsPerPlate : num % wellsPerPlate;
            int row = (pos - 1) / 12;
            int col = (pos - 1) % 12 + 1;
            char rowChar = (char)('A' + row);
            return $"{rowChar}{col:D2}";
        }
        private static void XmlFromExcel()
        {
            while (true)
            {
                var formatOptions = new Dictionary<string, int>
                {
                    { "subject ID in col B (standard in Intertek forms)", 1 },
                    { "subject ID in col A (typical for EIB)", 2 },
                    { "[red]Back[/]", 3 }
                };

                var selectedType = AnsiConsole.Prompt(
                    new SelectionPrompt<string>()
                        .Title("[yellow]Choose type:[/]")
                        .AddChoices("Routine", "Verification", "[red]Back[/]")
                );

                if (selectedType == "[red]Back[/]")
                    return;
                bool dartorder = false;
                if (selectedType == "Routine")
                {
                    var ordertype = AnsiConsole.Prompt(
                        new SelectionPrompt<string>()
                            .Title("[yellow]Choose routine type:[/]")
                            .AddChoices("Snp", "Snp+Dart", "[red]Back[/]")



                    );

                    if (ordertype == "[red]Back[/]")
                        return;
                    else if (ordertype == "Snp+Dart")
                    {
                        dartorder = true;
                    } 
                    }

                var selected = AnsiConsole.Prompt(
                    new SelectionPrompt<string>()
                        .Title("[yellow]Choose Excel format:[/]")
                        .AddChoices(formatOptions.Keys)
                );
                int formatChoice = formatOptions[selected];

                switch (formatChoice)
                {
                    case 1:
                    case 2:
                        ParseGenericFormat(FormatMap[formatChoice], selectedType, dartorder);
                        return;
                    case 3:
                        return;
                    default:
                        break;
                }
            }

        }

        private static void ParseGenericFormat(FormatDefinition def, string selectedType, bool dartorder = false)
        {

            string orderId = AnsiConsole.Ask<string>("Enter [green]Order ID[/] (e.g. SE-25-0130):");

            string excelPath = PromptForFilePath("Enter path to Excel file [grey](drag & drop works)[/]:");

            string outpath = PromptForOutputDirectory();

            int blankCount = 0, skippedCount = 0, correctedWell = 0;

            try
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                var sampleList = new List<(string SampleWellName, string CustomerPlateId, string Well, string PlateIdText)>();

                using var stream = File.Open(excelPath, FileMode.Open, FileAccess.Read);
                using var reader = ExcelReaderFactory.CreateReader(stream);

                var result = reader.AsDataSet();
                var sheet = result.Tables.Cast<DataTable>()
                    .FirstOrDefault(t => t.TableName.Equals("Sample List", StringComparison.OrdinalIgnoreCase));

                if (sheet == null)
                {
                    AnsiConsole.MarkupLine("[red]Sheet 'Sample List' not found.[/]");
                    return;
                }

                AnsiConsole.Status().Start("[green] Reading Excel file...[/]", ctx =>
                {
                    for (int i = def.StartRow; i < sheet.Rows.Count; i++)
                    {
                        var row = sheet.Rows[i];

                        if (row == null || row.ItemArray.All(cell => string.IsNullOrWhiteSpace(cell?.ToString())))
                            continue;

                        if (row.ItemArray.Length < Math.Max(def.SampleCol, Math.Max(def.PlateCol, def.WellCol)) + 1)
                            continue;

                        string sampleWellName = row[def.SampleCol]?.ToString()?.Trim() ?? "";
                        string customerPlateId = row[def.PlateCol]?.ToString()?.Trim() ?? "";
                        string rawWell = row[def.WellCol]?.ToString()?.Trim() ?? "";
                        string? well = NormalizeAndValidateWell(rawWell);
                        if (well == null)
                        {
                            skippedCount++;
                            continue;
                        }
                        // Add this warning after normalization
                        if (!string.IsNullOrEmpty(rawWell) && well != null && !rawWell.Equals(well, StringComparison.OrdinalIgnoreCase))
                        {
                            correctedWell++;
                        }
                        if (string.IsNullOrWhiteSpace(rawWell) || string.IsNullOrWhiteSpace(customerPlateId))
                        {
                            skippedCount++;
                            continue;
                        }

                        if (string.IsNullOrWhiteSpace(sampleWellName))
                        {
                            sampleWellName = $"BLANK_{customerPlateId}_{well?.Replace(" ", "")}";
                            blankCount++;
                        }

                        sampleList.Add((sampleWellName, customerPlateId, well ?? "", customerPlateId));
                    }
                });

                var uniquePlateIds = sampleList.Select(x => x.CustomerPlateId).Distinct().ToList();
                var plateIdMap = new Dictionary<string, (string IDText, string ITKPlateId)>();

                using var conn = new SqlConnection(connectionString);
                conn.Open();

                using var cmd = conn.CreateCommand();
                var paramNames = new List<string>();

                for (int i = 0; i < uniquePlateIds.Count; i++)
                {
                    string param = $"@p{i}";
                    cmd.Parameters.AddWithValue(param, uniquePlateIds[i]);
                    paramNames.Add(param);
                }

                cmd.Parameters.AddWithValue("@orderId", orderId);
                cmd.CommandText = $@"
            SELECT SAMPLE_NAME, ID_TEXT, ITK_PLATE_ID,ID_NUMERIC
            FROM sample
            WHERE SAMPLE_NAME IN ({string.Join(",", paramNames)})
            AND JOB_NAME = @orderId
            order by ID_NUMERIC";
                AnsiConsole.Status()
                    .Start("Reading from database...", ctx =>
                    {


                        using (var sqlReader = cmd.ExecuteReader())
                        {
                            while (sqlReader.Read())
                            {
                                string sampleName = sqlReader.GetString(0);
                                string idText = sqlReader.GetString(1);
                                string plateId = sqlReader.GetString(2);
                                plateIdMap[sampleName] = (idText, plateId);
                            }
                        }
                    });
                var enrichedSampleList = sampleList
                    .Where(s => plateIdMap.ContainsKey(s.CustomerPlateId))
                    .Select(s => (
                        s.SampleWellName,
                        s.CustomerPlateId,
                        s.Well,
                        PlateIdText: plateIdMap[s.CustomerPlateId].IDText
                    ))
                    .ToList();

                if(enrichedSampleList == null || enrichedSampleList.Count == 0){
                AnsiConsole.MarkupLine($"[red] Export aborted:[/] No samples found for {orderId}");
                return;
                }
                bool excel = true;
                // Ask for XML output path
                string xmlOutputFile = Path.Combine(outpath, $"{orderId}-Kraken_masterplates.xml");
                GenerateKrakenMasterPlatesXml(
                    projectType: orderId,
                    sampleList: enrichedSampleList,
                    outputPath: xmlOutputFile,
                    selectedType,
                    excel,
                    dartorder
                );
                PrintImportSummary(sheet.Rows.Count - def.StartRow, sampleList.Count, blankCount, correctedWell, skippedCount, xmlOutputFile);
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]An error occurred:[/] {ex.Message}");
            }
        }

        private static void XmlFromDbWithoutWells()
        {
            string orderId = AnsiConsole.Ask<string>("Enter [green]Order ID[/] (e.g. SE-25-0130):");
            string outpath = PromptForOutputDirectory();
            int wellsPerPlate = PromptForNumberOfWells();
            string query = @"
        SELECT ID_TEXT, SAMPLE_NAME, ID_NUMERIC
        FROM sample
        WHERE ENTITY_TEMPLATE_ID = 'PCR_SINGLE_SEED_PLATE' AND JOB_NAME = @order_id
        ORDER BY ID_NUMERIC
    ";

            try
            {
                using var conn = new SqlConnection(connectionString);
                conn.Open();

                using var cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@order_id", orderId!);

                var sampleList = new List<(string SampleWellName, string CustomerPlateId, string Well, string PlateIdText)>();
                AnsiConsole.Status()
                .Start("Reading from datase...", ctx =>
                {

                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        string plateIdText = reader.GetString(0);
                        string customerPlateId = reader.GetString(1);
                        // For each plate, generate wells
                        for (int i = 1; i <= wellsPerPlate; i++)
                        {
                            string well = GetWellFromIndex(i);
                            string sampleWellName = $"{plateIdText}_{well}";
                            sampleList.Add((sampleWellName, customerPlateId, well, plateIdText));
                        }
                    }
                }
                });

                if(sampleList == null || sampleList.Count == 0){
                AnsiConsole.MarkupLine($"[red] Export aborted:[/] No samples found for {orderId}");
                return;
                }
                // XML Export
                string xmlOutputFile = Path.Combine(outpath, $"{orderId}-Kraken_masterplates.xml");
                GenerateKrakenMasterPlatesXml(
                    projectType: orderId,
                    sampleList: sampleList,
                    outputPath: xmlOutputFile
                );
                AnsiConsole.MarkupLine($"[green] Export complete:[/] [underline]{xmlOutputFile}[/]");
            }
            catch (Exception ex)
            {
                AnsiConsole.WriteException(ex,
                    ExceptionFormats.ShortenPaths | ExceptionFormats.ShortenTypes);
            }
        }
        private static string GetWellFromIndex(int index)
        {
            int row = (index - 1) / 12;
            int col = (index - 1) % 12 + 1;
            char rowChar = (char)('A' + row);
            return $"{rowChar}{col:D2}";
        }

        static string PromptForFilePath(string promptText = "[green]Enter path to Excel file[/] [grey](drag & drop works)[/]:")
        {
            string input = AnsiConsole.Prompt(
                new TextPrompt<string>($"{promptText}")
                    .Validate(path =>
                    {
                        var trimmed = path.Trim('"');
                        return File.Exists(trimmed)
                            ? ValidationResult.Success()
                            : ValidationResult.Error("[red]Invalid file path. Try again.[/]");
                    })
            );
            return input.Trim('"');
        }
        private static void PrintImportSummary(int totalRows, int validSamples, int blankCount, int correctedWells, int skippedCount, string outputFile)
        {
            AnsiConsole.MarkupLine("\n[bold underline green]Import Summary[/]");
            var table = new Spectre.Console.Table();
            table.AddColumn("Metric");
            table.AddColumn("Value");
            table.AddRow("Total rows processed", totalRows.ToString());
            table.AddRow("Valid samples", validSamples.ToString());
            table.AddRow("Blanks auto-named", blankCount.ToString());
            table.AddRow("Corrected well positions", correctedWells.ToString());
            table.AddRow("Skipped rows", skippedCount.ToString());
            table.AddRow("Export complete", outputFile);
            AnsiConsole.Write(table);
        }
        static string PromptForOutputDirectory(string promptText = "Enter output path [grey](leave empty for current directory)[/]:")
        {
            string input = AnsiConsole.Prompt(
                new TextPrompt<string>($"{promptText}")
                    .AllowEmpty()
                    .PromptStyle("green")
            );

            string trimmed = input.Trim('"').Trim();

            return string.IsNullOrWhiteSpace(trimmed)
                ? Directory.GetCurrentDirectory()
                : trimmed;
        }
        static int PromptForNumberOfWells()
        {
            int wellsPerPlate = AnsiConsole.Prompt(
                new TextPrompt<int>("Enter number of wells per plate [grey]default[/]:")
                    .AllowEmpty()
                    .DefaultValue(88)
                    .Validate(val => val > 0 ? ValidationResult.Success() : ValidationResult.Error("[red]Must be positive[/]"))
            );
            return wellsPerPlate;


        }
        private static string NormalizeAndValidateWell(string well)
        {
            if (string.IsNullOrWhiteSpace(well))
                return null;

            well = well.Trim().ToUpper();

            // Fix common typos: e.g., "HO1" -> "H01"
            if (well.Length == 3 && well[1] == 'O' && char.IsDigit(well[2]))
                well = $"{well[0]}0{well[2]}";

            // Accept "A1" as "A01"
            var match = System.Text.RegularExpressions.Regex.Match(well, @"^([A-H])0?(\d{1,2})$");
            if (match.Success)
            {
                string row = match.Groups[1].Value;
                int col = int.Parse(match.Groups[2].Value);
                if (col >= 1 && col <= 12)
                    return $"{row}{col:D2}";
            }

            // If already in correct format (A01-H12)
            if (System.Text.RegularExpressions.Regex.IsMatch(well, @"^[A-H][0-9]{2}$"))
                return well;

            // Not valid
            return null;
        }


        private static void GenerateKrakenMasterPlatesXml(
            string projectType,
            List<(string SampleWellName, string CustomerPlateId, string Well, string PlateIdText)> sampleList,
            string outputPath,
            string? selectedType = null,
            bool? excel = false,
            bool? dartorder = false)

        {
            var plates = sampleList
                .GroupBy(s => s.PlateIdText)
                .ToDictionary(g => g.Key, g => g.ToList());

            var masterPlates = new XElement("MASTER_PLATES");

            void AddPlates(string suffix)
            {
                foreach (var plate in plates)
                {
                    var wells = new Dictionary<string, (string SampleWellName, string CustomerPlateId)>();
                    foreach (var s in plate.Value)
                        wells[s.Well] = (s.SampleWellName + suffix, s.CustomerPlateId + suffix);

                    var allWells = new List<string>();
                    for (char row = 'A'; row <= 'H'; row++)
                        for (int col = 1; col <= 12; col++)
                            allWells.Add($"{row}{col:D2}");

                    // Ensure all wells are present (fill with BLANK or NTC as needed)
                    foreach (var well in allWells)
                    {
                        if (dartorder == false)
                        {

                            if (!wells.ContainsKey(well))
                            {
                                if (well == "H11" || well == "H12")
                                    wells[well] = ($"NTC_{plate.Key}_{well}{suffix}", plate.Value.First().CustomerPlateId + suffix);
                                else
                                    wells[well] = ($"BLANK_{plate.Key}_{well}{suffix}", plate.Value.First().CustomerPlateId + suffix);
                            }
                            // After filling wells dictionary
                            wells["H11"] = ($"NTC_{plate.Key}_H11{suffix}", plate.Value.First().CustomerPlateId + suffix);
                            wells["H12"] = ($"NTC_{plate.Key}_H12{suffix}", plate.Value.First().CustomerPlateId + suffix);
                        }
                        else
                        {

                            if (!wells.ContainsKey(well))
                            {
                                if (well == "G12" || well == "H12")
                                    wells[well] = ($"NTC_{plate.Key}_{well}{suffix}", plate.Value.First().CustomerPlateId + suffix);
                                else
                                    wells[well] = ($"BLANK_{plate.Key}_{well}{suffix}", plate.Value.First().CustomerPlateId + suffix);
                            }
                            // After filling wells dictionary
                            wells["G12"] = ($"NTC_{plate.Key}_G12{suffix}", plate.Value.First().CustomerPlateId + suffix);
                            wells["H12"] = ($"NTC_{plate.Key}_H12{suffix}", plate.Value.First().CustomerPlateId + suffix);
                        }
                    }

                    var wellElements = allWells.Select(well =>
                    {
                        var (sample, customerPlateId) = wells[well];
                        var isNTC = dartorder == true
                            ? (well == "G12" || well == "H12")
                            : (well == "H11" || well == "H12");
                        return new XElement("Well",
                            new XElement("Location", well),
                            new XElement("Subject_id", sample),
                            (excel ?? false) ? new XElement("long_id", $"{customerPlateId}_{plate.Key + suffix}") : new XElement("long_id", $"{customerPlateId}"),
                            isNTC ? new XElement("class", "NTC") : null,
                            sample.IndexOf("BLANK_", StringComparison.OrdinalIgnoreCase) >= 0
                            || sample.IndexOf("EMPTY", StringComparison.OrdinalIgnoreCase) >= 0
                                ? new XElement("class", "EMPTY") : null
                        );
                    });

                    masterPlates.Add(
                        new XElement("MASTER",
                            new XElement("Plate", plate.Key + suffix),
                            new XElement("Density", "96"),
                            new XElement("Alias", ""),
                            new XElement("Barcode", ""),
                            wellElements
                        )
                    );
                }
            }

            AnsiConsole.Progress()
                .Start(ctx =>
                {
                    var task = ctx.AddTask("[green]Generating XML...[/]", maxValue: plates.Count * ((selectedType == "Verification") ? 2 : 1));
                    if (selectedType == "Verification")
                    {
                        AddPlates("_d1");
                        task.Increment(plates.Count);
                        AddPlates("_d2");
                        task.Increment(plates.Count);
                    }
                    else
                    {
                        AddPlates("");
                        task.Increment(plates.Count);
                    }
                });

            var xml =
                new XElement("LIMS",
                    masterPlates
                );

            var doc = new XDocument(new XDeclaration("1.0", "utf-8", "yes"), xml);
            AnsiConsole.Status()
            .Start("Saving XML file...", ctx =>
            {
                doc.Save(outputPath);
            });
        }
    }

}

