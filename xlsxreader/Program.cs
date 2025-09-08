using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using xlsxreader.Models;

public class XlsxReader
{
    private Dictionary<int, string> sharedStrings = new Dictionary<int, string>();
    private List<List<string>> rowsData = new List<List<string>>();
    internal string? filePath;
    internal string? sheetName;
    private Stopwatch stopwatch = new Stopwatch();
    static void Main(string[] args)
    {
        string file = "articoli.xlsx";//args[0];
        string sheet = "Articoli";//args.Length > 1 ? args[1] : "sheet1";
        string format = "model";// args.Length > 2 ? args[2] : "console";
        string output = "output.csv";//args.Length > 3 ? args[3] : "";

        var reader = new XlsxReader();
        reader.filePath = file;
        reader.sheetName = sheet;
        reader.ProcessFile(format, output);
    }

    public void ProcessFile(string exportFormat, string outputFile)
    {
        try
        {
            stopwatch.Start();

            Console.WriteLine("=== XLSX Reader Report ===");
            Console.WriteLine("File: " + Path.GetFileName(filePath));
            Console.WriteLine("Full path: " + Path.GetFullPath(filePath));
            Console.WriteLine("File size: " + new FileInfo(filePath).Length.ToString("N0") + " bytes");
            Console.WriteLine("Sheet: " + sheetName);
            Console.WriteLine("Started at: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            Console.WriteLine();

            ReadXlsxFile();

            stopwatch.Stop();

            // Statistics
            int totalRows = rowsData.Count;
            int totalCells = 0;
            foreach (var row in rowsData) totalCells += row.Count;
            int maxColumns = 0;
            foreach (var row in rowsData)
                if (row.Count > maxColumns) maxColumns = row.Count;

            Console.WriteLine("=== Statistics ===");
            Console.WriteLine("Reading time: " + stopwatch.ElapsedMilliseconds + " ms");
            Console.WriteLine("Total rows: " + totalRows.ToString("N0"));
            Console.WriteLine("Maximum columns: " + maxColumns);
            Console.WriteLine("Total cells: " + totalCells.ToString("N0"));
            Console.WriteLine("Total cols: " + sheetName);
            Console.WriteLine("Shared strings count: " + sharedStrings.Count.ToString("N0"));
            Console.WriteLine();

            // Export or display data
            switch (exportFormat.ToLower())
            {
                case "model":
            //      DisplayData();

                    ExportToModel(outputFile);
                    break;
                case "csv":
                    ExportToCsv(outputFile);
                    break;
                case "txt":
                    ExportToTxt(outputFile);
                    break;
                default:
                    DisplayData();
                    break;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("ERROR: " + ex.Message);
            Environment.Exit(1);
        }
    }

    private void ReadXlsxFile()
    {
        Console.WriteLine("Opening XLSX file...");

        using (ZipArchive archive = ZipFile.OpenRead(filePath))
        {
            // Load shared strings
            Console.WriteLine("Loading shared strings...");
            LoadSharedStrings(archive);

            Console.WriteLine("Resolving sheet path...");                                   //TEST
            string sheetPath = GetSheetPath(archive, sheetName);                            //TEST
            Console.WriteLine("Loading worksheet: " + sheetName + " (" + sheetPath + ")");  //TEST

            // Load worksheet data
            Console.WriteLine("Loading worksheet: " + sheetName + "...");
            LoadWorksheetData(archive);
        }
    }
    private string GetSheetPath(ZipArchive archive, string sheetName)
    {
        var workbookEntry = archive.GetEntry("xl/workbook.xml");
        if (workbookEntry == null) throw new Exception("workbook.xml non trovato");

        using (var reader = new StreamReader(workbookEntry.Open()))
        {
            var doc = XDocument.Load(reader);
            var ns = doc.Root.Name.Namespace;

            var sheets = doc.Descendants(ns + "sheet");
            foreach (var sheet in sheets)
            {
                var nameAttr = sheet.Attribute("name");
                var sheetIdAttr = sheet.Attribute("sheetId");
                if (nameAttr != null && sheetIdAttr != null &&
                    nameAttr.Value.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                {
                    string path = $"xl/worksheets/sheet{sheetIdAttr.Value}.xml";
                    if (archive.GetEntry(path) != null)
                        return path;
                }
            }
        }

        throw new Exception($"Foglio '{sheetName}' non trovato in workbook.xml");
    }

    /* private void LoadSharedStrings(ZipArchive archive)
       {
           var sharedEntry = archive.GetEntry("xl/sharedStrings.xml");
           if (sharedEntry != null)
           {
               using (var reader = new StreamReader(sharedEntry.Open()))
               {
                   var doc = XDocument.Load(reader);
                   var strings = doc.Descendants().Where(e => e.Name.LocalName == "t");
                   int index = 0;
                   foreach (var s in strings)
                   {
                       sharedStrings[index++] = s.Value;
                   }
               }
           }
       }
   */
    private void LoadSharedStrings(ZipArchive archive)
    {
        var sharedEntry = archive.GetEntry("xl/sharedStrings.xml");
        if (sharedEntry != null)
        {
            using (var reader = new StreamReader(sharedEntry.Open()))
            {
                var doc = XDocument.Load(reader);
                var items = doc.Descendants().Where(e => e.Name.LocalName == "si");
                int index = 0;
                foreach (var item in items)
                {
                    var textParts = item.Descendants().Where(e => e.Name.LocalName == "t");
                    sharedStrings[index++] = string.Join("", textParts.Select(t => t.Value));
                }
            }
        }
    }

    /*
        private void LoadWorksheetData(ZipArchive archive)
        {
            // Try different sheet paths
            string[] possiblePaths = {
                "xl/worksheets/" + sheetName + ".xml",
                "xl/worksheets/sheet" + sheetName + ".xml",
                "xl/worksheets/sheet1.xml"
            };

            ZipArchiveEntry sheetEntry = null;
            foreach (var path in possiblePaths)
            {
                sheetEntry = archive.GetEntry(path);
                if (sheetEntry != null) break;
            }

            if (sheetEntry == null)
            {
                Console.WriteLine("Available sheets:");
                foreach (var entry in archive.Entries)
                {
                    if (entry.FullName.StartsWith("xl/worksheets/") && entry.FullName.EndsWith(".xml"))
                    {
                        Console.WriteLine("  - " + entry.FullName);
                    }
                }
                throw new Exception("Sheet '" + sheetName + "' not found");
            }

            using (var reader = new StreamReader(sheetEntry.Open()))
            {
                var doc = XDocument.Load(reader);
                var rows = doc.Descendants().Where(e => e.Name.LocalName == "row");

                foreach (var row in rows)
                {
                    List<string> rowValues = new List<string>();
                    var cells = row.Elements().Where(e => e.Name.LocalName == "c");

                    foreach (var cell in cells)
                    {
                        string cellType = cell.Attribute("t") != null ? cell.Attribute("t").Value : null;
                        var valueElement = cell.Elements().FirstOrDefault(e => e.Name.LocalName == "v");

                        if (valueElement != null)
                        {
                            string rawValue = valueElement.Value;
                            if (cellType == "s")
                            {
                                int sIndex;
                                if (int.TryParse(rawValue, out sIndex) && sharedStrings.ContainsKey(sIndex))
                                {
                                    rowValues.Add(sharedStrings[sIndex]);
                                }
                                else
                                {
                                    rowValues.Add(rawValue);
                                }
                            }
                            else
                            {
                                rowValues.Add(rawValue);
                            }
                        }
                        else
                        {
                          rowValues.Add(""); // empty cell
                        }

                        Console.WriteLine(valueElement);
                    }

                    if (rowValues.Count > 0)
                    {
                        rowsData.Add(rowValues);
                    }
                }
            }
        }
    

    private void LoadWorksheetData(ZipArchive archive)
    {
        // Try different sheet paths
        string[] possiblePaths = {
        "xl/worksheets/" + sheetName + ".xml",
        "xl/worksheets/sheet" + sheetName + ".xml",
        "xl/worksheets/sheet1.xml"
    };

        ZipArchiveEntry sheetEntry = null;
        foreach (var path in possiblePaths)
        {
            sheetEntry = archive.GetEntry(path);
            if (sheetEntry != null) break;
        }

        if (sheetEntry == null)
        {
            Console.WriteLine("Available sheets:");
            foreach (var entry in archive.Entries)
            {
                if (entry.FullName.StartsWith("xl/worksheets/") && entry.FullName.EndsWith(".xml"))
                {
                    Console.WriteLine("  - " + entry.FullName);
                }
            }
            throw new Exception("Sheet '" + sheetName + "' not found");
        }

        using (var reader = new StreamReader(sheetEntry.Open()))
        {
            var doc = XDocument.Load(reader);
            var rows = doc.Descendants().Where(e => e.Name.LocalName == "row");

            foreach (var row in rows)
            {
                List<string> rowValues = new List<string>();
                var cells = row.Elements().Where(e => e.Name.LocalName == "c");

                foreach (var cell in cells)
                {
                    string cellType = cell.Attribute("t") != null ? cell.Attribute("t").Value : null;
                    var valueElement = cell.Elements().FirstOrDefault(e => e.Name.LocalName == "v");

                    if (valueElement != null)
                    {
                        string rawValue = valueElement.Value;
                        //                      Console.WriteLine($"Cell value: {rawValue} Type: {cellType} Row: " + rowc.ToString() + " Col: " + rowd.ToString()); // Log the cell value

                        if (cellType == "s")
                        {
                            int sIndex;
                            if (int.TryParse(rawValue, out sIndex) && sharedStrings.ContainsKey(sIndex))
                            {
                                rowValues.Add(sharedStrings[sIndex]);
                            }
                            else
                            {
                                rowValues.Add(rawValue);
                            }
                        }
                        else
                        {
                            rowValues.Add(rawValue);
                        }
                    }
                    else
                    {
                        //                      Console.WriteLine("Empty cell found " + rowc.ToString()); // Log empty cell
                        rowValues.Add(""); // empty cell
                    }
                }

                // Add every row, even if it contains empty cells
                rowsData.Add(rowValues);
            }
        }
    }
*/
    private void LoadWorksheetData(ZipArchive archive)
    {
        // Try different sheet paths
        string[] possiblePaths = {
        "xl/worksheets/" + sheetName + ".xml",
        "xl/worksheets/sheet" + sheetName + ".xml",
        "xl/worksheets/sheet1.xml"
    };

        ZipArchiveEntry sheetEntry = null;
        foreach (var path in possiblePaths)
        {
            sheetEntry = archive.GetEntry(path);
            if (sheetEntry != null) break;
        }

        if (sheetEntry == null)
        {
            Console.WriteLine("Available sheets:");
            foreach (var entry in archive.Entries)
            {
                if (entry.FullName.StartsWith("xl/worksheets/") && entry.FullName.EndsWith(".xml"))
                {
                    Console.WriteLine("  - " + entry.FullName);
                }
            }
            throw new Exception("Sheet '" + sheetName + "' not found");
        }

        using (var reader = new StreamReader(sheetEntry.Open()))
        {
            var doc = XDocument.Load(reader);
            var rows = doc.Descendants().Where(e => e.Name.LocalName == "row");

            foreach (var row in rows)
            {
                List<string> rowValues = new List<string>();
                var cells = row.Elements().Where(e => e.Name.LocalName == "c");

                foreach (var cell in cells)
                {
                    string cellType = cell.Attribute("t") != null ? cell.Attribute("t").Value : null;
                    var valueElement = cell.Elements().FirstOrDefault(e => e.Name.LocalName == "v");

                    if (valueElement != null)
                    {
                        string rawValue = valueElement.Value;
                        //                      Console.WriteLine($"Cell value: {rawValue} Type: {cellType} Row: " + rowc.ToString() + " Col: " + rowd.ToString()); // Log the cell value

                        if (cellType == "s")
                        {
                            int sIndex;
                            if (int.TryParse(rawValue, out sIndex) && sharedStrings.ContainsKey(sIndex))
                            {
                                rowValues.Add(sharedStrings[sIndex]);
                            }
                            else
                            {
                                rowValues.Add(rawValue);
                            }
                        }
                        else
                        {
                            rowValues.Add(rawValue);
                        }
                    }
                    else
                    {
                        //                      Console.WriteLine("Empty cell found " + rowc.ToString()); // Log empty cell
                        rowValues.Add(""); // empty cell
                    }
                }

                // Add every row, even if it contains empty cells
                rowsData.Add(rowValues);
            }
        }
    }

    private void ExportToCsv(string outputFile)
    {
        if (string.IsNullOrEmpty(outputFile))
        {
            outputFile = Path.ChangeExtension(filePath, ".csv");
        }

        Console.WriteLine("Exporting to CSV: " + outputFile);

        using (var writer = new StreamWriter(outputFile, false, Encoding.UTF8))
        {
            foreach (var row in rowsData)
            {
                List<string> csvRow = new List<string>();
                foreach (var cell in row)
                {
                    if (cell.Contains(",") || cell.Contains("\"") || cell.Contains("\n"))
                    {
                        csvRow.Add("\"" + cell.Replace("\"", "\"\"") + "\"");
                    }
                    else
                    {
                        csvRow.Add(cell);
                    }
                }
                writer.WriteLine(string.Join(",", csvRow));
            }
        }

        Console.WriteLine("CSV export completed: " + new FileInfo(outputFile).Length.ToString("N0") + " bytes");
    }

    private void DisplayData()
    {
        Console.WriteLine("=== Data Preview (First 10 rows) ===");
        int displayRows = Math.Min(10, rowsData.Count);
        int maxColumns = rowsData.Max(r => r.Count);

        for (int i = 0; i < displayRows; i++)
        {
            var row = rowsData[i];
            List<string> displayCols = new List<string>();
            for (int j = 0; j < Math.Min(maxColumns, row.Count); j++)
            {
                displayCols.Add(row[j]);
            }
            Console.WriteLine("Row " + (i + 1).ToString("D3") + ": " + string.Join(" | ", displayCols));

            if (row.Count > 10)
            {
                Console.WriteLine("      ... and " + (row.Count - 10) + " more columns ...");
            }
        }

        if (rowsData.Count > displayRows)
        {
            Console.WriteLine("... and " + (rowsData.Count - displayRows) + " more rows");
        }
    }

    private void ExportToModel(string outputFile)
    {
        List<Articoli> articolis = new List<Articoli>();
        int riga = -1;
        foreach (var row in rowsData)
        {
            List<string> csvRow = new List<string>();
            riga++;

            Console.WriteLine("Inizio Riga " + riga.ToString("N0"));
            Console.WriteLine("Riga: " + (riga).ToString("D3") + ": " + string.Join(" | ", row[riga]));

            if (riga > 17) riga = 0;

            articolis.Add(new Articoli
            {
                CodificatoCt = row[0],
                IdArticolo = row[1],
                Fornitore = row[2],
                SupplierCode = row[3],
                Ean = row[4],
                CodMatForn = row[5],
                Descrizione = row[6],
                Microcategory = row[7],
                CodSubcategory = row[8],
                Subcategory = row[9],
                Taric = row[10],
                Country = row[11],
                Stato = row[12],
                Brand = row[13],
                Taglia = row[14],
                Colore = row[15],
                Gender = row[16],
                CostoDiAcquisto = row[17]
            });
        }

        File.WriteAllText("models_out.txt", JsonSerializer.Serialize(articolis, new JsonSerializerOptions() { WriteIndented = true }));

        Console.WriteLine("Fine Riga " + riga.ToString("N0"));
    }
    /*
        private void ExportToCsv(string outputFile)
        {
            if (string.IsNullOrEmpty(outputFile))
            {
                outputFile = Path.ChangeExtension(filePath, ".csv");
            }

            Console.WriteLine("Exporting to CSV: " + outputFile);

            using (var writer = new StreamWriter(outputFile, false, Encoding.UTF8))
            {
                foreach (var row in rowsData)
                {
                    List<string> csvRow = new List<string>();
                    foreach (var cell in row)
                    {
                        if (cell.Contains(",") || cell.Contains("\"") || cell.Contains("\n"))
                        {
                            csvRow.Add("\"" + cell.Replace("\"", "\"\"") + "\"");
                        }
                        else
                        {
                            csvRow.Add(cell);
                        }
                    }
                    writer.WriteLine(string.Join(",", csvRow));
                }
            }

            Console.WriteLine("CSV export completed: " + new FileInfo(outputFile).Length.ToString("N0") + " bytes");
        }
    */
    private void ExportToTxt(string outputFile)
    {
        if (string.IsNullOrEmpty(outputFile))
        {
            outputFile = Path.ChangeExtension(filePath, ".txt");
        }

        Console.WriteLine("Exporting to TXT: " + outputFile);

        using (var writer = new StreamWriter(outputFile, false, Encoding.UTF8))
        {
            foreach (var row in rowsData)
            {
                writer.WriteLine(string.Join("\t", row));
            }
        }

        Console.WriteLine("TXT export completed: " + new FileInfo(outputFile).Length.ToString("N0") + " bytes");
    }
}