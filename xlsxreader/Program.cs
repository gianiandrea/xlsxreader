//
// XlsxReader
// Andrea Giani / Marina Lacetera - v0.31
//
// # Sintassi
// XlsxReader.exe data.xlsx json output.json true
// XlsxReader.exe data.xlsx csv 
// XlsxReader.exe data.xlsx console
//

using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.Json;
using System.Threading;
using System.Xml;
using System.Data.Common;

namespace XlsxReader
{
    class Program
    {
        static void Main(string[] args)
        {
			Stopwatch stopWatch = new Stopwatch();
			Process processo = Process.GetCurrentProcess();
            string processName = processo.ProcessName;

            Console.WriteLine("=======================================");
            Console.WriteLine("Excel XLSX Reader - Versione XML Nativa");
            Console.WriteLine("=======================================");

            string filePath = "";
            string outputFormat = "";
            string outputPath = "";
            bool showHeaders = false;

            // Gestione argomenti da linea di comando
            if (args.Length > 0)
            {
                filePath = args[0];

                if (args.Length > 1)
                    outputFormat = args[1].ToLower();

                if (args.Length > 2)
                    outputPath = args[2];

                if (args.Length > 3)
                    showHeaders = args[3].ToLower() == "true" || args[3] == "1";
            }
            else
            {
                Console.Write("Nome del file di Input(XLSX) mancante: ");
                filePath = Console.ReadLine();

				if (filePath == "") {

					string currentDirectory = Directory.GetCurrentDirectory();
					string[] xlsxFiles = Directory.GetFiles(currentDirectory, "*.xlsx");

					if (xlsxFiles.Length > 0)
					{
						filePath = xlsxFiles[0];
						Console.WriteLine("Autoselezione file .xlsx trovato: " + filePath);
					}
				}

                Console.WriteLine("\nFormati di output disponibili:");
                Console.WriteLine("1. console - Visualizza nella console (default)");
                Console.WriteLine("2. json - Esporta in formato JSON");
                Console.WriteLine("3. csv - Esporta in formato CSV");
                Console.WriteLine("4. txt - Esporta in formato testo");
                Console.WriteLine("5. sql - Formato non supportato");

                Console.Write("Selezionare il formato di output [console default]: ");
                var formatInput = Console.ReadLine();
                outputFormat = string.IsNullOrWhiteSpace(formatInput) ? "console" : formatInput.ToLower();

				if (outputFormat == "json" || outputFormat == "csv" || outputFormat == "txt") {
					Console.Write("Percorso file di output (lascia vuoto per generare automaticamente): ");
					outputPath = Console.ReadLine();
				}

                Console.Write("Includere la prima riga come intestazioni? (s/n) [n]: ");
                var headerInput = Console.ReadLine();
                showHeaders = headerInput?.ToLower() == "s" || headerInput?.ToLower() == "si";
            }

            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
         //     Console.WriteLine("File non trovato o percorso non valido!");
                Console.WriteLine("\nXLSXReader\n");
                Console.WriteLine("Sintassi: XlsxReader.exe <percorso_file> [formato_output] [percorso_output] [mostra_headers]");
                Console.WriteLine("\nEsempi:");
                Console.WriteLine("XlsxReader.exe data.xlsx json output.json true");
                Console.WriteLine("XlsxReader.exe data.xlsx csv");
                Console.WriteLine("XlsxReader.exe data.xlsx console\n");
                Console.WriteLine("Premi un tasto per uscire...");
                Console.ReadKey();
                return;
            }

			stopWatch.Start();
            TimeSpan startCpuTime = processo.TotalProcessorTime;
            DateTime startTime = DateTime.Now;

            try
            {
                // Leggi tutti i fogli disponibili
                var fogli = GetFogliDisponibili(filePath);

                if (fogli.Count == 0)
                {
                    Console.WriteLine("Nessun foglio trovato nel file Excel.");
                    return;
                }

                Console.WriteLine($"\nFogli disponibili ({fogli.Count}):");
                for (int i = 0; i < fogli.Count; i++)
                {
                    Console.WriteLine($"{i + 1}. {fogli[i].Nome}");
                }

                // Selezione foglio
                int sceltaFoglio = 0;
                if (fogli.Count > 1 && outputFormat == "console")
                {
                    Console.Write($"\nSeleziona il foglio da leggere (1-{fogli.Count}) [1]: ");
                    var input = Console.ReadLine();
                    if (!int.TryParse(input, out sceltaFoglio) || sceltaFoglio < 1 || sceltaFoglio > fogli.Count)
                    {
                        sceltaFoglio = 1;
                    }
                }
                else
                {
                    sceltaFoglio = 1;
                }

                var foglioSelezionato = fogli[sceltaFoglio - 1];
                Console.WriteLine($"\nLettura del foglio: {foglioSelezionato.Nome}");
                Console.WriteLine(new string('=', 50));

                // Leggi i dati del foglio
                var dati = LeggiExcel(filePath, foglioSelezionato.Id);

                if (dati.Count == 0)
                {
                    Console.WriteLine("Nessun dato trovato nel foglio selezionato.");
                    return;
                }

                // Processa i dati in base al formato richiesto
                switch (outputFormat)
                {
                    case "json":
                        EsportaJson(dati, outputPath, filePath, showHeaders, foglioSelezionato.Nome);
                        break;

                    case "csv":
                        EsportaCsv(dati, outputPath, filePath, showHeaders);
                        break;

                    case "txt":
                        EsportaTxt(dati, outputPath, filePath, showHeaders);
                        break;

                    case "sql":
         //             string connStr = "Server=localhost;Database=NomeDB;Trusted_Connection=True;";
         //             bool hasHeaders = false;
         //             EsportaSql(dati, connStr, "NomeTabella", hasHeaders);
                        break;

                    default:
                        MostraConsole(dati, showHeaders);
                        break;
                }

                Console.WriteLine($"\nReading completed.\n Total row: {dati.Count}");
				Console.WriteLine("Full path: " + Path.GetFullPath(filePath));
                Console.WriteLine("File size: " + new FileInfo(filePath).Length.ToString("N0") + " bytes");

				stopWatch.Stop();

				// Get the elapsed time as a TimeSpan value.
				TimeSpan ts = stopWatch.Elapsed;

                // Format and display the TimeSpan value.
                string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
					ts.Hours, ts.Minutes, ts.Seconds,
					ts.Milliseconds / 10);
				Console.WriteLine("Time elapsed: " + elapsedTime);	
				
				// Display Memory
				long memoriaPrivata = processo.PrivateMemorySize64;
				long memoriaFisica = processo.WorkingSet64;
				Console.WriteLine($"Private memory: {memoriaPrivata / 1024} KB");				//byte allocati esclusivamente dal processo
				Console.WriteLine($"Physical memory (working set): {memoriaFisica / 1024} KB"); //quantità di memoria fisica attualmente usata

                // CPU Usage
                TimeSpan endCpuTime = processo.TotalProcessorTime;
                DateTime endTime = DateTime.Now;

                double cpuUsedMs = (endCpuTime - startCpuTime).TotalMilliseconds;
                double totalMsPassed = (endTime - startTime).TotalMilliseconds;

                int processorCount = Environment.ProcessorCount;
                double cpuUsageTotal = (cpuUsedMs / (totalMsPassed * processorCount)) * 100;

                Console.WriteLine($"CPU used: {cpuUsageTotal:F2}%");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Errore durante la lettura del file: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
            }

            Console.WriteLine("\nPremi un tasto per uscire...");
            Console.ReadKey();
        }

        public class FoglioInfo
        {
            public string Nome { get; set; }
            public string Id { get; set; }
        }

        static List<FoglioInfo> GetFogliDisponibili(string filePath)
        {
            var fogli = new List<FoglioInfo>();

            using var archive = ZipFile.OpenRead(filePath);
            var workbookEntry = archive.GetEntry("xl/workbook.xml");

            if (workbookEntry != null)
            {
                using var workbookStream = workbookEntry.Open();
                var workbookDoc = LoadXmlSafe(workbookStream);

                var nsManager = new XmlNamespaceManager(workbookDoc.NameTable);
                nsManager.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

                var sheets = workbookDoc.SelectNodes("//x:sheet", nsManager);
                foreach (XmlNode sheet in sheets)
                {
                    var nome = sheet.Attributes?["name"]?.Value ?? "Senza nome";
                    var sheetId = sheet.Attributes?["sheetId"]?.Value ?? "1";

                    fogli.Add(new FoglioInfo
                    {
                        Nome = nome,
                        Id = sheetId
                    });
                }
            }

            // Se non troviamo fogli nel workbook.xml, assumiamo che esista sheet1
            if (fogli.Count == 0)
            {
                fogli.Add(new FoglioInfo { Nome = "Sheet1", Id = "1" });
            }

            return fogli;
        }

        static List<List<string>> LeggiExcel(string filePath, string sheetId = "1")
        {
            var risultato = new List<List<string>>();

            using var archive = ZipFile.OpenRead(filePath);

            // Carica le stringhe condivise
            var sharedStrings = CaricaStringheCondivise(archive);

            // Determina il percorso del foglio
            var sheetPath = $"xl/worksheets/sheet{sheetId}.xml";
            var sheetEntry = archive.GetEntry(sheetPath);

            // Se il foglio specificato non esiste, prova con sheet1.xml
            if (sheetEntry == null)
            {
                sheetPath = "xl/worksheets/sheet1.xml";
                sheetEntry = archive.GetEntry(sheetPath);
            }

            if (sheetEntry == null)
            {
                throw new FileNotFoundException($"Foglio non trovato: sheet{sheetId}.xml");
            }

            // Lettura del foglio
            using var sheetStream = sheetEntry.Open();
            var sheetDoc = LoadXmlSafe(sheetStream);

            var nsManager = new XmlNamespaceManager(sheetDoc.NameTable);
            nsManager.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

            var rows = sheetDoc.SelectNodes("//x:row", nsManager);

            foreach (XmlNode row in rows)
            {
                var riga = new List<string>();
                var cells = row.SelectNodes("x:c", nsManager);

                // Ottieni tutte le celle della riga
                var celleOrdinate = new SortedDictionary<int, XmlNode>();

                foreach (XmlNode cell in cells)
                {
                    var riferimento = cell.Attributes?["r"]?.Value;
                    if (!string.IsNullOrEmpty(riferimento))
                    {
                        var colonnaIndex = GetColonnaIndex(riferimento);
                        celleOrdinate[colonnaIndex] = cell;
                    }
                }

                // Se non ci sono celle, aggiungi una riga vuota
                if (celleOrdinate.Count == 0)
                {
                    risultato.Add(new List<string>());
                    continue;
                }

                // Riempi la riga fino alla colonna massima
                var maxColonna = celleOrdinate.Keys.Max();
                for (int col = 0; col <= maxColonna; col++)
                {
                    if (celleOrdinate.ContainsKey(col))
                    {
                        var valore = EstraiValoreCella(celleOrdinate[col], sharedStrings);
                        riga.Add(valore);
                    }
                    else
                    {
                        riga.Add(""); // Cella vuota
                    }
                }

                risultato.Add(riga);
            }

            return risultato;
        }

        static List<string> CaricaStringheCondivise(ZipArchive archive)
        {
            var sharedStrings = new List<string>();
            var sharedStringsEntry = archive.GetEntry("xl/sharedStrings.xml");

            if (sharedStringsEntry != null)
            {
                using var sharedStream = sharedStringsEntry.Open();
                var sharedDoc = LoadXmlSafe(sharedStream);

                var nsManager = new XmlNamespaceManager(sharedDoc.NameTable);
                nsManager.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

                var sis = sharedDoc.SelectNodes("//x:si", nsManager);
                foreach (XmlNode si in sis)
                {
                    // Cerca sia <t> che <r><t> per testo formattato
                    var tNode = si.SelectSingleNode("x:t", nsManager);
                    if (tNode != null)
                    {
                        sharedStrings.Add(tNode.InnerText);
                    }
                    else
                    {
                        // Gestisci rich text (con formattazione)
                        var rNodes = si.SelectNodes("x:r/x:t", nsManager);
                        var richText = "";
                        foreach (XmlNode rNode in rNodes)
                        {
                            richText += rNode.InnerText;
                        }
                        sharedStrings.Add(richText);
                    }
                }
            }

            return sharedStrings;
        }

        static string EstraiValoreCella(XmlNode cell, List<string> sharedStrings)
        {
            var tipo = cell.Attributes?["t"]?.Value;
            var vNode = cell["v"];
            var v = vNode?.InnerText;

            if (string.IsNullOrEmpty(v))
                return "";

            switch (tipo)
            {
                case "s": // Stringa condivisa
                    if (int.TryParse(v, out int idx) && idx < sharedStrings.Count)
                        return sharedStrings[idx];
                    return "";

                case "str": // Stringa inline
                    return v;

                case "inlineStr": // Stringa inline (alternativa)
                    var isNode = cell["is"];
                    return isNode?.InnerText ?? v;

                case "b": // Boolean
                    return v == "1" ? "TRUE" : "FALSE";

                case "d": // Data
                    if (DateTime.TryParse(v, out DateTime data))
                        return data.ToString("dd/MM/yyyy");
                    return v;

                case "e": // Errore
                    return $"#ERROR: {v}";

                default: // Numero o null
                    return v;
            }
        }

        static int GetColonnaIndex(string riferimentoCella)
        {
            var colonnaLetters = "";
            foreach (char c in riferimentoCella)
            {
                if (char.IsLetter(c))
                    colonnaLetters += c;
                else
                    break;
            }

            int risultato = 0;
            for (int i = 0; i < colonnaLetters.Length; i++)
            {
                risultato = risultato * 26 + (colonnaLetters[i] - 'A' + 1);
            }

            return risultato - 1; // Converti in base 0
        }

        static void MostraConsole(List<List<string>> dati, bool showHeaders)
        {
            Console.WriteLine($"Trovate {dati.Count} righe:");
            Console.WriteLine();

            for (int i = 0; i < dati.Count; i++)
            {
                if (i == 0 && showHeaders)
                {
                    Console.WriteLine($"HEADER: {string.Join(" | ", dati[i])}");
                    Console.WriteLine(new string('-', 80));
                }
                else
                {
                    Console.WriteLine($"Riga {(showHeaders ? i : i + 1),3}: {string.Join(" | ", dati[i])}");
                }
            }
        }
 
        static void EsportaJson(List<List<string>> dati, string outputPath, string originalFilePath, bool hasHeaders, string sheetName)
        {
            if (string.IsNullOrWhiteSpace(outputPath))
            {
                var baseName = Path.GetFileNameWithoutExtension(originalFilePath);
                outputPath = $"{baseName}_export.json";
            }

            object jsonData;

            if (hasHeaders && dati.Count > 1)
            {
                var headers = dati[0];
                var rows = dati.Skip(1).ToList();

                var records = new List<Dictionary<string, string>>();
                foreach (var row in rows)
                {
                    var record = new Dictionary<string, string>();
                    for (int i = 0; i < headers.Count; i++)
                    {
                        var key = string.IsNullOrWhiteSpace(headers[i]) ? $"Column_{i + 1}" : headers[i];
                        var value = i < row.Count ? row[i] : "";
                        record[key] = value;
                    }
                    records.Add(record);
                }

                jsonData = new
                {
                    metadata = new
                    {
                        source_file = Path.GetFileName(originalFilePath),
                        sheet_name = sheetName,
                        total_rows = dati.Count,
                        data_rows = rows.Count,
                        has_headers = hasHeaders,
                        export_date = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                    },
                    headers = headers,
                    data = records
                };
            }
            else
            {
                jsonData = new
                {
                    metadata = new
                    {
                        source_file = Path.GetFileName(originalFilePath),
                        sheet_name = sheetName,
                        total_rows = dati.Count,
                        has_headers = hasHeaders,
                        export_date = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                    },
                    data = dati
                };
            }

            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            };

            var jsonString = JsonSerializer.Serialize(jsonData, options);
            File.WriteAllText(outputPath, jsonString);

            Console.WriteLine($"Dati esportati in formato JSON: {outputPath}");
            Console.WriteLine($"Dimensione file: {new FileInfo(outputPath).Length} bytes");
        }

        static void EsportaCsv(List<List<string>> dati, string outputPath, string originalFilePath, bool hasHeaders)
        {
            if (string.IsNullOrWhiteSpace(outputPath))
            {
                var baseName = Path.GetFileNameWithoutExtension(originalFilePath);
                outputPath = $"{baseName}_export.csv";
            }

            using var writer = new StreamWriter(outputPath);

            foreach (var riga in dati)
            {
                var csvLine = string.Join(",", riga.Select(cell =>
                {
                    if (cell.Contains(",") || cell.Contains("\"") || cell.Contains("\n"))
                    {
                        return $"\"{cell.Replace("\"", "\"\"")}\"";
                    }
                    return cell;
                }));
                writer.WriteLine(csvLine);
            }

            Console.WriteLine($"Dati esportati in formato CSV: {outputPath}");
            Console.WriteLine($"Dimensione file: {new FileInfo(outputPath).Length} bytes");
        }

        static void EsportaTxt(List<List<string>> dati, string outputPath, string originalFilePath, bool hasHeaders)
        {
            if (string.IsNullOrWhiteSpace(outputPath))
            {
                var baseName = Path.GetFileNameWithoutExtension(originalFilePath);
                outputPath = $"{baseName}_export.txt";
            }

            using var writer = new StreamWriter(outputPath);

            writer.WriteLine($"Esportazione da: {Path.GetFileName(originalFilePath)}");
            writer.WriteLine($"Data esportazione: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            writer.WriteLine($"Totale righe: {dati.Count}");
            writer.WriteLine(new string('=', 80));
            writer.WriteLine();

            for (int i = 0; i < dati.Count; i++)
            {
                if (i == 0 && hasHeaders)
                {
                    writer.WriteLine($"HEADER: {string.Join(" | ", dati[i])}");
                    writer.WriteLine(new string('-', 80));
                }
                else
                {
                    writer.WriteLine($"Riga {(hasHeaders ? i : i + 1),3}: {string.Join(" | ", dati[i])}");
                }
            }

            Console.WriteLine($"Dati esportati in formato TXT: {outputPath}");
            Console.WriteLine($"Dimensione file: {new FileInfo(outputPath).Length} bytes");
        }

        // Metodo sicuro per caricare XML evitando XXE
        static XmlDocument LoadXmlSafe(Stream stream)
        {
            var settings = new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,  // Niente DTD
                XmlResolver = null                       // Nessun resolver esterno
            };

            using var reader = XmlReader.Create(stream, settings);
            var doc = new XmlDocument
            {
                XmlResolver = null // Anche qui, niente resolver
            };
            doc.Load(reader);
            return doc;
        }

        static string GetParameterSyntax(string providerName, string paramName)
        {
            return providerName switch
            {
                "System.Data.SqlClient" => $"@{paramName}",
                "Microsoft.Data.SqlClient" => $"@{paramName}",
                "Oracle.ManagedDataAccess.Client" => $":{paramName}",
                _ => $"@{paramName}" // default fallback
            };
        }

        static void EsportaSql(List<List<string>> dati, string connectionString, string tableName, bool hasHeaders)
        {
            if (dati == null || dati.Count == 0)
            {
                Console.WriteLine("Nessun dato da inserire.");
                return;
            }

            // "Oracle.ManagedDataAccess.Client"
            string providerName = "System.Data.SqlClient";

            var headers = hasHeaders ? dati[0] : Enumerable.Range(0, dati[0].Count).Select(i => $"Column_{i + 1}").ToList();
            var rows = hasHeaders ? dati.Skip(1) : dati;

            DbProviderFactory factory = DbProviderFactories.GetFactory(providerName); 
            using DbConnection conn = factory.CreateConnection();
            conn.ConnectionString = connectionString;
            conn.Open();

            foreach (var row in rows)
            {
                var columnNames = string.Join(", ", headers);
                var paramNames = string.Join(", ", headers.Select((h, i) => GetParameterSyntax(providerName, $"p{i}")));

                var query = $"INSERT INTO {tableName} ({columnNames}) VALUES ({paramNames})";

                using DbCommand cmd = conn.CreateCommand();
                cmd.CommandText = query;

                for (int i = 0; i < headers.Count; i++)
                {
                    var param = cmd.CreateParameter();
                    param.ParameterName = $"p{i}";
                    param.Value = i < row.Count ? row[i] ?? (object)DBNull.Value : DBNull.Value;
                    cmd.Parameters.Add(param);
                }

                cmd.ExecuteNonQuery();
            }

//          Console.WriteLine($"Inseriti {rows.Count} record nella tabella '{tableName}'.");
        }
		
    }
}
