using ClosedXML.Excel;
using System.Text.Json;

namespace ExcelScript;
internal class Program
{
    static void Main(string[] args)
    {
        if (args.Length > 0)
        {
            if (args[0] == "--help" || args[0] == "-h")
            {
                Console.WriteLine("Usage: ExcelScript -i(--input) <inputPath> -o(--output) <outputPath> -c(--config) <configPath>");
                Console.WriteLine("       If you dont specify a path, it will automatically search the scripts directory");
                return;
            }
        }
        
        string? inputPath = null;
        string? outputPath = null;
        string? configPath = null;
        foreach (var arg in args)
        {
            if (arg == "-i" || arg == "--input")
            {
                if (inputPath != null)
                {
                    Console.WriteLine("Error: Input path already specified: " + inputPath);
                    return;
                }

                int index = Array.IndexOf(args, arg);
                if (index + 1 < args.Length)
                {
                    var path = args[index + 1];
                    if (File.Exists(path))
                        inputPath = path;
                    else
                    {
                        Console.WriteLine("Error: Input path does not exist: " + path);
                        return;
                    }
                }
                else
                {
                    Console.WriteLine("Error: Missing input path after " + arg);
                    return;
                }
            }
            else if (arg == "-o" || arg == "--output")
            {
                if (outputPath != null)
                {
                    Console.WriteLine("Error: Output path already specified: " + outputPath);
                    return;
                }

                int index = Array.IndexOf(args, arg);
                if (index + 1 < args.Length)
                {
                    outputPath = args[index + 1];
                }
                else
                {
                    Console.WriteLine("Error: Missing output path after " + arg);
                    return;
                }
            }
            else if (arg == "-c" || arg == "--config")
            {
                if (configPath != null)
                {
                    Console.WriteLine("Error: Config path already specified: " + configPath);
                    return;
                }

                int index = Array.IndexOf(args, arg);
                if (index + 1 < args.Length)
                {
                    var path = args[index + 1];
                    if (File.Exists(path))
                        configPath = path;
                    else
                    {
                        Console.WriteLine("Error: Config path does not exist: " + path);
                        return;
                    }
                }
                else
                {
                    Console.WriteLine("Error: Missing config path after " + arg);
                    return;
                }
            }
        }
        var scripts = Directory.GetFiles(".", "*.xlsx");
        var configs = Directory.GetFiles(".", "*.json");
        Config? config = null;
        if (inputPath == null)
        {
            if (scripts.Length == 0)
            {
                Console.WriteLine("Error: No input path specified and no scripts found in the current directory.");
                return;
            }
            else
            {
                Console.WriteLine("Warning: No input path specified and multiple scripts found in the current directory, choosing first!");
                inputPath = scripts[0];
            }
        }
        if (configPath == null)
        {
            if (configs.Length == 0)
            {
                Console.WriteLine("Warning: No config path specified and no configs found in the current directory.");
                Console.WriteLine("Warning: No config path specified and no configs found in the current directory.");
                config = new Config();
                return;
            }
            else
            {
                Console.WriteLine("Warning: No config path specified and multiple configs found in the current directory, choosing first!");
                configPath = configs[0];
            }
        }
        if (outputPath == null)
        {
            Console.WriteLine("Warning: No output path specified, using default output.xlsx");
            outputPath = "output.xlsx";
        }
        Console.WriteLine();

        if (config == null) Manipulator.LoadConfig(configPath, out config);

        TimeSpan totalTime = TimeSpan.Zero;
        var callback = new ProgressCallback()
        {
            WorkbookLoaded = (time) =>
            {
                Console.WriteLine($"Workbook loaded in {time.TotalMilliseconds} ms.");
                totalTime += time;
            },
            DataCopied = (time) =>
            {
                Console.WriteLine($"Data copied in {time.TotalMilliseconds} ms.");
                totalTime += time;
            },
            ItemsGrouped = (time) =>
            {
                Console.WriteLine($"Items grouped in {time.TotalMilliseconds} ms.");
                totalTime += time;
            },
            DataFormatted = (time) =>
            {
                Console.WriteLine($"Data formatted in {time.TotalMilliseconds} ms.");
                totalTime += time;
            },
            SummaryGenerated = (time) =>
            {
                Console.WriteLine($"Summary generated in {time.TotalMilliseconds} ms.");
                totalTime += time;
            },
            WorkbookStyled = (time) =>
            {
                Console.WriteLine($"Workbook styled in {time.TotalMilliseconds} ms.");
                totalTime += time;
            },
            CommentsCopied = (time) =>
            {
                Console.WriteLine($"Comments copied in {time.TotalMilliseconds} ms.");
                totalTime += time;
            },
            WorkbookSaved = (time) =>
            {
                Console.WriteLine($"Workbook saved in {time.TotalMilliseconds} ms.");
                totalTime += time;
                Console.WriteLine($"\nTotal processing time: {totalTime.TotalSeconds:F2} seconds.");
            }
        };
        var manipulator = new Manipulator(config);
        manipulator.Run(inputPath, outputPath, callback);

        Console.WriteLine("Press any key to quit...");
        Console.ReadKey();
    }
}