using DymoDemo.Core;
using System.IO;

namespace DymoDemo.Cli;

internal class Program
{
    static async Task<int> Main(string[] args)
    {
        if (args.Length == 0 || args.Any(a => a.Equals("/help", StringComparison.OrdinalIgnoreCase)
                                           || a.Equals("/?", StringComparison.OrdinalIgnoreCase)))
        {
            PrintUsage();
            return 0;
        }

        string? printerSearch = null;
        string? labelFile = null;
        int copies = 1;
        int? roll = null;
        string? rollType = null;
        var labelValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (var arg in args)
        {
            if (TryParseArg(arg, "PRINTER", out var printerVal))
            {
                printerSearch = printerVal;
            }
            else if (TryParseArg(arg, "LABEL", out var labelVal))
            {
                labelFile = labelVal;
            }
            else if (TryParseArg(arg, "COPIES", out var copiesVal))
            {
                if (!int.TryParse(copiesVal, out copies) || copies < 1)
                {
                    Console.Error.WriteLine($"Error: Invalid copies value '{copiesVal}'.");
                    return 1;
                }
            }
            else if (TryParseArg(arg, "ROLLTYPE", out var rollTypeVal))
            {
                rollType = rollTypeVal;
            }
            else if (TryParseArg(arg, "ROLL", out var rollVal))
            {
                roll = rollVal.ToUpperInvariant() switch
                {
                    "AUTO" => 0,
                    "LEFT" => 1,
                    "RIGHT" => 2,
                    _ => null
                };
                if (roll == null)
                {
                    Console.Error.WriteLine($"Error: Invalid roll value '{rollVal}'. Use Auto, Left, or Right.");
                    return 1;
                }
            }
            else if (TryParseSetArg(arg, out var objName, out var objValue))
            {
                labelValues[objName] = objValue;
            }
            else
            {
                Console.Error.WriteLine($"Error: Unknown argument '{arg}'.");
                PrintUsage();
                return 1;
            }
        }

        // Validate required parameters
        if (string.IsNullOrEmpty(printerSearch))
        {
            Console.Error.WriteLine("Error: /PRINTER is required.");
            return 1;
        }

        // Run
        try
        {
            if (!DymoService.IsDymoConnectInstalled())
            {
                Console.Error.WriteLine("Error: DYMO Connect software is not installed.");
                Console.Error.WriteLine("Download it from https://www.dymo.com/support/downloads");
                return 1;
            }

            var service = new DymoService();

            // Find printer (wait for network printer discovery)
            Console.WriteLine("Discovering printers...");
            var printers = await service.GetPrintersAsync();
            if (printers.Count == 0)
            {
                Console.Error.WriteLine("Error: No Dymo printers found. Verify the printer is powered on and DYMO Connect Service is running.");
                return 1;
            }

            var printer = printers.FirstOrDefault(p =>
                p.Name.Contains(printerSearch, StringComparison.OrdinalIgnoreCase));

            if (printer == null)
            {
                Console.Error.WriteLine($"Error: No printer found matching '{printerSearch}'.");
                Console.Error.WriteLine("Available printers:");
                foreach (var p in printers)
                    Console.Error.WriteLine($"  - {p.Name}");
                return 1;
            }

            Console.WriteLine($"Using printer: {printer.Name}"); 

            bool hasRollStatus = service.IsRollStatusSupported(printer.Name);

            if (hasRollStatus && (rollType != null || !string.IsNullOrEmpty(labelFile)))
            {
                await Task.Delay(1000); // Small delay to ensure printer status is up to date
                var info = await service.GetConsumableInfoAsync(printer.Name);

                if (info != null && info.LabelsRemaining == "0")
                {
                    Console.Error.WriteLine("Error: No labels remaining. Please replace the label roll.");
                    return 1;
                }

                if (rollType != null)
                {
                    if (info == null)
                    {
                        Console.Error.WriteLine("Error: Unable to retrieve consumable information to verify roll type.");
                        return 1;
                    }

                    if (!info.Name.Contains(rollType, StringComparison.OrdinalIgnoreCase))
                    {
                        Console.Error.WriteLine($"Error: Expected roll type '{rollType}' but found '{info.Name}'.");
                        return 1;
                    }

                    Console.WriteLine($"Roll type verified: {info.Name}");
                }
            }
            else if (rollType != null)
            {
                Console.Error.WriteLine("Warning: Printer does not support roll status queries. Cannot verify roll type.");
            }

            // If no label file was supplied, just confirm the printer is available
            if (string.IsNullOrEmpty(labelFile))
            {
                Console.WriteLine("Printer is available.");
                return 0;
            }

            if (!File.Exists(labelFile))
            {
                Console.Error.WriteLine($"Error: Label file not found: {labelFile}");
                return 1;
            }

            // Load label
            service.LoadLabel(labelFile);
            Console.WriteLine($"Loaded label: {labelFile}");

            // Set label object values
            if (labelValues.Count > 0)
            {
                var labelObjects = service.GetLabelObjects();
                foreach (var (name, value) in labelValues)
                {
                    var obj = labelObjects.FirstOrDefault(o =>
                        o.Name.Equals(name, StringComparison.OrdinalIgnoreCase));

                    if (obj == null)
                    {
                        Console.Error.WriteLine($"Warning: Label object '{name}' not found. Available objects:");
                        foreach (var o in labelObjects)
                            Console.Error.WriteLine($"  - {o.Name}");
                        return 1;
                    }

                    service.UpdateLabelObject(obj, value);
                    Console.WriteLine($"Set '{obj.Name}' = '{value}'");
                }
            }

            // Print
            service.PrintLabel(printer.Name, copies, roll);
            Console.WriteLine($"Printed {copies} cop{(copies == 1 ? "y" : "ies")}.");

            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
            return 1;
        }
    }

    /// <summary>
    /// Tries to parse /KEY=value style arguments.
    /// </summary>
    static bool TryParseArg(string arg, string key, out string value)
    {
        var prefix = $"/{key}=";
        if (arg.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
        {
            value = arg[prefix.Length..];
            return true;
        }
        value = string.Empty;
        return false;
    }

    /// <summary>
    /// Tries to parse /SET:ObjectName=Value style arguments.
    /// </summary>
    static bool TryParseSetArg(string arg, out string objectName, out string objectValue)
    {
        const string prefix = "/SET:";
        if (arg.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
        {
            var rest = arg[prefix.Length..];
            var eqIndex = rest.IndexOf('=');
            if (eqIndex > 0)
            {
                objectName = rest[..eqIndex];
                objectValue = rest[(eqIndex + 1)..];
                return true;
            }
        }
        objectName = string.Empty;
        objectValue = string.Empty;
        return false;
    }

    static void PrintUsage()
    {
        Console.WriteLine("""
            DymoDemo.Cli - Command-line Dymo label printer

            Usage:
              DymoDemo.Cli /PRINTER=<search> [/LABEL=<file>] [/SET:<name>=<value> ...] [/COPIES=<n>] [/ROLL=<Auto|Left|Right>] [/ROLLTYPE=<name>]

            Parameters:
              /PRINTER=<search>       Required. Matches the first printer whose name contains <search>.
                                      Example: /PRINTER=500
              /LABEL=<file>           Path to the .label or .dymo file. Required for printing.
                                      If omitted, only checks whether the printer is available.
              /SET:<name>=<value>     Sets a label object value by its name. Repeat for multiple objects.
                                      Example: /SET:ProductName=Widget /SET:Price=9.99
              /COPIES=<n>             Number of copies to print (default: 1).
              /ROLL=<Auto|Left|Right> Roll selection for Twin Turbo 450 printers (default: not set).
              /ROLLTYPE=<name>        Verifies the loaded roll type contains <name> before printing.
                                      Requires a printer that supports roll status (e.g., LabelWriter 550).
                                      Example: /ROLLTYPE=30336

            Examples:
              DymoDemo.Cli /PRINTER=550
              DymoDemo.Cli /PRINTER=500 /LABEL=shipping.dymo /SET:Name="John Doe" /SET:Address="123 Main St"
              DymoDemo.Cli /PRINTER="LabelWriter" /LABEL=price.label /SET:Price=4.99 /COPIES=10
              DymoDemo.Cli /PRINTER=550 /LABEL=address.dymo /ROLLTYPE=30336
            """);
    }
}
