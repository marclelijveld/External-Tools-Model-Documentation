using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using Newtonsoft.Json;
using System.IO.Packaging;
using System.IO;
using Dax.Vpax.Tools;
using CommandLineParser.Arguments;
using CommandLineParser.Exceptions;
using System.ComponentModel;
using System.Reflection;

// TODO
// - Import from DMV 1100 (check for missing attributes?)
#pragma warning disable IDE0051 // Remove unused private members
namespace ModelDocumenter
{
    public enum ModelDocumenterErrEnum
    {
        [Description("The operation completed successfully.")]
        ERROR_SUCCESS,
        [Description("One or more arguments are not correct.")]
        ERROR_BAD_ARGUMENTS,
        [Description("This server is not supported.")]
        ERROR_NOT_SUPPORTED,
        [Description("Error exporting vpax file.")]
        ERROR_EXPORT_Vpax
    }

    class Program
    {
        public static Assembly execAssembly = Assembly.GetCallingAssembly();
        public static AssemblyName assemblyName = execAssembly.GetName();

        public static string applicationName = assemblyName.Name;
        public static string applicationVersion = $"{assemblyName.Version.Major:0}.{assemblyName.Version.Minor:0}";

        static void Main(string[] args)     
        {
            Console.WriteLine($"---");
            Console.WriteLine($"--- {applicationName} {applicationVersion} for.Net ({execAssembly.ImageRuntimeVersion})");
            Console.WriteLine($"---");

            var plugins = LoadPlugins();


            ModelDocumenterErrEnum exitCode = ModelDocumenterErrEnum.ERROR_SUCCESS;
            CommandLineParser.CommandLineParser parser = new CommandLineParser.CommandLineParser();

            ValueArgument<string> serverArgument = new ValueArgument<string>('s', "server", "This parameter specifies the server name.");
            serverArgument.Optional = false;

            ValueArgument<string> databaseArgument = new ValueArgument<string>('d', "database", "This parameter specifies the database name.");
            databaseArgument.Optional = false;

            ValueArgument<string> filenameArgument = new ValueArgument<string>('f', "filename", "This parameter specifies the export filename.");
            databaseArgument.Optional = false;

            ValueArgument<string> pbitemplateArgument = new ValueArgument<string>('p', "pbitemplate", "This parameter specifies the location of the power bi template file.");
            pbitemplateArgument.Optional = false;

            parser.Arguments.Add(serverArgument);
            parser.Arguments.Add(databaseArgument);
            parser.Arguments.Add(filenameArgument);
            parser.Arguments.Add(pbitemplateArgument);

            try
            {
                /* parse command line arguments */
                parser.ParseCommandLine(args);

                /* this prints results to the console */
                parser.ShowParsedArguments();
            }
            catch (CommandLineException e)
            {
                Console.WriteLine(e.Message);
                parser.ShowUsage();
                exitCode = ModelDocumenterErrEnum.ERROR_BAD_ARGUMENTS;
            }

            if (exitCode == ModelDocumenterErrEnum.ERROR_SUCCESS)
            {
                if (serverArgument.Value.StartsWith("pbiazure://api.powerbi.com"))
                {
                    Console.WriteLine($"Server {serverArgument.Value} not supported.");
                    exitCode = ModelDocumenterErrEnum.ERROR_NOT_SUPPORTED;
                }
                else
                {
                    exitCode = VpaxExport(serverArgument.Value, databaseArgument.Value, filenameArgument.Value);
                }
                if (exitCode == ModelDocumenterErrEnum.ERROR_SUCCESS)
                {
                    System.Diagnostics.Process.Start(pbitemplateArgument.Value);
                }                    
            }

            if (exitCode != ModelDocumenterErrEnum.ERROR_SUCCESS)
            {
                Console.WriteLine("Error 0x{0:x4}: {1}", (int)exitCode, exitCode.GetDescription());
                Console.ReadLine();
            }
            Environment.Exit((int)exitCode);
        }

        static IList<Assembly> LoadPlugins()
        {
            List<Assembly> pluginAssemblies = new List<Assembly>();

            foreach (var dll in Directory.EnumerateFiles(AppDomain.CurrentDomain.BaseDirectory, "*.dll"))
            {
                try
                {
                    var pluginAssembly = Assembly.LoadFile(dll);
                    if (pluginAssembly != null)
                    {
                        pluginAssemblies.Add(pluginAssembly);
                        Console.WriteLine("Successfully loaded plugin " + pluginAssembly.FullName + " from assembly " + Path.GetFileName(dll));
                    }
                }
                catch
                {

                }
            }

            return pluginAssemblies;
        }

        static ModelDocumenterErrEnum VpaxExport(string serverName, string databaseName, string fileName)
        {
            ModelDocumenterErrEnum exitCode = ModelDocumenterErrEnum.ERROR_SUCCESS;

            bool includeTomModel = true;

            //
            // Create log filename, use filename with .log extension
            //
            string logFile = Path.ChangeExtension(fileName, ".log");
            string outputPath = Path.GetDirectoryName(fileName);
            if (!Directory.Exists(outputPath))
            {
                try
                {
                    // Try to create the directory.
                    DirectoryInfo di = Directory.CreateDirectory(outputPath);
                }
                catch (Exception e)
                {
                    Console.WriteLine("Cannot create directory '{0}': {1}", outputPath, e.ToString());
                }
            }

            try
            {
                using (StreamWriter w = File.AppendText(logFile))
                {
                    Log("Exporting...", w);
                    //
                    // Get Dax.Model object from the SSAS engine
                    //
                    Log("[INFO] Get Dax.Model object from the SSAS engine", w);
                    Dax.Metadata.Model model = Dax.Metadata.Extractor.TomExtractor.GetDaxModel(serverName, databaseName, applicationName, applicationVersion);

                    //
                    // Get TOM model from the SSAS engine
                    //
                    Log("[INFO] Get TOM model from the SSAS engine", w);
                    Microsoft.AnalysisServices.Database database = includeTomModel ? Dax.Metadata.Extractor.TomExtractor.GetDatabase(serverName, databaseName) : null;

                    // 
                    // Create VertiPaq Analyzer views
                    //
                    Log("[INFO] Create VertiPaq Analyzer views", w);
                    Dax.ViewVpaExport.Model viewVpa = new Dax.ViewVpaExport.Model(model);

                    //
                    // Save VPAX file
                    // 
                    // TODO: export of database should be optional
                    Log("[INFO] Exporting VPAX file", w);
                    Dax.Vpax.Tools.VpaxTools.ExportVpax(fileName, model, viewVpa, database);

                    Log("Completed", w);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                exitCode = ModelDocumenterErrEnum.ERROR_EXPORT_Vpax;
            }

            return exitCode;
        }

        public static void Log(string logMessage, TextWriter w)
        {
            w.Write($"{DateTime.Now.ToLongTimeString()} {DateTime.Now.ToLongDateString()}");
            w.WriteLine($"  :{logMessage}");
        }
    }
}