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
namespace VertiPaqExport
{
    public enum VpaqExportErrEnum
    {
        [Description("The operation completed successfully.")]
        ERROR_SUCCESS,
        [Description("One or more arguments are not correct.")]
        ERROR_BAD_ARGUMENTS,
        [Description("This server is not supported.")]
        ERROR_NOT_SUPPORTED,
        [Description("Error exporting vpaq file.")]
        ERROR_EXPORT_VPAQ
    }

    class Program
    {
        public static string applicationName = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
        public static string applicationVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();

        static void Main(string[] args)
        {
            VpaqExportErrEnum exitCode = VpaqExportErrEnum.ERROR_SUCCESS;
            CommandLineParser.CommandLineParser parser = new CommandLineParser.CommandLineParser();

            ValueArgument<string> serverArgument = new ValueArgument<string>('s', "server", "This parameter specifies the server name.");
            serverArgument.Optional = false;

            ValueArgument<string> databaseArgument = new ValueArgument<string>('d', "database", "This parameter specifies the database name.");
            databaseArgument.Optional = false;

            ValueArgument<string> filenameArgument = new ValueArgument<string>('f', "filename", "This parameter specifies the export filename.");
            databaseArgument.Optional = false;

            parser.Arguments.Add(serverArgument);
            parser.Arguments.Add(databaseArgument);
            parser.Arguments.Add(filenameArgument);

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
                exitCode = VpaqExportErrEnum.ERROR_BAD_ARGUMENTS;
            }

            if (exitCode == VpaqExportErrEnum.ERROR_SUCCESS)
            {
                if (serverArgument.Value.StartsWith("pbiazure://api.powerbi.com"))
                {
                    Console.WriteLine($"Server {serverArgument.Value} not supported.");
                    exitCode = VpaqExportErrEnum.ERROR_NOT_SUPPORTED;
                }
                else
                {
                    exitCode = VpaqExport(serverArgument.Value, databaseArgument.Value, filenameArgument.Value);
                }
            }

            if (exitCode != VpaqExportErrEnum.ERROR_SUCCESS)
            {
                Console.WriteLine("Error 0x{0:x4}: {1}", (int)exitCode, exitCode.GetDescription());
                Console.ReadLine();
            }
            Environment.Exit((int)exitCode);
        }

        static VpaqExportErrEnum VpaqExport(string serverName, string databaseName, string fileName)
        {
            VpaqExportErrEnum exitCode = VpaqExportErrEnum.ERROR_SUCCESS;

            bool includeTomModel = true;

            //
            // Create log filename, use filename with .log extension
            //
            string logFile = Path.ChangeExtension(fileName, ".log");

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
                exitCode = VpaqExportErrEnum.ERROR_EXPORT_VPAQ;
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