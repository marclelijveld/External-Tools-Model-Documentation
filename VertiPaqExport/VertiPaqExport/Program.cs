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

// TODO
// - Import from DMV 1100 (check for missing attributes?)
#pragma warning disable IDE0051 // Remove unused private members
namespace VertiPaqExport
{
    class Program
    {


        static void Main(string[] args)
        {
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
                Console.ReadLine();
            }

            VpaqExport(serverArgument.Value, databaseArgument.Value, filenameArgument.Value);
        }

        static void VpaqExport(string serverName, string databaseName, string fileName)
        {
            const string applicationName = "VpaqExport";
            const string applicationVersion = "0.1";
            bool includeTomModel = true;

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
                Console.ReadLine();
            }
        }

        public static void Log(string logMessage, TextWriter w)
        {
            w.Write($"{DateTime.Now.ToLongTimeString()} {DateTime.Now.ToLongDateString()}");
            w.WriteLine($"  :{logMessage}");
        }
    }
}
