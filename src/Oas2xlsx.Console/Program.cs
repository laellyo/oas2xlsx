using ClosedXML.Excel;
using Microsoft.OpenApi.Models;
using Microsoft.OpenApi.Readers;
using System;
using System.Collections.Generic;
using System.IO;
using Oas2xlsx.Console.Settings;
using Oas2xlsx.Console.Helpers;

namespace Oas2xlsx.Console
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                OasSettings.DisplayUsage();
                return;
            }
            OasSettings oasSettings;
            try
            {
                oasSettings = new OasSettings(args);
            }
            catch (Exception)
            {
                ColorConsole.WriteError("Impossible to parse your command. Please check it respects tool parameters describe below.");
                OasSettings.DisplayUsage();
                return;
            }

            try
            {
                using (Stream oasFile = File.OpenRead(oasSettings.Source))
                {

                    var settings = new OpenApiReaderSettings()
                    {
                        ReferenceResolution = ReferenceResolutionSetting.ResolveLocalReferences
                    };
                    var reader = new OpenApiStreamReader(settings);

                    var diagnostic = new OpenApiDiagnostic();
                    OpenApiDocument document = reader.Read(oasFile, out diagnostic);
                    if (diagnostic.Errors.Count > 0)
                    {
                        ColorConsole.WriteError("Unable to generate an Excel file, the OAS parser are detected some issues:");
                        foreach (var error in diagnostic.Errors)
                        {
                            ColorConsole.WriteError(string.Format("[{0}] : {1}", error.Pointer, error.Message));
                        }
                        return;
                    }

                    var generator = new ExcelGenerator(document);
                    generator.Generate(oasSettings.Target);
                }
            }
            catch (FileNotFoundException e)
            {
                ColorConsole.WriteError(e.Message);
                return;
            }

            ColorConsole.WriteSuccess("Excel document has been successfully created!");

        }

    }
}
