using Oas2xlsx.Console.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Oas2xlsx.Console.Settings
{
    public class OasSettings
    {
        public SourceType SourceType { get; set; }
        public string Source { get; set; }

        public string Target { get; set; }


        public OasSettings(string[] args)
        {
            SourceType = SourceType.FileSystem;

            if(args.Length <= 1)
            {
                throw new ArgumentOutOfRangeException("args", "Empty arguments");
            }
            for (int index = 0; index < args.Length - 1; index = index + 2)
            {
                var argName = args[index];
                var argValue = args[index + 1];

                switch (argName)
                {
                    case "-type":
                        SourceType = (SourceType)Enum.Parse(typeof(SourceType), argValue);
                        break;
                    case "-oas":
                        Source = argValue;
                        break;
                    case "-xlsx":
                        Target = argValue;
                        break;
                    default:
                        throw new ArgumentOutOfRangeException(argName, "Not supported argument");
                }
            }
            if(Source == null)
            {
                throw new ArgumentNullException("-oas");
            }
            if (Target == null)
            {
                throw new ArgumentNullException("-xlsx");
            }
        }

        public static void DisplayUsage()
        {
            StringBuilder builder = new StringBuilder();
            ColorConsole.WriteInfo("Tool usage:");
            ColorConsole.WriteInfo("oas2xslx.exe [-type <source type>] -oas <oas source file> -xlsx <xlsx target file>");

            ColorConsole.Write("-type ", ConsoleColor.Gray);
            ColorConsole.Write("<source type>", ConsoleColor.DarkGray);
            ColorConsole.Write(": ", ConsoleColor.Gray);
            ColorConsole.WriteInfo("Possible values are [filesystem, url]. Define the way to retrieve OAS definition. Optional, default value is FileSystem.");
            
            ColorConsole.Write("-oas ", ConsoleColor.Gray);
            ColorConsole.Write("<oas source file>", ConsoleColor.DarkGray);
            ColorConsole.Write(": ", ConsoleColor.Gray);
            ColorConsole.WriteInfo("path (file system or url) to access to the oas definition. Mandatory parameter.");

            ColorConsole.Write("-xlsx ", ConsoleColor.Gray);
            ColorConsole.Write("<xlsx target file>", ConsoleColor.DarkGray);
            ColorConsole.Write(": ", ConsoleColor.Gray);
            ColorConsole.WriteInfo("path to write the generated Excel file. Mandatory parameter.");           
        }
    }
}
