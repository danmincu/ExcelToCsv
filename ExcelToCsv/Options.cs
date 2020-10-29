using CommandLine;
using System.Collections.Generic;

namespace ExcelToCsv
{

    using CommandLine.Text;

    class Options
    {
        [Option("filename", Required = true, HelpText = "Input filename.")]
        public string filename { get; set; }

        [Option("csvfilename", Required = true, HelpText = "Input filename.")]
        public string csvfilename { get; set; }


        [Option("fromline", Required = false, Default = 1, HelpText = "The first line where the extractraction should begin.")]
        public int fromline { get; set; }

        [Option("sheetname", Required = false, HelpText = "The name of the sheet to extract as CSV. First one if absent.")]
        public string sheetname { get; set; }


        [Usage(ApplicationAlias = "ExcelToCsv")]
        public static IEnumerable<Example> Examples
        {
            get
            {
                return new List<Example>() {

          new Example("Convert XLSX file to a CSV format", new Options { filename = "file.xlsx", csvfilename = "file.csv" })
         };
            }
        }
    }

}
