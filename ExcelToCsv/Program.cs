using CommandLine;
using System;

namespace ExcelToCsv
{
    class Program
    {
        static void Main(string[] args)
        {
            var parsedArgs = Parser.Default.ParseArguments<Options>(args);

            if (parsedArgs.Tag == ParserResultType.NotParsed)
            {
                parsedArgs.WithNotParsed(o =>
                {
                    foreach (var item in o)
                    {
                        Console.WriteLine(item);
                    }
                });
            }
            else
            {
                parsedArgs.WithParsed(o => {

                    if (!System.IO.File.Exists(o.filename))
                    {
                        Console.WriteLine($"{o.filename} does not exists!");
                        return;
                    }
                    var exporter = new Exporter(o.filename, o.csvfilename, o.fromline, o.sheetname);
                    exporter.Export();
                });
            }
            Console.ReadLine();
        }
    }
}
