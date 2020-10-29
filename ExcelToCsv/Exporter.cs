using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Configuration;
using System.Linq;

namespace ExcelToCsv
{
    public class Exporter
    {
        private static string[] dateTypes;
        private string fileName;
        private string csvFileName;
        private int startFromLine;
        private string sheetName;
        public Exporter(string fileName, string csvFileName, int startFromLine, string sheetName)
        {
            this.fileName = fileName;
            this.startFromLine = startFromLine;
            this.sheetName = sheetName;
            this.csvFileName = csvFileName;
            dateTypes = ConfigurationManager.AppSettings["dateTypes"].Split(',');
        }

        public void Export()
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(this.fileName, false))
            {
                var wbPart = document.WorkbookPart;

                var theSheet = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name.ToString().Equals(this.sheetName, StringComparison.OrdinalIgnoreCase));
                var wsPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));
                var theRows = wsPart.Worksheet.Descendants<Row>();
                var meaninfulRows = theRows.Skip(this.startFromLine - 1).ToList();
                foreach (var row in meaninfulRows)
                {
                    var cells = row.Descendants<Cell>().ToList();
                    Console.WriteLine("");
                    foreach (var item in cells)
                    {
                        var text = GetCellValue(item, wbPart).ToString().Trim().Replace("\n", " ");

                        Console.Write(text + ",");
                    }
                    Console.WriteLine("");
                }
            }
        }



        private static object GetCellValue(Cell theCell, WorkbookPart wbPart)
        {
            var attrib = theCell.GetAttributes();
            Object value = theCell.InnerText;

            // If the cell represents an integer number, you are done. 
            // For dates, this code returns the serialized value that 
            // represents the date. The code handles strings and 
            // Booleans individually. For shared strings, the code 
            // looks up the corresponding value in the shared string 
            // table. For Booleans, the code converts the value into 
            // the words TRUE or FALSE.
            if (theCell.DataType != null)
            {
                switch (theCell.DataType.Value)
                {
                    case CellValues.Date:

                        TimeSpan datefromexcel = new TimeSpan(int.Parse(value.ToString()), 0, 0, 0);
                        value = (new DateTime(1899, 12, 30).Add(datefromexcel)).ToString();
                        break;
                    case CellValues.SharedString:

                        // For shared strings, look up the value in the
                        // shared strings table.
                        var stringTable =
                            wbPart.GetPartsOfType<SharedStringTablePart>()
                            .FirstOrDefault();

                        // If the shared string table is missing, something 
                        // is wrong. Return the index that is in
                        // the cell. Otherwise, look up the correct text in 
                        // the table.
                        if (stringTable != null)
                        {
                            var v = stringTable.SharedStringTable
                                .ElementAt(int.Parse(value.ToString()));

                            value = v.InnerText;
                        }
                        break;

                    case CellValues.Boolean:
                        switch (value.ToString())
                        {
                            case "0":
                                value = "FALSE";
                                break;
                            default:
                                value = "TRUE";
                                break;
                        }
                        break;
                }
            }
            else
            {
                if (decimal.TryParse(theCell.CellValue.Text, out var dd))
                {
                    value = dd.ToString("F99").TrimEnd('0');
                }
                else
                {
                    if (theCell.CellValue.Text.ToUpper().Contains("E-"))
                    {
                        try
                        {
                            // 5.3651762733430766E-2
                            value = decimal.Parse(theCell.CellValue.Text, System.Globalization.NumberStyles.Float).ToString("F99").TrimEnd('0');
                        }
                        catch
                        {
                            value = theCell.CellValue.Text;
                        }
                    }

                }
                /*  THIS SECTION CAN BE USED IF WE WANT TO PRESERVE NUMBER FORMATS
                var cellFormats = wbPart.WorkbookStylesPart.Stylesheet.CellFormats;
                var numberingFormats = wbPart.WorkbookStylesPart.Stylesheet.NumberingFormats;


                 bool isDate = false;
                 var styleIndex = (int)theCell.StyleIndex.Value;
                 var cellFormatt = (CellFormat)cellFormats.ElementAt(styleIndex);

                 if (cellFormatt.NumberFormatId != null)
                 {
                     var numberFormatId = cellFormatt.NumberFormatId.Value;
                     var numberingFormat = numberingFormats.Cast<NumberingFormat>()
                         .SingleOrDefault(f => f.NumberFormatId.Value == numberFormatId);

                     // Here's yer string! Example: $#,##0.00_);[Red]($#,##0.00)
                     if (numberingFormat != null && numberingFormat.FormatCode.Value.Contains("mmm"))
                     {
                         string formatString = numberingFormat.FormatCode.Value;
                         isDate = true;
                     }
                 }


                 int dateInteger = 0;
                 if (!string.IsNullOrEmpty(value.ToString()) && int.TryParse(value.ToString(), out dateInteger) && attrib.Count >= 1 && (dateTypes.Contains(attrib[1].Value) || isDate)) 
                 {
                     value = (new DateTime(1899, 12, 30).Add(new TimeSpan(dateInteger, 0, 0, 0)));
                 }
                 */
            }
            return value;

        }

    }
}
