﻿using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using Jitbit.Utils;
using System;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;

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
            dateTypes = "76,75,611,108,109,685,779,615".Split(',');
        }

        public void Export()
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(this.fileName, false))
            {
                var wbPart = document.WorkbookPart;

                var theSheet = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => String.IsNullOrEmpty(this.sheetName)
                || s.Name.ToString().Equals(this.sheetName, StringComparison.OrdinalIgnoreCase));
                var wsPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));
                var theRows = wsPart.Worksheet.Descendants<Row>();
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(this.csvFileName))
                {
                    foreach (var row in theRows.Skip(this.startFromLine - 1))
                    {
                        var cells = row.Descendants<Cell>();
                        var line = string.Join(",", cells.Select(c =>
                         CsvExport.MakeValueCsvFriendly(GetCellValue(c, wbPart).ToString().Trim())));
                        file.WriteLine(line);
                    }
                }
            }
        }

        // TODO - test this with a large number of samples
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
                const int EXCEL_TRUNCATE_DIGITS_COUNT = 15;

                if (decimal.TryParse(theCell.CellValue?.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out var dd))
                {
                    var stringValue = dd.ToString("F99").TrimEnd('0').TrimEnd('.');
                    string pattern = @"(-?0\.0*)([1-9]\d*)";
                    var matches = Regex.Matches(stringValue, pattern);

                    // THE FOLLOWING BLOCK INSURES THAT EXCEPT THE TRAILING + - SIGN AND PREFIX ZEROS THERE ARE NO MORE THAN 15 DIGITS.
                    // IF MORE THE LAST DIGITS EXCEEDING 15 ARE ROUNDED IN THE 15th VALUE
                    /*
                     *     Input 0.0000759497104417289 
                     *     Group 0.   0.0000759497104417289 
                           Group 1.	  0.0000
                           Group 2.	        759497104417289
                     */
                    if (matches.Count > 0 && matches[0].Success && matches[0].Groups.Count == 3 && matches[0].Groups[2].Value.Length > EXCEL_TRUNCATE_DIGITS_COUNT)
                    {
                        var zeroes = matches[0].Groups[1].Value;
                        var remaining = matches[0].Groups[2].Value;
                        var suffix = remaining.Substring(0, EXCEL_TRUNCATE_DIGITS_COUNT - 1);
                        var prefix = remaining.Substring(EXCEL_TRUNCATE_DIGITS_COUNT - 1, remaining.Length - EXCEL_TRUNCATE_DIGITS_COUNT + 1);
                        var compose = prefix[0] + "." + prefix.Substring(1, prefix.Length - 1);
                        var rounded = (int)Math.Round(double.Parse(compose, CultureInfo.InvariantCulture));
                        value = zeroes + suffix + rounded.ToString(CultureInfo.InvariantCulture);
                    }
                    else
                      value = stringValue;
                }
                else
                {
                    value = theCell.CellValue?.Text ?? "";
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
