/* Davis Lee
 * 1/23/18
 * This program parses a spreadsheet file and checks for valid strings and numbers
 * Row counting starts at 1 with the first data row, data header is row 0
 */

using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System.Configuration;
using System.Collections.Specialized;
namespace DataParser1
{
    class MainClass
    {
        public static void Main(string[] args)
        {
            int cellErr = 0, cellTot = 0, rowErr = 0, rowTot = 0, strCnt = 0, numCnt = 0;
            string fileName = getData(ref strCnt, ref numCnt);
            Parser parser = new CsvParser();

            if (fileName.EndsWith(".csv", StringComparison.CurrentCulture))
            {
                parser = new CsvParser();
            }

            else if (fileName.EndsWith(".xlsx", StringComparison.CurrentCulture))
            {
                parser = new ExcelParser();
            }

            parser.Parse(fileName, ref cellErr, ref cellTot, ref rowErr, ref rowTot, strCnt, numCnt);

            Console.WriteLine("");
            Console.WriteLine(cellErr + " invalid cells found out of " + cellTot + " checked");
            Console.WriteLine(rowErr + " invalid rows found out of " + rowTot + " checked\n");
            Console.WriteLine("Parse complete");
        }
        public static string getData(ref int strCnt, ref int numCnt)
        {
            Console.WriteLine("Welcome to the data table parser");
            Console.Write("\nWhat is the name of the file? ");
            string fileName = Console.ReadLine();
            while (!File.Exists(fileName))
            {
                Console.WriteLine("Invalid input, please enter valid file name");
                Console.Write("What is the name of the file? ");
                fileName = Console.ReadLine();
            }
            string input = "";
            Console.Write("\nHow many data columns are strings? ");
            input = Console.ReadLine();

            int n;
            while (!int.TryParse(input, out n) || Convert.ToInt32(input) < 0)
            {
                Console.WriteLine("Invalid input, please enter valid integer");
                Console.Write("How many data columns are strings? ");
                input = Console.ReadLine();
            }

            strCnt = Convert.ToInt32(input);
            Console.Write("\nHow many data columns are numbers? ");
            input = Console.ReadLine();

            while (!int.TryParse(input, out n) || Convert.ToInt32(input) < 0)
            {
                Console.WriteLine("Invalid input, please enter valid integer");
                Console.Write("How many data columns are numbers? ");
                input = Console.ReadLine();
            }
            numCnt = Convert.ToInt32(input);

            return fileName;
        }
    }

    interface Parser
    {
        void Parse(string fileName, ref int cellErr, ref int cellTot,
                   ref int rowErr, ref int rowTot, int strCnt, int numCnt);
    }

    class CsvParser : Parser
    {
        public void Parse(string fileName, ref int cellErr, ref int cellTot,
                   ref int rowErr, ref int rowTot, int strCnt, int numCnt)
        {
            var fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            using (var streamReader = new StreamReader(fileStream, System.Text.Encoding.UTF8))
            {
                string line;
                string[] columnHeaders = ColumnHeaders(streamReader, strCnt, numCnt);
                while ((line = streamReader.ReadLine()) != null)
                {
                    ParseRow(line, ref columnHeaders, ref rowTot, ref rowErr,
                             ref cellTot, ref cellErr, strCnt, numCnt);
                }
            }
        }
        //Returns string array of data column headers
        public string[] ColumnHeaders(StreamReader streamReader, int strCnt, int numCnt)
        {
            string line;

            //check for empty document
            if ((line = streamReader.ReadLine()) == null)
            {
                Console.WriteLine("Error, empty document file");
                Environment.Exit(-1);
            }

            string[] columnHeaders = line.Split(',');

            //check for valid number of headers
            if (columnHeaders.Length != strCnt + numCnt)
            {
                Console.WriteLine("Error, invalid number of headers");
                Environment.Exit(-1);
            }
            return columnHeaders;
        }

        //Parses each row of data and writes to console if errors are found
        public void ParseRow(string line, ref string[] columnHeaders, ref int rowTot,
                             ref int rowErr, ref int cellTot, ref int cellErr,
                             int strCnt, int numCnt)
        {
            rowTot++;
            Console.WriteLine();
            bool err = false;
            string[] columns = line.Split(',');

            //if there are missing data entries no individual cell data is checked
            if (columns.Length < strCnt + numCnt)
            {
                Console.WriteLine("Row entry " + rowTot +
                              " is corrupted, missing data fields");
                err = true;
            }
            else
            {
                //if a row has extraneous data entries then the first entries are still checked
                if (columns.Length > strCnt + numCnt)
                {
                    Console.WriteLine("Row entry " + rowTot +
                                  " contains extraneous data fields");
                    err = true;
                }

                for (int i = 0; i < strCnt; i++)
                {
                    cellTot++;
                    for (int j = 0; j < columns[i].Length; j++)
                    {
                        //invalid strings are classified as having a digit
                        if (Char.IsDigit(columns[i][j]))
                        {
                            Console.WriteLine("Row entry " + rowTot +
                                 " contains invalid " + columnHeaders[i]);
                            Console.WriteLine("\"" + columns[i] + "\" is not a valid string");
                            err = true;
                            cellErr++;
                            break;
                        }
                    }
                }

                for (int i = 4; i <= 9; i++)
                {
                    cellTot++;
                    double n;

                    if (!double.TryParse(columns[i], out n))
                    {
                        Console.WriteLine("Row entry " + rowTot +
                            " contains invalid " + columnHeaders[i]);
                        Console.WriteLine("\"" + columns[i] + "\" is not a valid number");
                        err = true;
                        cellErr++;
                    }
                }
            }

            if (err)
            {
                rowErr++;
            }
            else
            {
                Console.WriteLine("Row entry " + rowTot + " confirmed as valid");
            }
        }
    }

    class ExcelParser : Parser
    {
        public void Parse(string fileName, ref int cellErr, ref int cellTot,
                   ref int rowErr, ref int rowTot, int strCnt, int numCnt)
        {
            FileInfo fi = new FileInfo(fileName);
            using (ExcelPackage xlPackage = new ExcelPackage(fi))
            {
                //Get the worksheet in the workbook 
                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets[1];
                string[] columnHeaders = ColumnHeaders(worksheet, strCnt, numCnt);
                for (int i = worksheet.Dimension.Start.Row + 1; i <= worksheet.Dimension.End.Row; i++)
                {
                    ParseRow(worksheet, i, ref columnHeaders,
                             ref rowErr, ref cellErr, ref rowTot, ref cellTot, strCnt, numCnt);
                }
            }
        }

        //Returns string array of data column headers
        public string[] ColumnHeaders(ExcelWorksheet workSheet, int strCnt, int numCnt)
        {

            string[] columnHeaders = new string[strCnt + numCnt];

            //check for valid number of headers
            if (workSheet.Dimension.End.Column - workSheet.Dimension.Start.Column + 1 !=
                strCnt + numCnt)
            {
                Console.WriteLine("Error, invalid number of headers");
                Environment.Exit(-1);
            }
            for (int i = workSheet.Dimension.Start.Column; i <= workSheet.Dimension.Start.Column; i++)
            {
                columnHeaders[i - 1] = (string)workSheet.Cells[workSheet.Dimension.Start.Row, i].Value;
            }
            return columnHeaders;
        }

        //Parses each row of data and writes to console if errors are found
        public void ParseRow(ExcelWorksheet workSheet, int row, ref string[] columnHeaders,
                             ref int rowErr, ref int cellErr, ref int rowTot,
                             ref int cellTot, int strCnt, int numCnt)
        {
            rowTot++;
            Console.WriteLine();
            bool err = false;

            for (int i = workSheet.Dimension.Start.Column; i < workSheet.Dimension.Start.Column + strCnt; i++)
            {
                string cell = workSheet.Cells[row, i].Value.ToString();
                cellTot++;
                for (int j = 0; j < cell.Length; j++)
                {
                    //invalid strings are classified as having a digit
                    if (Char.IsDigit(cell[j]))
                    {
                        Console.WriteLine("Row entry " + rowTot +
                             " contains invalid " + columnHeaders[i - 1]);
                        Console.WriteLine("\"" + cell + "\" is not a valid string");
                        err = true;
                        cellErr++;
                        break;
                    }
                }
            }

            for (int i = workSheet.Dimension.Start.Column + strCnt; i < workSheet.Dimension.Start.Column + strCnt + numCnt; i++)
            {
                string cell = workSheet.Cells[row, i].Value.ToString();
                cellTot++;
                double n;

                if (!double.TryParse(cell, out n))
                {
                    Console.WriteLine("Row entry " + rowTot +
                        " contains invalid " + columnHeaders[i - 1]);
                    Console.WriteLine("\"" + cell + "\" is not a valid number");
                    err = true;
                    cellErr++;
                }
            }

            if (err)
            {
                rowErr++;
            }
            else
            {
                Console.WriteLine("Row entry " + rowTot + " confirmed as valid");
            }
        }
    }
}