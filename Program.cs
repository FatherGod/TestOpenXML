using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;

namespace TestOpenXML
{
    class Program
    {

        static void Main()
        {
            Excel.Application testChart = new Excel.Application();
            //Excel.Application xlApp = new Excel.Application();
            //xlApp.Visible = false;
            //Excel.Workbook wb = xlApp.Workbooks.Open(@"C:\Users\edoua\Desktop\a detruire.xlsx");
            Excel.Workbook openXmlTest = testChart.Workbooks.Open(@"C:\Users\edoua\Desktop\Résultats Test.xlsx");
            Excel.Worksheet openXmlChartSheet = openXmlTest.Worksheets[2];
            //Excel.Worksheet dest = wb.Worksheets.Add();
            //dest.Name = "Destination";

            //SpreadsheetDocument document = SpreadsheetDocument.Open(@"C:\Users\edoua\Desktop\a detruire out.xlsx", true);
            //WorkbookPart wbPart = document.WorkbookPart;
            //Sheet sheet = wbPart.Workbook.Descendants<Sheet>().Where(
            //    (s) => s.Name == "source").FirstOrDefault();
            //Worksheet ws = ((WorksheetPart)(wbPart.GetPartById(sheet.Id))).Worksheet;

            //List<List<String>> data = readRangeFromExcel(wb.Worksheets["source"], 1, 1, 7, 7);

            //List<List<String>> data;
            DateTime start;
            TimeSpan dur;
            String cell;

            //for (int i = 1; i <= 400; i++)
            //{
            //    data = GenerateData(i);

            //    start = DateTime.Now;

            //    writeOnSheet(wb.Worksheets["Destination"], data);

            //    dur = DateTime.Now - start;

            //    openXmlChartSheet.Cells[1, i].Value = i;
            //    cell = dur.ToString().Split(':')[2];
            //    openXmlChartSheet.Cells[2, i].Value = cell;
            //}

            //wb.Save();
            //wb.Close();
            //xlApp.Quit();

            //String output = "";
            //foreach(List<String> ls in data)
            //{
            //    foreach(String s in ls)
            //    {
            //        output += s + "\t";
            //    }
            //    output += "\n";
            //}

            //Console.WriteLine(output);
            //Console.ReadLine();

            System.Threading.Thread.Sleep(100);

            SpreadsheetDocument document = SpreadsheetDocument.Open(@"C:\Users\edoua\Desktop\a detruire out.xlsx", true);
            WorkbookPart wbPart = document.WorkbookPart;
            Sheet sheet = wbPart.Workbook.Descendants<Sheet>().Where(
                (s) => s.Name == "source").FirstOrDefault();
            Worksheet ws = ((WorksheetPart)(wbPart.GetPartById(sheet.Id))).Worksheet;

            List<List<UInt32Value>> style; //= readStyleFromExcel(ws, 1, 1, 7, 7);

            Sheet destSheet = wbPart.Workbook.Descendants<Sheet>().Where(
                (s) => s.Name == "Destination").FirstOrDefault();
            Worksheet dest = ((WorksheetPart)(wbPart.GetPartById(sheet.Id))).Worksheet;

            //applyStyleOnSheet(ws, style);
            //output = "";
            //foreach (List<UInt32Value> ls in style)
            //{
            //    foreach (UInt32Value s in ls)
            //    {
            //        output += s.ToString() + "\t";
            //    }
            //    output += "\n";
            //}

            for (long i = 1; i <= 400; i++)
            {
                style = readStyleFromExcel(ws,1, 1, i, i);

                start = DateTime.Now;

                applyStyleOnSheet(dest, style);

                dur = DateTime.Now - start;

                openXmlChartSheet.Cells[1, i].Value = i;
                cell = dur.ToString().Split(':')[2];
                openXmlChartSheet.Cells[2, i].Value = cell;
            }

            //Console.WriteLine(output);

            document.Save();
            document.Close();

            openXmlChartSheet.SaveAs(@"C:\Users\edoua\Desktop\Résultats Test.xlsx");
            openXmlTest.Close();
            testChart.Quit();
        }

        private static List<List<string>> GenerateData(int v)
        {
            Random r = new Random();
            List<List<String>> output = new List<List<string>>();
            for(int i = 0; i < v; i++)
            {
                output.Add(new List<string>());
                for(int j = 0; j < v; j++)
                {
                    output.ElementAt(i).Add(r.Next().ToString());
                }
            }
            return output;
        }

        public static void writeOnSheet(Excel.Worksheet ws, List<List<String>> data)
        {
            //long rowIndex = 1, colIndex = 1;
            //foreach(List<String> row in data)
            //{
            //    foreach(String col in row)
            //    {
            //        ws.Cells[rowIndex, colIndex].Value = col;
            //        colIndex++;
            //    }
            //    rowIndex++;
            //    colIndex = 1;
            //}
            Excel.Range r = ws.Range[ws.Cells[1, 1], ws.Cells[data.Count, data.ElementAt(0).Count]];

            //String[,] input = temp.ToArray();
            r.Value = ArrayToList(data);
        }

        public static void writeOnSheet(Worksheet ws, List<List<String>> style)
        {
            Cell c;
            long rowIndex = 1, colIndex = 1;
            foreach (List<String> row in style)
            {
                foreach (String col in row)
                {
                    c = InsertCellInWorksheet(ws, AddressToString(rowIndex, colIndex));
                    c.CellValue = new CellValue(col);
                    colIndex++;
                }
                rowIndex++;
                colIndex = 1;
            }
        }

        public static void applyStyleOnSheet(Worksheet ws, List<List<UInt32Value>> style)
        {
            Cell c;
            long rowIndex = 1, colIndex = 1;
            foreach (List<UInt32Value> row in style)
            {
                foreach(UInt32Value col in row)
                {
                    c = InsertCellInWorksheet(ws, AddressToString(rowIndex, colIndex));
                    c.StyleIndex = col;
                    colIndex++;
                }
                rowIndex++;
                colIndex = 1;
            }
        }

        public static List<List<String>> readRangeFromExcel(Excel.Worksheet sheet, long startRow, long startColumn, long endRow, long endColumn)
        {
            object[,] o = (object[,])sheet.Range[sheet.Cells[startRow, startColumn], sheet.Cells[endRow, endColumn]].Value;
            long rows = o.GetLength(0);
            long columns = o.GetLength(1);
            List<List<String>> output = new List<List<string>>();
            for (int row = 0; row < rows; row++)
            {
                output.Add(new List<string>());
                for (int column = 0; column < columns; column++)
                    output.ElementAt(row).Add((o[row + 1, column + 1] != null) ? o[row + 1, column + 1].ToString() : "");
            }
            return output;

        }

        public static List<List<String>> readRangeFromExcel(Worksheet sheet, long startRow, long startColumn, long endRow, long endColumn)
        {
            Cell c;
            List<List<String>> output = new List<List<String>>();
            for (long row = 0; row <= endRow - startRow; row++)
            {
                output.Add(new List<String>());
                for (long column = 0; column <= endColumn - startColumn; column++)
                {
                    c = InsertCellInWorksheet(sheet, AddressToString(row + startRow, column + startColumn));
                    output.ElementAt((int)row).Add(c.InnerText);
                }
            }
            return output;

        }

        public static List<List<UInt32Value>> readStyleFromExcel(Worksheet sheet, long startRow, long startColumn, long endRow, long endColumn)
        {
            Cell c;
            List<List<UInt32Value>> output = new List<List<UInt32Value>>();
            for (long row = 0; row <= endRow - startRow; row++)
            {
                output.Add(new List<UInt32Value>());
                for (long column = 0; column <= endColumn - startColumn; column++)
                {
                    c = InsertCellInWorksheet(sheet, AddressToString(row + startRow, column + startColumn));
                    output.ElementAt((int)row).Add(c.StyleIndex);
                }
            }
            return output;

        }

        // Given a Worksheet and an address (like "AZ254"), either return a 
        // cell reference, or create the cell reference and return it.
        public static Cell InsertCellInWorksheet(Worksheet ws, string addressName)
        {
            SheetData sheetData = ws.GetFirstChild<SheetData>();
            Cell cell = null;

            UInt32 rowNumber = GetRowIndex(addressName);
            Row row = GetRow(sheetData, rowNumber);

            // If the cell you need already exists, return it.
            // If there is not a cell with the specified column name, insert one.  
            Cell refCell = row.Elements<Cell>().
                Where(c => c.CellReference.Value == addressName).FirstOrDefault();
            if (refCell != null)
            {
                cell = refCell;
            }
            else
            {
                cell = CreateCell(row, addressName);
            }
            return cell;
        }

        // Add a cell with the specified address to a row.
        private static Cell CreateCell(Row row, String address)
        {
            Cell cellResult;
            Cell refCell = null;

            // Cells must be in sequential order according to CellReference. 
            // Determine where to insert the new cell.
            foreach (Cell cell in row.Elements<Cell>())
            {
                if (string.Compare(cell.CellReference.Value, address, true) > 0)
                {
                    refCell = cell;
                    break;
                }
            }

            cellResult = new Cell();
            cellResult.CellReference = address;

            row.InsertBefore(cellResult, refCell);
            return cellResult;
        }

        // Return the row at the specified rowIndex located within
        // the sheet data passed in via wsData. If the row does not
        // exist, create it.
        private static Row GetRow(SheetData wsData, UInt32 rowIndex)
        {
            var row = wsData.Elements<Row>().
            Where(r => r.RowIndex.Value == rowIndex).FirstOrDefault();
            if (row == null)
            {
                row = new Row();
                row.RowIndex = rowIndex;
                wsData.Append(row);
            }
            return row;
        }

        // Given an Excel address such as E5 or AB128, GetRowIndex
        // parses the address and returns the row index.
        private static UInt32 GetRowIndex(string address)
        {
            string rowPart;
            UInt32 l;
            UInt32 result = 0;

            for (int i = 0; i < address.Length; i++)
            {
                if (UInt32.TryParse(address.Substring(i, 1), out l))
                {
                    rowPart = address.Substring(i, address.Length - i);
                    if (UInt32.TryParse(rowPart, out l))
                    {
                        result = l;
                        break;
                    }
                }
            }
            return result;
        }

        public static String AddressToString(long row, long col)
        {
            String address = "";
            long mod;
            do
            {
                mod = col % 26;
                address += (char)('A' + mod - 1);
                col /= 26;
            } while (col > 1);
            address += row;
            return address;
        }

        public static String[,] ArrayToList(List<List<String>> data)
        {
            String[,] list = new String[data.Count, data.ElementAt(0).Count];
            int i = 0, j = 0;
            foreach(List<String> stringList in data)
            {
                foreach (String s in stringList)
                {
                    list[i, j] = s;
                    j++;
                }
                i++;
                j = 0;
            }
            return list;
        }
    }

}
