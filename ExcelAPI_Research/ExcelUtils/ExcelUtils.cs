using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

// Additional namespaces
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelUtils
{
    public class ExcelUtils
    {
        public static Excel.Application OpenExcelApplication()
        {
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true;

            return excelApp;
        }

        public static bool CloseExcelApplication(Excel.Application excelApp)
        {
            bool isClosed = false;

            try
            {
                excelApp.Quit();
                isClosed = true;
            }
            catch (Exception)
            {
                Console.WriteLine("Fail to close current excel application!");
            }

            return isClosed;
        }

        public static string PromptUser(string message)
        {
            string input = string.Empty;

            do
            {
                Console.Write(message);
                input = Console.ReadLine();
            } while (!input.Any());

            return input;
        }

        public static Excel.Workbook OpenWorkbook(Excel.Application excelApp, string excelFilePath)
        {
            Excel.Workbook excelWorkbook = null;

            try
            {
                excelWorkbook = excelApp.Workbooks.Open(excelFilePath);
            }
            catch (Exception)
            {
                Console.WriteLine($"Cannot find the workbook located in {excelFilePath}. Please check it again!");
            }

            return excelWorkbook;
        }

        public static bool CloseWorkbook(Excel.Workbook excelWorkbook, bool isSaved)
        {
            bool isClosed = false;

            try
            {
                if (excelWorkbook != null)
                {
                    excelWorkbook.Close(isSaved);
                    isClosed = true;
                }
            }
            catch (Exception)
            {
                Console.WriteLine("Fail to close the current workbook!");
            }

            return isClosed;
        }

        public static string GetValueOfCellInRange(Excel.Range selectedRange, int rowIndex, int columnIndex)
        {
            string cellValue = string.Empty;

            try
            {
                cellValue = (selectedRange.Item[rowIndex, columnIndex]).Value();
            }
            catch (Exception)
            {
                Console.WriteLine($"The selected cell is not in a valid range!");
            }

            return cellValue;
        }

        public static void SetValueOfCellInRange(Excel.Range selectedRange, int rowIndex, int columnIndex, string value)
        {
            try
            {
                selectedRange.Item[rowIndex, columnIndex] = value;
            }
            catch (Exception)
            {
                Console.WriteLine($"The selected cell is not in a valid range!");
            }
        }

        public static Excel.Range SelectRangeInSheet(Excel.Worksheet currSheet, string startCell, string endCell)
        {
            return currSheet.Range[startCell, endCell];
        }

        public static List<Excel.Worksheet> GetWorksheetsList(Excel.Workbook workbook)
        {
            List<Excel.Worksheet> sheetList = new List<Excel.Worksheet>();
            Excel.Sheets sheets = workbook.Sheets;

            for (int i = 1; i <= sheets.Count; i++)
            {
                sheetList.Add(sheets.Item[i]);
            }

            return sheetList;
        }

        public static List<Excel.Range> SearchAllValuesInColumn(Excel.Worksheet worksheet, object columnIndex, object searchValue)
        {
            List<Excel.Range> results = new List<Excel.Range>();
            bool hasNext = true;
            Excel.Range searchRange = worksheet.Range[worksheet.Cells[1, columnIndex], worksheet.Cells[worksheet.UsedRange.Rows.Count, columnIndex]];
            Excel.Range tempCell = null;

            try
            {
                tempCell = searchRange.Find(What: searchValue, LookIn: Excel.XlFindLookIn.xlValues, LookAt: Excel.XlLookAt.xlPart);
                if (tempCell != null)
                {
                    results.Add(tempCell);

                    do
                    {
                        tempCell = searchRange.FindNext(After: results.LastOrDefault());
                        if (results.First().Row != tempCell.Row)
                        {
                            results.Add(tempCell);
                        }
                        else
                        {
                            hasNext = false;
                        }
                    } while (hasNext);

                }
            }
            catch (Exception)
            {
                throw;
            }

            return results;
        }

        public static List<Excel.Range> SearchAllValuesInRow(Excel.Worksheet worksheet, object rowIndex, object searchValue)
        {
            List<Excel.Range> results = new List<Excel.Range>();
            bool hasNext = true;
            Excel.Range searchRange = worksheet.Range[worksheet.Cells[rowIndex, 1], worksheet.Cells[rowIndex, worksheet.UsedRange.Columns.Count]];
            Excel.Range tempCell = null;

            try
            {
                tempCell = searchRange.Find(What: searchValue, LookIn: Excel.XlFindLookIn.xlValues, LookAt: Excel.XlLookAt.xlPart);
                if (tempCell != null)
                {
                    results.Add(tempCell);

                    do
                    {
                        tempCell = searchRange.FindNext(After: results.LastOrDefault());
                        if (results.First().Column != tempCell.Column)
                        {
                            results.Add(tempCell);
                        }
                        else
                        {
                            hasNext = false;
                        }
                    } while (hasNext);

                }
            }
            catch (Exception)
            {
                throw;
            }

            return results;
        }
    }
}
