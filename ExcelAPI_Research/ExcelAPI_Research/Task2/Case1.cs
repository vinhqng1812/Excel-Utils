using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;

// Additional namespaces
using Excel = Microsoft.Office.Interop.Excel;
using MyExcelUtils = ExcelUtils.ExcelUtils;
using System.Diagnostics;

/// <summary>
///     RESEARCH EXCEL INTEROP - TASK 2
/// Task from: 9/26/2018
/// Task end: 9/28/2018
/// 
/// Status: IN PROGRESS
/// 
/// Case 1: Search a value
/// 
/// <Source: https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel?view=excel-pia >
/// </summary>

namespace Task2
{
    class Case1
    {
        // Global constants
        const string excelFolderPath = "C:\\Users\\vinhqnguyen\\Documents\\0_VinhQNguyen\\Excel API Research\\TEMP\\";
        const string lookinXlValue = "";
        const string lookatXlPart = "";
        const string searchorderXlByColumns = "";
        const string searchorderXlByRows = "";
        const string searchdirectionXlNext = "";


        static void Main(string[] args)
        {
            string excelFilePath = string.Concat(excelFolderPath, "abc.xlsx");

            // Open excel appication
            Excel.Application excelApp = MyExcelUtils.OpenExcelApplication();

            // Open excel workbook
            Excel.Workbook workbook = MyExcelUtils.OpenWorkbook(excelApp, excelFilePath);
            if (workbook != null)
            {
                Excel.Sheets sheets = workbook.Sheets;
                Excel.Worksheet worksheet = sheets.Item["Test 1"];

                Stopwatch sw = new Stopwatch();
                Console.WriteLine("Worksheet name: Test 1");
                Console.WriteLine($"Number of rows in worksheet: {worksheet.UsedRange.Rows.Count}");
                Console.WriteLine($"Number of columns in worksheet: {worksheet.UsedRange.Columns.Count}");
                Console.WriteLine();

                ///
                Console.WriteLine($"Search All Values in Column - RESULT");
                sw.Start();
                List<Excel.Range> searchResults1 = SearchAllValuesInColumn(worksheet, 2, "12");
                sw.Stop();

                long timeInMiliSeconds1 = sw.ElapsedMilliseconds;
                double timeInSeconds1 = (double)timeInMiliSeconds1 / 1000;

                Console.WriteLine($"Found {searchResults1.Count} result(s)\n" +
                    $"\tin {timeInSeconds1} second(s)\n" +
                    $"\tin {timeInMiliSeconds1} milisecond(s)");
                Console.WriteLine();

                ///
                Console.WriteLine($"Search All Values in Row - RESULT");
                sw.Start();
                List<Excel.Range> searchResults2 = SearchAllValuesInRow(worksheet, 5, "12");
                sw.Stop();

                long timeInMiliSeconds2 = sw.ElapsedMilliseconds;
                double timeInSeconds2 = (double)timeInMiliSeconds2 / 1000;

                Console.WriteLine($"Found {searchResults2.Count} result(s)\n" +
                    $"\tin {timeInSeconds2} second(s)\n" +
                    $"\tin {timeInMiliSeconds2} milisecond(s)");


            }


            // Save and Close workbook
            //workbook.Save();
            MyExcelUtils.CloseWorkbook(workbook, false);

            // Close Excel application
            MyExcelUtils.CloseExcelApplication(excelApp);
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

        public static List<Excel.Range> SearchDirectionByColumns(Excel.Worksheet worksheet, string searchValue)
        {
            List<Excel.Range> result = new List<Excel.Range>();
            bool hasNext = true;
            result.Add(worksheet.UsedRange[RowIndex: 4].Find(What: searchValue, LookIn: Excel.XlFindLookIn.xlValues, LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByRows, SearchDirection: Excel.XlSearchDirection.xlNext));
            do
            {
                result.Add(worksheet.UsedRange[RowIndex: 4].FindNext(result.LastOrDefault()));
                if (result.Last() == null)
                {
                    hasNext = false;
                }

            } while (hasNext);

            return result;
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

        public static void CopyRanges(Excel.Worksheet worksheet)
        {
            int countRow = 0;
            int maxColCount = 0;
            int colPos = 0;
            int rowPos = 1;

            do
            {
                if (countRow == 0)
                {
                    colPos = 6;
                    maxColCount = 99;
                }
                else
                {
                    colPos = 1;
                    maxColCount = 200;
                }
                int countCol = 0;
                while (countCol < maxColCount)
                {
                    Excel.Range tempCell = worksheet.Cells[rowPos, colPos];
                    worksheet.Range["a1", "e4"].Copy(tempCell);

                    colPos += 5;
                    ++countCol;
                }
                rowPos += 4;
                ++countRow;
            } while (countRow < 300);
        }
    }
}
