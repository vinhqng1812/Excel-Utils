using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

// Additional namespaces
using Excel = Microsoft.Office.Interop.Excel;
using MyExcelUtils = ExcelUtils.ExcelUtils;

/// <summary>
///     RESEARCH EXCEL INTEROP - TASK 1
/// Task from: 9/26/2018
/// Task end: 9/28/2018

/// Status: DONE (09/26/2018)

/// 1. Create a console application
/// 2. Prompt user to enter path
/// 3. Open excel from the enter path
/// 4. Prompt to enter a range(from cell to cell)
/// 5. Enter a cell in range
/// 6. Get the value in that cell

/// <Source: https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel?view=excel-pia >
/// </summary>

namespace Task1
{
    class Program
    {
        static void Main(string[] args)
        {
            Excel.Application excelApp = MyExcelUtils.OpenExcelApplication();
            string excelFilePath = MyExcelUtils.PromptUser("Type in excel file path: ");
            if (!excelFilePath.Equals("-q") || !excelFilePath.Equals("-Q"))
            {
                Excel.Workbook workbook = MyExcelUtils.OpenWorkbook(excelApp, excelFilePath);
                if (workbook != null)
                {
                    Excel.Sheets sheets = workbook.Sheets;
                    Excel.Worksheet worksheet = sheets.Item[1];

                    string startCell = MyExcelUtils.PromptUser("Enter start cell: ");
                    string endCell = MyExcelUtils.PromptUser("Enter end cell: ");

                    if ((!startCell.Equals("-q") && !endCell.Equals("-q")) ||
                        (!startCell.Equals("-Q") && !endCell.Equals("-Q")))
                    {
                        Excel.Range selectedRange = MyExcelUtils.SelectRangeInSheet(worksheet, startCell, endCell);
                        int.TryParse(MyExcelUtils.PromptUser("Enter row index: "), out int rowIndex);
                        int.TryParse(MyExcelUtils.PromptUser("Enter column index: "), out int columnIndex);

                        string cellValue = MyExcelUtils.GetValueOfCellInRange(selectedRange, rowIndex, columnIndex);

                        Console.WriteLine($"The value of cell({rowIndex}, {columnIndex}) is \"{cellValue}\"");
                    }
                    MyExcelUtils.CloseWorkbook(workbook, true);
                }
            }
            MyExcelUtils.CloseExcelApplication(excelApp);
        }
    }
}
