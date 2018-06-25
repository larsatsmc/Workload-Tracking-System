using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Toolroom_Scheduler
{
    class ExcelInteractions
    {
        public QuoteInfo getQuoteInfo(string filePath = @"X:\TOOLROOM\Josh Meservey\Workload Tracking System\Simple Quote  Template - 2018-05-25.xlsx")
        {
            QuoteInfo quote;
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbooks workbooks = excelApp.Workbooks;
            Excel.Workbook workbook = workbooks.Open(Filename:filePath, Password:"ENG505");
            Excel.Worksheet quoteWorksheet = workbook.Sheets[1];
            Excel.Worksheet quoteLetter = workbook.Sheets[2];

            quote = new QuoteInfo(

                customer: quoteLetter.Cells[8, 3].value,
                partName: quoteLetter.Cells[10, 3].value,
                programRoughHours: Convert.ToInt16(quoteWorksheet.Cells[22, 8].value),
                programFinishHours: Convert.ToInt16(quoteWorksheet.Cells[23,8].value),
                programElectrodeHours: Convert.ToInt16(quoteWorksheet.Cells[24,8].value),
                cncRoughHours: Convert.ToInt16(quoteWorksheet.Cells[25,8].value),
                cncFinishHours: Convert.ToInt16(quoteWorksheet.Cells[26,8].value),
                cncElectrodeHours: Convert.ToInt16(quoteWorksheet.Cells[27,8].value),
                edmSinkerHours: Convert.ToInt16(quoteWorksheet.Cells[29,8].value)
                
                );

            Marshal.FinalReleaseComObject(quoteWorksheet);
            Marshal.FinalReleaseComObject(quoteLetter);

            workbook.Close(false);
            Marshal.FinalReleaseComObject(workbook);
            workbook = null;

            workbooks.Close();
            Marshal.FinalReleaseComObject(workbooks);
            workbooks = null;

            excelApp.Quit();
            Marshal.FinalReleaseComObject(excelApp);
            excelApp = null;

            return quote;

        }

        public void UpdateKanBanWorkbook(string filePath, ProjectInfo project)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbooks workbooks = excelApp.Workbooks;
            Excel.Workbook workbook = workbooks.Open(filePath);
        }

        public bool WorkbookHasMatchingComponent(Excel.Workbook workbook, string component)
        {
            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                if(sheet.Cells[1, 2].value.ToString().Trim() == component)
                {
                    return true;
                }
            }

            return false;
        }

        public void InsertNewComponentSheets()
        {

        }

        public void UpdateExistingComponentSheet()
        {

        }
    }
}
