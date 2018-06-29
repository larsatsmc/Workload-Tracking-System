using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Toolroom_Scheduler
{
    class ExcelInteractions
    {
        public QuoteInfo GetQuoteInfo(string filePath = @"X:\TOOLROOM\Josh Meservey\Workload Tracking System\Simple Quote  Template - 2018-05-25.xlsx")
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

        private string getPrinterPort()
        {
            var devices = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows NT\CurrentVersion\Devices"); //Read-accessible even when using a locked-down account
            string printerName = "Microsoft XPS Document Writer";

            try
            {

                foreach (string name in devices.GetValueNames())
                {
                    if (Regex.IsMatch(name, printerName, RegexOptions.IgnoreCase))
                    {
                        var value = (String)devices.GetValue(name);
                        var port = Regex.Match(value, @"(Ne\d+:)", RegexOptions.IgnoreCase).Value;
                        //MessageBox.Show(printerName + " on " + port);
                        return port;
                    }
                }
                return "";
            }
            catch
            {
                throw;
            }
        }

        public void GenerateKanBanWorkbook(ProjectInfo pi)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbooks workbooks = excelApp.Workbooks;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            Excel.Borders border;

            string activePrinterString, dateTime;
            int r;
            Database db = new Database();

            try
            {
                excelApp.ScreenUpdating = false;
                excelApp.EnableEvents = false;
                excelApp.DisplayAlerts = false;
                excelApp.Visible = true;

                wb = workbooks.Open(@"X:\TOOLROOM\Workload Tracking System\Resource Files\Kan Ban Base File.xlsm", ReadOnly: true);

                // Remember active printer.
                activePrinterString = excelApp.ActivePrinter;

                // Change active printer to XPS Document Writer.
                excelApp.ActivePrinter = "Microsoft XPS Document Writer on " + getPrinterPort(); // This speeds up page setup operations.

                //ws = wb.Sheets[1];
                ws = wb.Sheets.Add(After: wb.Sheets[1]);
                ws.Name = pi.JobNumber;

                r = 1;

                ws.Cells[r, 1].value = "Job Number";
                ws.Cells[r, 2].value = "   Component";
                ws.Cells[r, 3].value = "Task ID";
                ws.Cells[r, 4].value = "   Task Name";
                ws.Cells[r, 5].value = "Duration";
                ws.Cells[r, 6].value = "Start Date";
                ws.Cells[r, 7].value = "Finish Date";
                ws.Cells[r, 8].value = "   Predecessors";
                ws.Cells[r, 9].value = "Status";
                ws.Cells[r, 10].value = "Initials";
                ws.Cells[r, 11].value = "Date";

                r = 2;

                ws.Range["H1"].EntireColumn.NumberFormat = "@";

                foreach (Component component in pi.ComponentList)
                {
                    border = ws.Range[ws.Cells[r - 1, 1], ws.Cells[r - 1, 11]].Borders;
                    border[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                    foreach (TaskInfo task in component.TaskList)
                    {
                        ws.Cells[r, 1].value = pi.JobNumber;
                        ws.Cells[r, 2].value = $"   {component.Name}";
                        ws.Cells[r, 3].value = task.ID;
                        ws.Cells[r, 4].value = $"   {task.TaskName}";
                        ws.Cells[r, 5].value = $"   {task.Duration}";
                        //ws.Cells[r, 6].value = nrow["StartDate"];
                        //ws.Cells[r, 7].value = nrow["FinishDate"];
                        ws.Cells[r, 8].value = $"  {task.Predecessors}";
                        ws.Cells[r, 9].value = task.Status;
                        ws.Cells[r, 10].value = task.Initials;
                        ws.Cells[r, 11].value = task.DateCompleted;

                        if (r % 2 == 0)
                            ws.Range[ws.Cells[r, 1], ws.Cells[r, 11]].Interior.Color = Excel.XlRgbColor.rgbPink;

                        r++;
                    }
                }

                border = ws.Range[ws.Cells[2, 1], ws.Cells[r - 1, 11]].Borders;

                border[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                border[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                border[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                border[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;

                for (int c = 2; c <= 11; c++)
                {
                    border = ws.Range[ws.Cells[2, c], ws.Cells[r - 1, c]].Borders;
                    border[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                }

                ws.Columns["B:B"].Autofit();
                ws.Columns["D:D"].Autofit();

                ws.Range["A1"].EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                ws.Range["A1"].EntireColumn.ColumnWidth = 11; // - 1
                ws.Range["B1"].EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                ws.Range["C1"].EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                ws.Range["C1"].EntireColumn.ColumnWidth = 6.25; // - 2.18
                ws.Range["E1"].EntireColumn.ColumnWidth = 7.71; // - .72
                ws.Range["F1:G1"].EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                ws.Range["F1:G1"].EntireColumn.ColumnWidth = 10.29; // - 2
                ws.Range["H1"].EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                ws.Range["H1"].EntireColumn.ColumnWidth = 13.25; // - 1
                ws.Range["I1"].EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                ws.Range["I1"].EntireColumn.ColumnWidth = 12; // 0
                ws.Range["J1"].EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                ws.Range["J1"].EntireColumn.ColumnWidth = 12.71; // + 4.28
                ws.Range["K1"].EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                ws.Range["K1"].EntireColumn.ColumnWidth = 10.43; // + 2

                //ws.Range[ws.Cells[1, 1], ws.Cells[1, 9]].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                ws.Range[ws.Cells[1, 1], ws.Cells[1, 11]].Font.Bold = true;

                ws.PageSetup.LeftHeader = "&\"Arial,Bold\"&18" + "Project #: " + pi.ProjectNumber;
                ws.PageSetup.CenterHeader = "&\"Arial,Bold\"&18" + "Lead: " + pi.ToolMaker;
                dateTime = DateTime.Today.ToShortDateString() + " " + DateTime.Now.ToShortTimeString();
                ws.PageSetup.RightHeader = "&\"Arial,Bold\"&18" + " Due Date: " + pi.DueDate.ToShortDateString();
                ws.PageSetup.RightFooter = "&\"Arial,Bold\"&12" + " Generated: " + dateTime;
                ws.PageSetup.HeaderMargin = excelApp.InchesToPoints(.2);
                ws.PageSetup.Zoom = 67;
                ws.PageSetup.TopMargin = excelApp.InchesToPoints(.5);
                ws.PageSetup.BottomMargin = excelApp.InchesToPoints(.5);
                ws.PageSetup.LeftMargin = excelApp.InchesToPoints(.2);
                ws.PageSetup.RightMargin = excelApp.InchesToPoints(.2);

                CreateKanBanComponentSheets(pi, excelApp, wb, ws);

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel files (*.xlsm)|*.xlsm";
                saveFileDialog.FilterIndex = 0;
                saveFileDialog.RestoreDirectory = true;
                saveFileDialog.CreatePrompt = false;
                saveFileDialog.FileName = pi.JobNumber + "- Proj #" + pi.ProjectNumber + " Checkoff Sheet";
                saveFileDialog.Title = "Save Path of Kan Ban";

                excelApp.Visible = true;
                excelApp.ScreenUpdating = true;
                excelApp.EnableEvents = true;
                excelApp.ActivePrinter = activePrinterString;

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Save. The selected path can be got with saveFileDialog.FileName.ToString()
                    wb.SaveAs(saveFileDialog.FileName.ToString());
                    db.setKanBanWorkbookPath(saveFileDialog.FileName.ToString(), pi.JobNumber, pi.ProjectNumber);
                }

                excelApp.DisplayAlerts = true;  // So I get prompted to save after adding pictures to the Kan Bans.

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message + " CreateKanBanWorkbook");

                // TODO: Need to close and release workbooks variable.
                // TODO: Need to remove garbage collection and have excel shutdown without it.

                wb.Close();
                Marshal.ReleaseComObject(wb);

                workbooks.Close();
                Marshal.ReleaseComObject(workbooks);

                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);


                Marshal.ReleaseComObject(ws);

            }

        }

        private void CreateKanBanComponentSheets(ProjectInfo pi, Excel.Application excelApp, Excel.Workbook wb, Excel.Worksheet ws)
        {
            Excel.Borders border;
            int r, n;
            string dateTime;

            n = 2;
            ws = wb.Sheets[1]; // Blank Sheet that contains VBA Code.

            ws.PageSetup.LeftHeader = "&\"Arial,Bold\"&18" + "Project #: " + pi.ProjectNumber;
            ws.PageSetup.CenterHeader = "&\"Arial,Bold\"&18" + "Lead: " + pi.ToolMaker;
            dateTime = DateTime.Today.ToShortDateString() + " " + DateTime.Now.ToShortTimeString();
            ws.PageSetup.RightHeader = "&\"Arial,Bold\"&18" + " Due Date: " + pi.DueDate.ToShortDateString();
            ws.PageSetup.RightFooter = "&\"Arial,Bold\"&12" + " Generated: " + dateTime;
            ws.PageSetup.HeaderMargin = excelApp.InchesToPoints(.2);
            ws.PageSetup.Zoom = 67;
            ws.PageSetup.TopMargin = excelApp.InchesToPoints(.5);
            ws.PageSetup.BottomMargin = excelApp.InchesToPoints(.5);
            ws.PageSetup.LeftMargin = excelApp.InchesToPoints(.2);
            ws.PageSetup.RightMargin = excelApp.InchesToPoints(.2);
            ws.Select();

            foreach (Component component in pi.ComponentList)
            {
                wb.Sheets[1].Copy(After: wb.Sheets[n++]);
                ws = wb.Sheets[n];

                Console.WriteLine(component.Name);

                if (component.Name.Length <= 31)
                {
                    ws.Name = component.Name;
                }
                else if (component.Name.Length > 31)
                {
                }
                else
                {
                    ws.Name = "Mold";
                }

                r = 1;

                ws.Cells[r, 1].value = "Job Number";
                ws.Cells[r, 2].value = "   Component";
                ws.Cells[r, 3].value = "Task ID";
                ws.Cells[r, 4].value = "   Task Name";
                ws.Cells[r, 5].value = "Duration";
                ws.Cells[r, 6].value = "Start Date";
                ws.Cells[r, 7].value = "Finish Date";
                ws.Cells[r, 8].value = "   Predecessors";
                ws.Cells[r, 9].value = "Status";
                ws.Cells[r, 10].value = "Initials";
                ws.Cells[r, 11].value = "Date";

                r++;

                ws.Range["H1"].EntireColumn.NumberFormat = "@";

                foreach (TaskInfo task in component.TaskList)
                {
                    border = ws.Range[ws.Cells[r - 1, 1], ws.Cells[r - 1, 11]].Borders;

                    ws.Cells[r, 1].NumberFormat = "@"; // Allows for a number with a 0 in front to be entered otherwise the 0 gets dropped.
                    ws.Cells[r, 1].value = pi.JobNumber;
                    ws.Cells[r, 2].value = "   " + component.Name;
                    ws.Cells[r, 3].value = task.ID;
                    ws.Cells[r, 4].value = "   " + task.TaskName;
                    ws.Cells[r, 5].value = "   " + task.Duration;
                    ws.Cells[r, 6].value = task.StartDate;
                    ws.Cells[r, 7].value = task.FinishDate;
                    ws.Cells[r, 8].value = "  " + task.Predecessors;
                    ws.Cells[r, 9].value = task.Status;
                    ws.Cells[r, 10].value = task.Initials;
                    ws.Cells[r, 11].value = task.DateCompleted;

                    if (r % 2 == 0)
                        ws.Range[ws.Cells[r, 1], ws.Cells[r, 11]].Interior.Color = Excel.XlRgbColor.rgbPink;

                    r++;

                }

                border = ws.Range[ws.Cells[2, 1], ws.Cells[r - 1, 11]].Borders;

                border[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                border[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                border[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                border[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;

                for (int c = 2; c <= 11; c++)
                {
                    border = ws.Range[ws.Cells[2, c], ws.Cells[r - 1, c]].Borders;
                    border[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                }

                ws.Columns["B:B"].Autofit();
                ws.Columns["D:D"].Autofit();

                ws.Range["A1"].EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                ws.Range["A1"].EntireColumn.ColumnWidth = 11; // - 1
                ws.Range["B1"].EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                ws.Range["C1"].EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                ws.Range["C1"].EntireColumn.ColumnWidth = 6.25; // - 2.18
                ws.Range["E1"].EntireColumn.ColumnWidth = 7.71; // - .72
                ws.Range["F1:G1"].EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                ws.Range["F1:G1"].EntireColumn.ColumnWidth = 10.29; // - 2
                ws.Range["H1"].EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                ws.Range["H1"].EntireColumn.ColumnWidth = 13.25; // - 1
                ws.Range["I1"].EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                ws.Range["I1"].EntireColumn.ColumnWidth = 12; // 0
                ws.Range["J1"].EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                ws.Range["J1"].EntireColumn.ColumnWidth = 12.71; // + 4.28
                ws.Range["K1"].EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                ws.Range["K1"].EntireColumn.ColumnWidth = 10.43; // + 2

                ws.Range[ws.Cells[1, 1], ws.Cells[1, 11]].Font.Bold = true;

                if(component.Picture != null)
                {
                    Clipboard.SetImage(component.Picture);
                    ws.Paste((Excel.Range)ws.Cells[r + 2, 2]);
                }
            }

            if (sheetNExists("Sheet1", wb))
            {
                wb.Sheets["Sheet1"].delete();
            }

            if (sheetNExists("Mold", wb))
                wb.Sheets["Mold"].Visible = Excel.XlSheetVisibility.xlSheetHidden;
        }

        private Boolean sheetNExists(string sheetname, Excel.Workbook wb)
        {
            foreach (Excel.Worksheet sheet in wb.Sheets)
            {
                if (sheet.Name == sheetname)
                {
                    return true;
                }
            }

            return false;
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
