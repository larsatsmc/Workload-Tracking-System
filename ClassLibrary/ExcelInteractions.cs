using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using VBIDE = Microsoft.Vbe.Interop;

namespace ClassLibrary
{
    public class ExcelInteractions
    {
        private static readonly string ColorPrinterString = "P-1336 HP CP5225 - Color";
        private readonly string XPSDocWriterString = "Microsoft XPS Document Writer";
        private static string KanBanBaseFilePath = @"X:\TOOLROOM\Workload Tracking System\Resource Files\Kan Ban Base File.xlsm";
        private static string KanBanSheetCode = 
                   "Function IsMarkedComplete(row As Integer) As String\r\n" +
                   "\r\n" +
                   "  If IsEmpty(Cells(row, 9)) = False And IsEmpty(Cells(row, 10)) = False Then\r\n" +
                   "\r\n" +
                   "    IsMarkedComplete = \"True\"\r\n" +
                   "\r\n" +
                   "  ElseIf IsEmpty(Cells(row, 9)) = True And IsEmpty(Cells(row, 10)) = True Then\r\n" +
                   "\r\n" +
                   "    IsMarkedComplete = \"False\"\r\n" +
                   "\r\n" +
                   "  Else\r\n" +
                   "\r\n" +
                   "    IsMarkedComplete = \"\r\n" +
                   "\r\n" +
                   "  End If\r\n" +
                   "\r\n" +
                   "End Function\r\n" +
                   "\r\n" +
                   "Private Sub Worksheet_Change(ByVal Target As Range)\r\n" +
                   "\r\n" +
                   "  Dim leftHeaderArr As Variant\r\n" +
                   "  Dim jobNumberArr As Variant\r\n" +
                   "  Dim componentArr As Variant\r\n" +
                   "  jobNumberArr = Split(Range(\"A1\").Value, \":\")\r\n" +
                   "  componentArr = Split(Range(\"A2\").Value, \":\")\r\n" +
                   "  leftHeaderArr = Split(Me.PageSetup.LeftHeader, \" \")\r\n" +
                   "\r\n" +
                   "  If Target.column = 9 Or Target.column = 10 Then\r\n" +
                   "\r\n" +
                   "    Dim Completed As String\r\n" +
                   "\r\n" +
                   "    Completed = IsMarkedComplete(Target.row)\r\n" +
                   "\r\n" +
                   "    ThisWorkbook.Save\r\n" +
                   "\r\n" +
                   "    If Completed = \"True\" Then\r\n" +
                   "\r\n" +
                   "      Database.SetTaskAsCompleted _\r\n" +
                   "      jobNumber:=jobNumberArr(1), _\r\n" +
                   "      projectNumber:=CLng(leftHeaderArr(2)), _\r\n" +
                   "      component:=componentArr(1), _\r\n" +
                   "      taskID:=CInt(Cells(Target.row, 1).Value), _\r\n" +
                   "      initials:=Cells(Target.row, 9).Value, _\r\n" +
                   "      dateCompleted:=Cells(Target.row, 10).Value\r\n" +
                   "\r\n" +
                   "    ElseIf Completed = \"False\" Then\r\n" +
                   "\r\n" +
                   "      Database.SetTaskAsIncomplete _\r\n" +
                   "      jobNumber:=jobNumberArr(1), _\r\n" +
                   "      projectNumber:=CLng(leftHeaderArr(2)), _\r\n" +
                   "      component:=componentArr(1), _\r\n" +
                   "      taskID:=CInt(Cells(Target.row, 1).Value)\r\n" +
                   "\r\n" +
                   "    End If\r\n" +
                   "\r\n" +
                   "  ElseIf Target.column = 7 Then\r\n" +
                   "\r\n" +
                   "    ThisWorkbook.Save\r\n" +
                   "\r\n" +
                   "    Database.SetNote _\r\n" +
                   "    jobNumber:=jobNumberArr(1), _\r\n" +
                   "    projectNumber:=CLng(leftHeaderArr(2)), _\r\n" +
                   "    component:=componentArr(1), _\r\n" +
                   "    taskID:=CInt(Cells(Target.row, 1).Value), _\r\n" +
                   "    notes:=Cells(Target.row, 7).Value\r\n" +
                   "\r\n" +
                   "  End If\r\n" +
                   "\r\n" +
                   "End Sub\r\n"
                   ;

        public QuoteModel GetQuoteInfo(string filePath = @"X:\TOOLROOM\Josh Meservey\Workload Tracking System\Simple Quote  Template - 2018-05-25.xlsx")
        {
            QuoteModel quote;            
            Excel.Application excelApp;
            //var app = excelApp.Application;
            Excel.Workbooks workbooks;
            Excel.Workbook workbook;
            Excel.Sheets worksheets;
            Excel.Worksheet quoteWorksheet;
            Excel.Worksheet quoteLetter;
            Excel.Range quoteLetterCells;
            Excel.Range quoteWorksheetCells;

            excelApp = new Excel.Application();
            workbooks = excelApp.Workbooks;
            workbook = workbooks.Open(Filename: filePath, Password: "ENG505");
            worksheets = workbook.Worksheets;
            quoteWorksheet = worksheets["Quote Worksheet"];
            quoteLetter = worksheets["Quote Letter"];
            quoteLetterCells = quoteLetter.Cells;
            quoteWorksheetCells = quoteWorksheet.Cells;

            try
            {

                //TODO: Change this method to read the Quote sheet by iterating through the list of tasks rather then by reading from specified task locations.

                quote = new QuoteModel(

                    customer: quoteLetterCells[8, 3].value,
                    partName: quoteLetterCells[10, 3].value,
                    designHours: quoteWorksheetCells[18, 8].value,
                    designElectrodeHours: quoteWorksheetCells[21, 8].value,
                    programRoughHours: quoteWorksheetCells[22, 8].value,
                    programFinishHours: quoteWorksheetCells[23, 8].value,
                    programElectrodeHours: quoteWorksheetCells[24, 8].value,
                    cncRoughHours: quoteWorksheetCells[25, 8].value,
                    cncFinishHours: quoteWorksheetCells[26, 8].value,
                    grindFittingHours: quoteWorksheetCells[28, 8].value,
                    cncElectrodeHours: quoteWorksheetCells[27, 8].value,
                    edmSinkerHours: quoteWorksheetCells[29, 8].value

                    );

            }
            finally
            {
                while (Marshal.ReleaseComObject(quoteLetterCells) != 0);
                while (Marshal.ReleaseComObject(quoteWorksheetCells) != 0);
                while (Marshal.ReleaseComObject(quoteWorksheet) != 0);
                while (Marshal.ReleaseComObject(quoteLetter) != 0);
                while (Marshal.ReleaseComObject(worksheets) != 0);

                quoteLetterCells = null;
                quoteWorksheetCells = null;
                quoteWorksheet = null;
                quoteLetter = null;
                worksheets = null;

                workbook.Close(0);
                while (Marshal.ReleaseComObject(workbook) != 0);
                workbook = null;

                workbooks.Close();
                while (Marshal.ReleaseComObject(workbooks) != 0);
                workbooks = null;

                //app.Quit();
                excelApp.Quit();
                while (Marshal.ReleaseComObject(excelApp) != 0);
                excelApp = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }


            return quote;
        }

        private static bool DeviceExists(string device)
        {
            var devices = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows NT\CurrentVersion\Devices"); //Read-accessible even when using a locked-down account

            return devices.GetValueNames().ToList().Exists(x => x.ToString().Contains(device));
        }

        private static string GetDevice(string device)
        {
            var devices = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows NT\CurrentVersion\Devices"); //Read-accessible even when using a locked-down account

            return devices.GetValueNames().ToList().FirstOrDefault(x => x.ToString().Contains(device));
        }

        private static string GetDevicePort(string device)
        {
            var devices = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows NT\CurrentVersion\Devices"); //Read-accessible even when using a locked-down account
            string printerName = device;

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

                MessageBox.Show(printerName + " is not installed on this computer.");
                return "";
            }
            catch
            {
                throw;
            }
        }

        public string GenerateKanBanWorkbook(ProjectModel pi)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbooks workbooks = excelApp.Workbooks;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            Excel.Borders border;
            VBIDE.VBComponents vBComponents;

            string activePrinterString, dateTime;
            int r, n;

            try
            {
                excelApp.ScreenUpdating = false;
                excelApp.EnableEvents = false;
                excelApp.DisplayAlerts = false;

                wb = workbooks.Open(KanBanBaseFilePath, ReadOnly: true);

                // Set to color printer.
                if (DeviceExists(ColorPrinterString))
                {
                    excelApp.ActivePrinter = GetDevice(ColorPrinterString) + " on " + GetDevicePort(ColorPrinterString);
                }

                // Remember active printer.
                activePrinterString = excelApp.ActivePrinter;

                // Change active printer to XPS Document Writer.
                if (DeviceExists(XPSDocWriterString))
                {
                    excelApp.ActivePrinter = XPSDocWriterString + " on " + GetDevicePort(XPSDocWriterString); // This speeds up page setup operations.
                }

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
                ws.Cells[r, 8].value = "   Hours";
                ws.Cells[r, 9].value = "Status";
                ws.Cells[r, 10].value = "Initials";
                ws.Cells[r, 11].value = "Date";

                r = 2;

                ws.Range["H1"].EntireColumn.NumberFormat = "@";

                foreach (ComponentModel component in pi.Components)
                {
                    border = ws.Range[ws.Cells[r - 1, 1], ws.Cells[r - 1, 11]].Borders;
                    border[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                    foreach (TaskModel task in component.Tasks)
                    {
                        ws.Cells[r, 1].value = pi.JobNumber;
                        ws.Cells[r, 2].value = $"   {component.Component}";
                        ws.Cells[r, 3].value = task.TaskID;
                        ws.Cells[r, 4].value = $"   {task.TaskName}";
                        ws.Cells[r, 5].value = $"   {task.Duration}";

                        if(task.StartDate != null)
                        {
                            ws.Cells[r, 6].value = task.StartDate;
                        }

                        if (task.FinishDate != null)
                        {
                            ws.Cells[r, 7].value = task.FinishDate;
                        }

                        ws.Cells[r, 8].value = $"  {task.Hours}";
                        ws.Cells[r, 9].value = task.Status;
                        ws.Cells[r, 10].value = task.Initials;
                        if (task.DateCompleted != null)
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
                ws.PageSetup.Zoom = 71;
                ws.PageSetup.TopMargin = excelApp.InchesToPoints(.5);
                ws.PageSetup.BottomMargin = excelApp.InchesToPoints(.5);
                ws.PageSetup.LeftMargin = excelApp.InchesToPoints(.2);
                ws.PageSetup.RightMargin = excelApp.InchesToPoints(.2);

                ws = wb.Sheets[1];

                //FormatComponentSheet(pi, ws);

                n = 2;

                vBComponents = wb.VBProject.VBComponents;

                foreach (ComponentModel component in pi.Components)
                {
                    wb.Sheets.Add(After: wb.Sheets[n++]);
                    //wb.Sheets[1].Copy(After: wb.Sheets[n++]);

                    ws = wb.Sheets[n];

                    PopulateKanBanComponentSheet(pi, component, ws);

                    vBComponents = wb.VBProject.VBComponents;

                    foreach (VBIDE.VBComponent wsMod in vBComponents)
                    {
                        if (wsMod.Name == ws.CodeName)
                        {
                            Console.WriteLine($"{wsMod.Name} is {ws.Name}");
                            wsMod.CodeModule.AddFromString(KanBanSheetCode);
                        }
                    }
                }

                if (SheetNExists("Sheet1", wb))
                {
                    wb.Sheets["Sheet1"].delete();
                }

                if (SheetNExists("Mold", wb))
                    wb.Sheets["Mold"].Visible = Excel.XlSheetVisibility.xlSheetHidden;

                ws = wb.Sheets.Add(After: wb.Sheets[--n]);

                CreateHoursSheet(pi, wb, ws.Index);

                string initialDirectory = GetInitialDirectory(pi.KanBanWorkbookPath);

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel files (*.xlsm)|*.xlsm";
                saveFileDialog.FilterIndex = 0;
                saveFileDialog.RestoreDirectory = true;
                saveFileDialog.CreatePrompt = false;
                saveFileDialog.InitialDirectory = initialDirectory;
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
                    return saveFileDialog.FileName.ToString();
                }

                excelApp.DisplayAlerts = true;  // So I get prompted to save after adding pictures to the Kan Bans.

                return "";
            }
            catch(Exception ex)
            {

                if (wb != null)
                {
                    wb.Close();
                    Marshal.ReleaseComObject(wb); 
                }

                workbooks.Close();
                Marshal.ReleaseComObject(workbooks);

                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);


                if (ws != null)
                {
                    Marshal.ReleaseComObject(ws); 
                }

                MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);

                return "";
            }
        }

        private void FormatComponentSheet(ProjectModel pi, Excel.Worksheet ws)
        {
            Excel.Application excelApp = new Excel.Application();
            //Excel.Worksheet ws;
            string dateTime;

            //ws = wb.Sheets[1]; // Blank Sheet that contains VBA Code.

            ws.PageSetup.LeftHeader = "&\"Arial,Bold\"&18" + "Project #: " + pi.ProjectNumber;
            ws.PageSetup.CenterHeader = "&\"Arial,Bold\"&18" + "Lead: " + pi.ToolMaker;
            dateTime = DateTime.Today.ToShortDateString() + " " + DateTime.Now.ToShortTimeString();
            ws.PageSetup.RightHeader = "&\"Arial,Bold\"&18" + " Due Date: " + pi.DueDate.ToShortDateString();
            ws.PageSetup.RightFooter = "&\"Arial,Bold\"&12" + " Generated: " + dateTime;
            ws.PageSetup.HeaderMargin = excelApp.InchesToPoints(.2);
            ws.PageSetup.Zoom = 79;
            ws.PageSetup.TopMargin = excelApp.InchesToPoints(.5);
            ws.PageSetup.BottomMargin = excelApp.InchesToPoints(.5);
            ws.PageSetup.LeftMargin = excelApp.InchesToPoints(.2);
            ws.PageSetup.RightMargin = excelApp.InchesToPoints(.2);

            // Task ID
            ws.Range["A1"].EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            ws.Range["A1"].EntireColumn.ColumnWidth = 6.29;
            // Task Name
            ws.Range["B1"].EntireColumn.ColumnWidth = 27.4;
            // Duration
            ws.Range["C1"].EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            ws.Range["C1"].EntireColumn.ColumnWidth = 9.86;
            // Start Date & Finish Date
            ws.Range["D1:E1"].EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            ws.Range["D1:E1"].EntireColumn.ColumnWidth = 10.86;
            // Hours
            ws.Range["F1"].EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            ws.Range["F1"].EntireColumn.ColumnWidth = 6;
            // Notes
            ws.Range["G1"].EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            ws.Range["G1"].EntireColumn.ColumnWidth = 55.79;
            // Date
            ws.Range["H1"].EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            ws.Range["H1"].EntireColumn.ColumnWidth = 8.57;
            // Initials
            ws.Range["I1"].EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            ws.Range["I1"].EntireColumn.ColumnWidth = 10.43;

            //ws.Range["A1"].Select();
            ws.Select();
        }

        // FreeSpire is limited to 200 rows and 5 sheets.
        // My current installation of DevExpress can only generate spreadsheets.  Loading and editing are unavailable.  Can add subscription for $500.

        private int CreateKanBanComponentSheets(ProjectModel pi, Excel.Application excelApp, Excel.Workbook wb)
        {
            Excel.Worksheet ws;
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

            foreach (ComponentModel component in pi.Components)
            {
                wb.Sheets[1].Copy(After: wb.Sheets[n++]);
                ws = wb.Sheets[n];

                Console.WriteLine(component.Component);

                if (component.Component.Length <= 31)
                {
                    ws.Name = component.Component;
                }
                else if (component.Component.Length > 31)
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

                foreach (TaskModel task in component.Tasks)
                {
                    border = ws.Range[ws.Cells[r - 1, 1], ws.Cells[r - 1, 11]].Borders;

                    ws.Cells[r, 1].NumberFormat = "@"; // Allows for a number with a 0 in front to be entered otherwise the 0 gets dropped.
                    ws.Cells[r, 1].value = pi.JobNumber;
                    ws.Cells[r, 2].value = "   " + component.Component;
                    ws.Cells[r, 3].value = task.TaskID;
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
                //ws.Columns["D:D"].Autofit();

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

                if (component.Notes.Contains('\n'))
                {
                    foreach (string line in component.Notes.Split('\n'))
                    {
                        ws.Cells[r++ + 1, 2].value = line;
                    }
                }
                else
                {
                    ws.Cells[r++ + 1, 2].value = component.Notes;
                }


                if (component.Picture != null)
                {
                    Clipboard.SetImage(component.picture);
                    ws.Paste((Excel.Range)ws.Cells[r + 2, 2]);
                }
            }

            if (SheetNExists("Sheet1", wb))
            {
                wb.Sheets["Sheet1"].delete();
            }

            if (SheetNExists("Mold", wb))
                wb.Sheets["Mold"].Visible = Excel.XlSheetVisibility.xlSheetHidden;

            return n;
        }
        public static string GenerateKanBanWorkbook2(ProjectModel pi)
        {
            //try
            //{
            string kanBanSavePath = ChooseKanBanSavePath(pi);

            if (kanBanSavePath == "")
            {
                return "";
            }

            //Stopwatch sw = new Stopwatch();
            //sw.Start();

            using (var wb = new XLWorkbook(KanBanBaseFilePath, XLEventTracking.Disabled))
            {
                var wsBase = wb.Worksheet(1);
                var componentInfoCellRange = wsBase.Range("A1:J3");
                int[] borderedColumnsArr = new int[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 };
                int r;
                int taskRowCount;

                wsBase.PageSetup.Header.Left.AddText("Project #: " + pi.ProjectNumber).SetBold().SetFontSize(18);
                wsBase.PageSetup.Header.Center.AddText("Lead: " + pi.ToolMaker).SetBold().SetFontSize(18);
                wsBase.PageSetup.Header.Right.AddText("Due Date: " + pi.DueDate.ToString("M/d/yyyy")).SetBold().SetFontSize(18);

                //wsBase.Range("A1:G1").Merge();
                //wsBase.Range("A2:G2").Merge();
                //wsBase.Range("A3:G3").Merge();

                //wsBase.Range("H1:J1").Merge();
                //wsBase.Range("H2:J2").Merge();
                //wsBase.Range("H3:J3").Merge();

                //componentInfoCellRange.Style.Fill.BackgroundColor = XLColor.White;
                //wsBase.Range("A1:G1").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                //wsBase.Range("A1:A3").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                //wsBase.Range("A3:G3").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                //wsBase.Range("A3:G3").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                //wsBase.Range("H1:J1").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                //wsBase.Range("J1:J3").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                //wsBase.Range("H1:H3").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                //wsBase.Range("H3:J3").Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                wsBase.Range("A5:J5").Style.Font.Bold = true;

                wsBase.Cell(5, 1).Value = "Task ID";
                wsBase.Cell(5, 2).Value = "   Task Name";
                wsBase.Cell(5, 3).Value = "Duration";
                wsBase.Cell(5, 4).Value = "Start Date";
                wsBase.Cell(5, 5).Value = "Finish Date";
                wsBase.Cell(5, 6).Value = "Hours";
                wsBase.Cell(5, 7).Value = "Notes";
                wsBase.Cell(5, 9).Value = "Initials";
                wsBase.Cell(5, 10).Value = "Date";

                foreach (ComponentModel component in pi.Components)
                {
                    taskRowCount = component.Tasks.Count();
                    var componentWs = wb.Worksheet(1).CopyTo(component.Component);
                    r = 6;

                    foreach (int c in borderedColumnsArr)
                    {
                        componentWs.Range(componentWs.Cell(6, c), componentWs.Cell(5 + taskRowCount, c)).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                    }

                    componentWs.Range("A6:J6").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                    componentWs.Range($"A6:A{5 + taskRowCount}").Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                    componentWs.Range($"A{5 + taskRowCount}:J{5 + taskRowCount}").Style.Border.BottomBorder = XLBorderStyleValues.Thin;

                    componentWs.Range("A1").Value = " Job Number: " + pi.JobNumber;
                    componentWs.Range("A2").Value = " Component: " + component.Component;
                    componentWs.Range("A3").Value = " Material: " + component.Material;

                    componentWs.Range("H1").Value = " Qty: " + component.Quantity;
                    componentWs.Range("H2").Value = " Spares: " + component.Spares;
                    componentWs.Range("H3").Value = " Finish: " + component.Finish;

                    foreach (TaskModel task in component.Tasks)
                    {
                        if (r % 2 == 1)
                        {
                            componentWs.Range(componentWs.Cell(r, 1), componentWs.Cell(r, 10)).Style.Fill.BackgroundColor = XLColor.Pink;
                        }
                        else
                        {
                            componentWs.Range(componentWs.Cell(r, 1), componentWs.Cell(r, 10)).Style.Fill.BackgroundColor = XLColor.White;
                        }

                        componentWs.Range(componentWs.Cell(r, 7), componentWs.Cell(r, 8)).Merge();

                        componentWs.Cell(r, 1).Value = task.TaskID;
                        componentWs.Cell(r, 2).Value = task.TaskName;
                        componentWs.Cell(r, 3).Value = task.Duration;
                        componentWs.Cell(r, 4).Value = task.StartDate;
                        componentWs.Cell(r, 5).Value = task.FinishDate;
                        componentWs.Cell(r, 6).Value = task.Hours;
                        componentWs.Cell(r, 6).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        componentWs.Cell(r, 7).Value = task.Notes;
                        componentWs.Cell(r, 9).Value = task.Initials;
                        componentWs.Cell(r++, 10).Value = task.DateCompleted;
                    }

                    //ws.Range(ws.Cell(6, 1), ws.Cell(taskRowCount + 5, 8));
                    componentWs.Cell(taskRowCount + 7, 1).Value = "Notes: " + component.Notes;
                    componentWs.Cell(taskRowCount + 7, 1).Style.Font.Bold = true;

                    var noteArea = componentWs.Range(componentWs.Cell(taskRowCount + 7, 1), componentWs.Cell(taskRowCount + 10, 10));
                    noteArea.Style.Fill.BackgroundColor = XLColor.White;
                    noteArea.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    noteArea.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                    noteArea.Style.Alignment.WrapText = true;
                    noteArea.Merge();

                    componentWs.Range(componentWs.Cell(taskRowCount + 7, 1), componentWs.Cell(taskRowCount + 10, 1)).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                    componentWs.Range(componentWs.Cell(taskRowCount + 10, 1), componentWs.Cell(taskRowCount + 10, 10)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    componentWs.Range(componentWs.Cell(taskRowCount + 7, 1), componentWs.Cell(taskRowCount + 7, 10)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                    componentWs.Range(componentWs.Cell(taskRowCount + 7, 10), componentWs.Cell(taskRowCount + 10, 10)).Style.Border.RightBorder = XLBorderStyleValues.Thin;

                    //var noteContentsRange = componentWs.Range(componentWs.Cell(taskRowCount + 7, 1), componentWs.Cell(taskRowCount + 10, 10));
                    //noteContentsRange.Style.Fill.BackgroundColor = XLColor.White;
                    //noteContentsRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    //noteContentsRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                    //noteContentsRange.Style.Alignment.WrapText = true;
                    //noteContentsRange.Merge();

                    if (component.picture != null)
                    {
                        var image = componentWs.AddPicture((Bitmap)component.picture);
                        image.MoveTo(componentWs.Cell(taskRowCount + 12, 2));
                    }

                    componentWs.PageSetup.PrintAreas.Clear();

                    //componentWs.PageSetup.PrintAreas.Add("A1:J61");
                }

                var sumWs = wb.Worksheets.Add("Summary");

                var col1 = sumWs.Column("A");
                col1.Style.Font.Bold = true;
                sumWs.Range("A1").Value = "Work Type";
                sumWs.Range("B1").Style.Font.Bold = true;
                sumWs.Range("B1").Value = "Total Hours";

                ProjectSummary ps = GetProjectSummary(pi);

                r = 2;

                foreach (Hours hour in ps.HoursList)
                {
                    sumWs.Cell(r, 1).Value = hour.WorkType;
                    sumWs.Cell(r++, 2).Value = hour.Qty;
                }

                sumWs.Cell(r, 1).Value = "Total";
                sumWs.Cell(r, 2).FormulaA1 = "=Sum(" + sumWs.Cell(r - 1, 2).Address + ":" + sumWs.Cell(r - ps.HoursList.Count, 2).Address + ")";
                sumWs.Cell(r, 2).Style.Font.Bold = true;

                wsBase.Delete();

                wb.SaveAs(kanBanSavePath);
                InsertVbaCode(kanBanSavePath); // Printer set to color here.


                //sw.Stop();
                //MessageBox.Show($"Kan Ban Generated: {sw.ElapsedMilliseconds}");

                Process.Start(kanBanSavePath); // Opens up generated Kan Ban.

                return kanBanSavePath;
            }
            
            //}
            //catch (Exception ex)
            //{
            //    throw ex;
            //}
        }
        public static string ChooseKanBanSavePath(ProjectModel project)
        {
            string initialDirectory = GetInitialDirectory(project.KanBanWorkbookPath);

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel files (*.xlsm)|*.xlsm";
            saveFileDialog.FilterIndex = 0;
            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.CreatePrompt = false;
            saveFileDialog.InitialDirectory = initialDirectory;
            saveFileDialog.FileName = project.JobNumber + "- Proj #" + project.ProjectNumber + " Checkoff Sheet";
            saveFileDialog.Title = "Save Path of Kan Ban";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                //wb.SaveAs(saveFileDialog.FileName.ToString());
                //InsertVbaCode(saveFileDialog.FileName.ToString()); // This code also sets the printer to the color printer.
                //Process.Start(saveFileDialog.FileName.ToString());
                return saveFileDialog.FileName.ToString();
            }
            else
            {
                MessageBox.Show("No Kan Ban was created since a file location wasn't chosen.");
                return "";
            }
        }
        public static void InsertVbaCode(string destinationFilePath)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbooks workbooks = excelApp.Workbooks;
            Excel.Workbook workbook = workbooks.Open(destinationFilePath);
            VBIDE.VBComponents vBComponents = workbook.VBProject.VBComponents;

            // Set to color printer.
            if (DeviceExists(ColorPrinterString))
            {
                excelApp.ActivePrinter = GetDevice(ColorPrinterString) + " on " + GetDevicePort(ColorPrinterString);
            }

            try
            {
                foreach (Excel.Worksheet ws in workbook.Worksheets)
                {
                    if (ws.Name != "Summary")
                    {
                        foreach (VBIDE.VBComponent wsMod in vBComponents)
                        {
                            if (wsMod.Name == ws.CodeName)
                            {
                                Console.WriteLine($"{wsMod.Name} is {ws.Name}");
                                wsMod.CodeModule.AddFromString(KanBanSheetCode);
                            }
                        }
                    }
                }

                workbook.Save();
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message + "\n\n" + ex.StackTrace);

                throw new Exception(ex.Message + "\n\n" + ex.StackTrace);
            }
            finally
            {
                if (workbook != null)
                {
                    workbook.Close();
                    Marshal.ReleaseComObject(workbook);
                }

                workbooks.Close();
                Marshal.ReleaseComObject(workbooks);

                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }
        }
        private void PopulateKanBanComponentSheet(ProjectModel pi, ComponentModel component, Excel.Worksheet ws)
        {
            Excel.Borders border;
            int r;

            // Checks if sheet has been formatted.  If it hasn't then format it.
            if (ws.PageSetup.LeftHeader == "")
            {
                FormatComponentSheet(pi, ws);
            }

            Console.WriteLine(component.Component);

            if (component.Component.Length <= 31)
            {
                ws.Name = component.Component;
            }
            else if (component.Component.Length > 31)
            {
            }
            else
            {
                ws.Name = "Mold";
            }

            ws.Range["D1"].EntireColumn.Hidden = true;
            ws.Range["E1"].EntireColumn.Hidden = true;

            Excel.Shape textBox = ws.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 300, 65);
            textBox.TextFrame2.TextRange.Characters.Text = "Job Number: " + pi.JobNumber + "\n" + "Component: " + component.Component + "\n" + "Material: " + component.Material;
            textBox.TextFrame2.TextRange.Font.Size = 14;
            textBox.TextFrame2.TextRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue;
            textBox.ShapeStyle = Microsoft.Office.Core.MsoShapeStyleIndex.msoShapeStylePreset1;

            Excel.Shape textBox2 = ws.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 315, 0, 150, 65);
            textBox2.TextFrame2.TextRange.Characters.Text = "Qty: " + component.Quantity + "\n" + "Spares: " + component.Spares + "\n" + "Finish: " + component.Finish;
            textBox2.TextFrame2.TextRange.Font.Size = 14;
            textBox2.TextFrame2.TextRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue;
            textBox2.ShapeStyle = Microsoft.Office.Core.MsoShapeStyleIndex.msoShapeStylePreset1;

            //ws.Range["A1:I73"].Font.Size = 12;

            r = 6;

            ws.Range[ws.Cells[r, 1], ws.Cells[r, 9]].Font.Bold = true;

            ws.Cells[r, 1].value = "Task ID";
            ws.Cells[r, 2].value = "   Task Name";
            ws.Cells[r, 3].value = "Duration";
            ws.Cells[r, 4].value = "Start Date";
            ws.Cells[r, 5].value = "Finish Date";
            ws.Cells[r, 6].value = "Hours";
            ws.Cells[r, 7].value = "Notes";
            ws.Cells[r, 8].value = "Initials";
            ws.Cells[r, 9].value = "Date";

            r++;

            ws.Range["F1"].EntireColumn.NumberFormat = "@";

            foreach (TaskModel task in component.Tasks)
            {
                border = ws.Range[ws.Cells[r - 1, 1], ws.Cells[r - 1, 9]].Borders;

                ws.Cells[r, 1].value = task.TaskID;
                ws.Cells[r, 2].value = "   " + task.TaskName;
                ws.Cells[r, 3].value = "" + task.Duration;
                ws.Cells[r, 4].value = " " + String.Format("{0:M/d/yyyy}", task.StartDate);
                ws.Cells[r, 5].value = " " + String.Format("{0:M/d/yyyy}", task.FinishDate);
                ws.Cells[r, 6].value = task.Hours;
                ws.Cells[r, 7].value = task.Notes;
                ws.Cells[r, 8].value = task.Initials;
                ws.Cells[r, 9].value = task.DateCompleted;

                if (r % 2 == 0)
                    ws.Range[ws.Cells[r, 1], ws.Cells[r, 9]].Interior.Color = Excel.XlRgbColor.rgbPink;

                r++;
            }

            border = ws.Range[ws.Cells[7, 1], ws.Cells[r - 1, 9]].Borders;

            border[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            border[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            border[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            border[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;

            for (int c = 2; c <= 9; c++)
            {
                border = ws.Range[ws.Cells[7, c], ws.Cells[r - 1, c]].Borders;
                border[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            }

            //ws.Columns["B:B"].Autofit();

            //if (component.Notes.Contains('\n'))
            //{
            //    foreach (string line in component.Notes.Split('\n'))
            //    {
            //        ws.Cells[r++ + 1, 2].value = line;
            //    }
            //}
            //else
            //{
            //    ws.Cells[r++ + 1, 2].value = component.Notes;
            //}

            //ws.Cells[r++ + 1, 2].Top();
            //ws.Range[r++, 2].Left();

            Excel.Shape textBox3 = ws.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 0, ws.Cells[r + 1, 1].Top(), 677, 47);
            textBox3.TextFrame2.TextRange.Characters.Text = "Notes: " + component.Notes;
            textBox3.TextFrame2.TextRange.Font.Size = 11;
            textBox3.TextFrame2.TextRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue;
            textBox3.ShapeStyle = Microsoft.Office.Core.MsoShapeStyleIndex.msoShapeStylePreset1;

            if (component.Picture != null)
            {
                Clipboard.SetImage(component.picture);
                ws.Paste((Excel.Range)ws.Cells[r + 5, 2]);  // This line throws an error when Brian tries to make a KanBan.
            }
        }

        public void OpenKanBanWorkbook(string filepath, string component)
        {
            if (filepath != null && filepath != "")
            {
                FileInfo fi = new FileInfo(filepath);

                if (fi.Exists)
                {
                    Excel.Application excelApp = new Excel.Application();
                    Excel.Workbook workbook = excelApp.Workbooks.Open(fi.FullName);

                    try
                    {
                        //var attributes = File.GetAttributes(fi.FullName);    

                        foreach (Excel.Worksheet ws in workbook.Worksheets)
                        {
                            if (ws.Name.Trim() == component)
                            {
                                workbook.Sheets[ws.Index].Select();
                                workbook.Save();
                            }
                        }

                        workbook.Close();
                        excelApp.Quit();

                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        Marshal.ReleaseComObject(workbook);

                        //Marshal.ReleaseComObject(ws);
                        Marshal.ReleaseComObject(excelApp);

                        var res = Process.Start("EXCEL.EXE", "\"" + fi.FullName + "\"");

                    }
                    catch (Exception e)
                    {
                        MessageBox.Show(e.Message);

                        //workbook.Close();
                        excelApp.Quit();
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        Marshal.ReleaseComObject(workbook);

                        //Marshal.ReleaseComObject(ws);
                        Marshal.ReleaseComObject(excelApp);
                    }
                }
                else
                {
                    MessageBox.Show("Can't find a Kan Ban Workbook with path " + filepath + ".");
                }
            }
            else
            {
                MessageBox.Show("No Kan Ban exists for this project.");
            }

        }

        public void EditKanBanWorkbook(ProjectModel pi, string kanBanWorkbookPath, List<string> componentsList)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbooks workbooks = excelApp.Workbooks;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            int index = 0;
            VBIDE.VBComponents vBComponents;
            ComponentModel component;
            int compareResult;

            try
            {
                wb = workbooks.Open(kanBanWorkbookPath);

                // Set to color printer.
                if (DeviceExists(ColorPrinterString))
                {
                    excelApp.ActivePrinter = GetDevice(ColorPrinterString) + " on " + GetDevicePort(ColorPrinterString);
                }

                vBComponents = wb.VBProject.VBComponents;

                foreach (string componentName in componentsList)
                {
                    if (WorkbookHasMatchingComponent(wb, componentName))
                    {
                        //ShowSheetIndexes(wb);

                        ws = MatchingComponentSheet(wb, componentName);

                        index = ws.Index - 1;

                        excelApp.DisplayAlerts = false;

                        ws.Delete();

                        excelApp.DisplayAlerts = true;
                    }
                    else
                    {
                        foreach (Excel.Worksheet sheet in wb.Sheets)
                        {
                            compareResult = sheet.Name.CompareTo(componentName);

                            if (compareResult < 0 && sheet.Index > 1)
                            {
                                index = sheet.Index - 1;
                            }
                        }

                        if (index == 0)
                        {
                            index = wb.Sheets.Count - 1;
                        }
                    }

                    //ShowSheetIndexes(wb);

                    ws = wb.Sheets.Add(After: wb.Sheets[index]);

                    component = pi.Components.Find(x => x.Component == componentName);

                    PopulateKanBanComponentSheet(pi, component, ws);

                    vBComponents = wb.VBProject.VBComponents;

                    foreach (VBIDE.VBComponent wsMod in vBComponents)
                    {
                        if (wsMod.Name == ws.CodeName)
                        {
                            Console.WriteLine($"{wsMod.Name} is {ws.Name}");
                            wsMod.CodeModule.AddFromString(KanBanSheetCode);
                        }
                    }
                }

                if (!WorkbookHasSummarySheet(wb))
                {
                    ws = wb.Sheets.Add(After: wb.Sheets[index + 1]);
                }
                else
                {
                    ws = wb.Sheets[SheetNIndex("Summary", wb)];
                }

                CreateHoursSheet(pi, wb, ws.Index);

                excelApp.Visible = true;
            }
            catch (Exception)
            {
                if (ws != null)
                {
                    Marshal.ReleaseComObject(ws);
                    ws = null;
                }


                if (wb != null)
                {
                    wb.Close(false);
                    Marshal.ReleaseComObject(wb);
                    wb = null; 
                }

                workbooks.Close();
                Marshal.ReleaseComObject(workbooks);
                workbooks = null;

                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
                excelApp = null;

                throw;
            }
        }

        private void CreateHoursSheet(ProjectModel pi, Excel.Workbook wb, int sheetIndex)
        {
            //Excel.Application excelApp = new Excel.Application();
            Excel.Worksheet ws = wb.Sheets[sheetIndex];
            ProjectSummary ps = GetProjectSummary(pi);
            int r = 1, c = 1;

            ws.Name = "Summary";
            ws.Range["A1"].ColumnWidth = 20;
            ws.Range["A1"].EntireColumn.Font.Bold = true;
            ws.Range["A1"].EntireRow.Font.Bold = true;
            ws.Cells[r, c].Value = "Work Type";
            ws.Cells[r++, c + 1].Value = "Total Hours";

            foreach (Hours hour in ps.HoursList)
            {
                ws.Cells[r, c].Value = hour.WorkType;
                ws.Cells[r++, c + 1].Value = hour.Qty;
            }

            ws.Cells[r, c].Value = "Total";
            ws.Cells[r, c + 1].Formula = "=Sum(" + ws.Cells[r - 1, c + 1].Address + ":" + ws.Cells[r - ps.HoursList.Count , c + 1].Address + ")";
            ws.Cells[r, c + 1].Font.Bold = true;
        }

        private static ProjectSummary GetProjectSummary(ProjectModel pi)
        {
            List<TaskModel> taskList = new List<TaskModel>();
            List<TaskModel> summaryTaskList = new List<TaskModel>();
            ProjectSummary ps = new ProjectSummary();

            foreach (ComponentModel component in pi.Components)
            {
                taskList.AddRange(component.Tasks);
            }

            foreach (Hours hours in ps.HoursList)
            {
                if (hours.WorkType == "Ordering")
                {
                    summaryTaskList = taskList.FindAll(x => x.TaskName.Contains("Order"));
                }
                else if (hours.WorkType == "Design")
                {
                    summaryTaskList = taskList.FindAll(x => x.TaskName.Contains("Design"));
                }
                else if (hours.WorkType == "Grind")
                {
                    summaryTaskList = taskList.FindAll(x => x.TaskName.Contains("Grind"));
                }
                else if (hours.WorkType == "Inspection")
                {
                    summaryTaskList = taskList.FindAll(x => x.TaskName.Contains("Inspection"));
                }
                else
                {
                    summaryTaskList = taskList.FindAll(x => x.TaskName == hours.WorkType);
                }

                ps.HoursList.Find(x => x.WorkType == hours.WorkType).Qty += summaryTaskList.Sum(p => p.Hours);
            }

            return ps;
        }

        private class ProjectSummary
        {
            public List<Hours> HoursList { get; set; }

            public ProjectSummary()
            {
                HoursList = new List<Hours>();

                string[] workTypeArr = {"Design", "Ordering", "Program Rough", "Program Finish", "Program Electrodes",
                                   "CNC Rough", "CNC Finish", "CNC Electrodes", "Grind", "EDM Sinker", "EDM Wire (In-House)", "Inspection"};

                foreach (string workType in workTypeArr)
                {
                    HoursList.Add(new Hours { WorkType = workType });
                }
            }
        }

        private class Hours
        {
            public int Qty { get; set; }
            public string WorkType { get; set; }
        }
        private static string GetInitialDirectory(string path)
        {
            string initialDirectory;

            if (File.Exists(path))
            {
                initialDirectory = path.Substring(0, path.LastIndexOf('\\'));
            }
            else
            {
                initialDirectory = @"X:\TOOLROOM\";
            }

            return initialDirectory;
        }
        private Boolean SheetNExists(string sheetName, Excel.Workbook wb)
        {
            foreach (Excel.Worksheet sheet in wb.Sheets)
            {
                if (sheet.Name == sheetName)
                {
                    return true;
                }
            }

            return false;
        }

        private int SheetNIndex(string sheetName, Excel.Workbook wb)
        {
            foreach (Excel.Worksheet sheet in wb.Sheets)
            {
                if (sheet.Name == sheetName)
                {
                    return sheet.Index;
                }
            }

            return 0;
        }

        private void ShowSheetIndexes(Excel.Workbook wb)
        {
            foreach (Excel.Worksheet sheet in wb.Sheets)
            {
                Console.WriteLine(sheet.Name + " " + sheet.Index);
            }
        }

        public void UpdateKanBanWorkbook(string filePath, ProjectModel project)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbooks workbooks = excelApp.Workbooks;
            Excel.Workbook workbook = workbooks.Open(filePath);
            Excel.Worksheet matchingSheet;
            VBIDE.VBComponents vBComponents;
            VBIDE.VBComponent wsMod;
            int matchingSheetIndex;

            foreach (ComponentModel component in project.Components)
            {
                matchingSheet = MatchingComponentSheet(workbook, component.Component);

                if (matchingSheet != null)
                {
                    matchingSheetIndex = matchingSheet.Index;
                    matchingSheet.Delete();
                    workbook.Sheets[1].Copy(After: workbook.Sheets[matchingSheetIndex - 1]);
                    vBComponents = workbook.VBProject.VBComponents;
                    wsMod = vBComponents.Item(1);

                    wsMod.CodeModule.AddFromString(KanBanSheetCode);
                }
                else
                {

                }
            }
        }

        public bool WorkbookHasMatchingComponent(Excel.Workbook workbook, string component)
        {
            string componentSheet = "";
            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.Cells[2, 2].value != null)
                {
                    componentSheet = sheet.Cells[2, 2].value.ToString().Trim();
                }
                else
                {
                    foreach (Excel.Shape shape in sheet.Shapes)
                    {
                        if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoTextBox)
                        {
                            if (shape.TextFrame2.TextRange.Characters.Text.Contains("Job Number:") && shape.TextFrame2.TextRange.Characters.Text.Contains("Component:"))
                            {
                                componentSheet = shape.TextFrame2.TextRange.Characters.Text.Split('\n')[1].Split(':')[1].Trim();
                            }
                        }
                    }
                }

                if(componentSheet == component)
                {
                    return true;
                }
            }

            return false;
        }

        public bool WorkbookHasSummarySheet(Excel.Workbook workbook)
        {
            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.Name == "Summary")
                {
                    return true;
                }
            }

            return false;
        }

        public Excel.Worksheet MatchingComponentSheet(Excel.Workbook workbook, string component)
        {
            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.Index != 1)
                {
                    if (sheet.Cells[2, 2].value != null && sheet.Cells[2, 2].value.ToString().Trim() == component)
                    {
                        return sheet;
                    }
                    else
                    {
                        foreach (Excel.Shape shape in sheet.Shapes)
                        {
                            if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoTextBox)
                            {
                                if (shape.TextFrame2.TextRange.Characters.Text.Contains("Job Number:") && shape.TextFrame2.TextRange.Characters.Text.Contains("Component:"))
                                {
                                    if(shape.TextFrame2.TextRange.Characters.Text.Split('\n')[1].Split(':')[1].Trim()  == component)
                                    {
                                        return sheet;
                                    }
                                }
                            }
                        }
                    }
                }

            }

            return null;
        }
    }
}
