using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
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
        public QuoteInfo GetQuoteInfo(string filePath = @"X:\TOOLROOM\Josh Meservey\Workload Tracking System\Simple Quote  Template - 2018-05-25.xlsx")
        {
            QuoteInfo quote;
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbooks workbooks = excelApp.Workbooks;
            Excel.Workbook workbook;
            Excel.Worksheet quoteWorksheet;
            Excel.Worksheet quoteLetter;

            workbook = workbooks.Open(Filename: filePath, Password: "ENG505");
            quoteWorksheet = workbook.Sheets[1];
            quoteLetter = workbook.Sheets[2];

            quote = new QuoteInfo(

                customer: quoteLetter.Cells[8, 3].value,
                partName: quoteLetter.Cells[10, 3].value,
                programRoughHours: Convert.ToInt16(quoteWorksheet.Cells[22, 8].value),
                programFinishHours: Convert.ToInt16(quoteWorksheet.Cells[23, 8].value),
                programElectrodeHours: Convert.ToInt16(quoteWorksheet.Cells[24, 8].value + quoteWorksheet.Cells[21, 8].value),
                cncRoughHours: Convert.ToInt16(quoteWorksheet.Cells[25, 8].value),
                cncFinishHours: Convert.ToInt16(quoteWorksheet.Cells[26, 8].value),
                grindFittingHours: Convert.ToInt16(quoteWorksheet.Cells[28, 8].value),
                cncElectrodeHours: Convert.ToInt16(quoteWorksheet.Cells[27, 8].value),
                edmSinkerHours: Convert.ToInt16(quoteWorksheet.Cells[29, 8].value)

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



        private string GetPrinterPort()
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

        public string GenerateKanBanWorkbook(ProjectInfo pi)
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

                wb = workbooks.Open(@"X:\TOOLROOM\Workload Tracking System\Resource Files\Kan Ban Base File.xlsm", ReadOnly: true);

                // Remember active printer.
                activePrinterString = excelApp.ActivePrinter;

                // Change active printer to XPS Document Writer.
                excelApp.ActivePrinter = "Microsoft XPS Document Writer on " + GetPrinterPort(); // This speeds up page setup operations.

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

                        if (task.StartDate == null)
                        {

                        }
                        else
                        {
                            ws.Cells[r, 6].value = task.StartDate;
                        }

                        ws.Cells[r, 7].value = task.FinishDate;
                        ws.Cells[r, 8].value = $"  {task.Predecessors}";
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
                ws.PageSetup.Zoom = 67;
                ws.PageSetup.TopMargin = excelApp.InchesToPoints(.5);
                ws.PageSetup.BottomMargin = excelApp.InchesToPoints(.5);
                ws.PageSetup.LeftMargin = excelApp.InchesToPoints(.2);
                ws.PageSetup.RightMargin = excelApp.InchesToPoints(.2);

                ws = wb.Sheets[1];

                //FormatComponentSheet(pi, ws);

                n = 2;

                vBComponents = wb.VBProject.VBComponents;

                foreach (Component component in pi.ComponentList)
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
                            wsMod.CodeModule.AddFromString(KanBanSheetCode());
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

                string initialDirectory = "";

                if (pi.KanBanWorkbookPath != "")
                {
                    initialDirectory = pi.KanBanWorkbookPath.Substring(0, pi.KanBanWorkbookPath.LastIndexOf('\\'));
                }

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
            catch (Exception ex)
            {
                //MessageBox.Show(e.Message + " GenerateKanBanWorkbook");

                // TODO: Need to close and release workbooks variable.
                // TODO: Need to remove garbage collection and have excel shutdown without it.

                wb.Close();
                Marshal.ReleaseComObject(wb);

                workbooks.Close();
                Marshal.ReleaseComObject(workbooks);

                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);


                Marshal.ReleaseComObject(ws);

                throw ex;

                //return "";
            }

            //vBComponents = wb.VBProject.VBComponents;
            //wsMod = vBComponents.Item(3);

            //for (int i = 1; i <= vBComponents.Count; i++)
            //{
            //    MessageBox.Show(vBComponents.Item(i).Name);
            //}

            //string macroCode = "Sub main()\r\n" +
            //                   "   MsgBox \"Hello world\"\r\n" +
            //                   "end Sub";

            //wsMod.CodeModule.AddFromString(macroCode);
        }

        private void FormatComponentSheet(ProjectInfo pi, Excel.Worksheet ws)
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
            ws.PageSetup.Zoom = 75;
            ws.PageSetup.TopMargin = excelApp.InchesToPoints(.5);
            ws.PageSetup.BottomMargin = excelApp.InchesToPoints(.5);
            ws.PageSetup.LeftMargin = excelApp.InchesToPoints(.2);
            ws.PageSetup.RightMargin = excelApp.InchesToPoints(.2);

            // Task ID
            ws.Range["A1"].EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            ws.Range["A1"].EntireColumn.ColumnWidth = 6.29;
            // Task Name
            ws.Range["B1"].EntireColumn.ColumnWidth = 31.29;
            // Duration
            ws.Range["C1"].EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            ws.Range["C1"].EntireColumn.ColumnWidth = 9.86;
            // Start Date & Finish Date
            ws.Range["D1:E1"].EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            ws.Range["D1:E1"].EntireColumn.ColumnWidth = 10.86;
            // Predecessors
            ws.Range["F1"].EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            ws.Range["F1"].EntireColumn.ColumnWidth = 13.57;
            // Notes
            ws.Range["G1"].EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            ws.Range["G1"].EntireColumn.ColumnWidth = 57.86;
            // Date
            ws.Range["H1"].EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            ws.Range["H1"].EntireColumn.ColumnWidth = 8.57;
            // Initials
            ws.Range["I1"].EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            ws.Range["I1"].EntireColumn.ColumnWidth = 10.43;

            //ws.Range["A1"].Select();
            ws.Select();
        }

        // TODO: Find an alternative to this method that does not use COM interop.
        // FreeSpire is limited to 200 rows and 5 sheets.
        // My current installation of DevExpress can only generate spreadsheets.  Loading and editing are unavailable.  Can add subscription for $500.

        private int CreateKanBanComponentSheets(ProjectInfo pi, Excel.Application excelApp, Excel.Workbook wb)
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
                    Clipboard.SetImage(component.Picture);
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

        private void PopulateKanBanComponentSheet(ProjectInfo pi, Component component, Excel.Worksheet ws)
        {
            Excel.Borders border;
            int r;

            // Checks if sheet has been formatted.  If it hasn't then format it.
            if (ws.PageSetup.LeftHeader == "")
            {
                FormatComponentSheet(pi, ws);
            }

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

            ws.Range["D1"].EntireColumn.Hidden = true;
            ws.Range["E1"].EntireColumn.Hidden = true;

            Excel.Shape textBox = ws.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 300, 65);
            textBox.TextFrame2.TextRange.Characters.Text = "Job Number: " + pi.JobNumber + "\n" + "Component: " + component.Name + "\n" + "Material: " + component.Material;
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
            ws.Cells[r, 6].value = " Predecessors";
            ws.Cells[r, 7].value = "Notes";
            ws.Cells[r, 8].value = "Initials";
            ws.Cells[r, 9].value = "Date";

            r++;

            ws.Range["F1"].EntireColumn.NumberFormat = "@";

            foreach (TaskInfo task in component.TaskList)
            {
                border = ws.Range[ws.Cells[r - 1, 1], ws.Cells[r - 1, 9]].Borders;

                ws.Cells[r, 1].value = task.ID;
                ws.Cells[r, 2].value = "   " + task.TaskName;
                ws.Cells[r, 3].value = "" + task.Duration;
                ws.Cells[r, 4].value = " " + String.Format("{0:M/d/yyyy}", task.StartDate);
                ws.Cells[r, 5].value = " " + String.Format("{0:M/d/yyyy}", task.FinishDate);
                ws.Cells[r, 6].value = "  " + task.Predecessors;
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

            ws.Columns["B:B"].Autofit();

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

            Excel.Shape textBox3 = ws.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 0, ws.Cells[r + 1, 1].Top(), 700, 47);
            textBox3.TextFrame2.TextRange.Characters.Text = "Notes: " + component.Notes;
            textBox3.TextFrame2.TextRange.Font.Size = 11;
            textBox3.TextFrame2.TextRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue;
            textBox3.ShapeStyle = Microsoft.Office.Core.MsoShapeStyleIndex.msoShapeStylePreset1;

            if (component.Picture != null)
            {
                Clipboard.SetImage(component.Picture);
                ws.Paste((Excel.Range)ws.Cells[r + 5, 2]);
            }
        }

        public void OpenKanBanWorkbook(string filepath, string component)
        {
            //Excel.Worksheet ws;

            if (filepath != null)
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
                MessageBox.Show("There is no Kan Ban Workbook for this project.");
            }

        }

        public void EditKanBanWorkbook(ProjectInfo pi, string kanBanWorkbookPath, List<string> componentsList)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbooks workbooks = excelApp.Workbooks;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            int index = 0;
            VBIDE.VBComponents vBComponents;
            Component component;

            try
            {
                wb = workbooks.Open(kanBanWorkbookPath);

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
                            if (sheet.Name.CompareTo(componentName) < 0)
                            {
                                index = sheet.Index;
                            }
                        }
                    }

                    //ShowSheetIndexes(wb);

                    ws = wb.Sheets.Add(After: wb.Sheets[index]);

                    component = pi.ComponentList.Find(x => x.Name == componentName);

                    PopulateKanBanComponentSheet(pi, component, ws);

                    vBComponents = wb.VBProject.VBComponents;

                    foreach (VBIDE.VBComponent wsMod in vBComponents)
                    {
                        if (wsMod.Name == ws.CodeName)
                        {
                            Console.WriteLine($"{wsMod.Name} is {ws.Name}");
                            wsMod.CodeModule.AddFromString(KanBanSheetCode());
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
                    Marshal.ReleaseComObject(ws);
                ws = null;

                wb.Close(false);
                Marshal.ReleaseComObject(wb);
                wb = null;

                workbooks.Close();
                Marshal.ReleaseComObject(workbooks);
                workbooks = null;

                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
                excelApp = null;

                throw;
            }
        }

        private void CreateHoursSheet(ProjectInfo pi, Excel.Workbook wb, int sheetIndex)
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

        private ProjectSummary GetProjectSummary(ProjectInfo pi)
        {
            List<TaskInfo> taskList = new List<TaskInfo>();
            List<TaskInfo> summaryTaskList = new List<TaskInfo>();
            ProjectSummary ps = new ProjectSummary();

            foreach (Component component in pi.ComponentList)
            {
                taskList.AddRange(component.TaskList);
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

        public void UpdateKanBanWorkbook(string filePath, ProjectInfo project)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbooks workbooks = excelApp.Workbooks;
            Excel.Workbook workbook = workbooks.Open(filePath);
            Excel.Worksheet matchingSheet;
            VBIDE.VBComponents vBComponents;
            VBIDE.VBComponent wsMod;
            int matchingSheetIndex;

            foreach (Component component in project.ComponentList)
            {
                matchingSheet = MatchingComponentSheet(workbook, component.Name);

                if (matchingSheet != null)
                {
                    matchingSheetIndex = matchingSheet.Index;
                    matchingSheet.Delete();
                    workbook.Sheets[1].Copy(After: workbook.Sheets[matchingSheetIndex - 1]);
                    vBComponents = workbook.VBProject.VBComponents;
                    wsMod = vBComponents.Item(1);

                    wsMod.CodeModule.AddFromString(KanBanSheetCode());
                }
                else
                {

                }
            }
        }

        private string KanBanSheetCode()
        {
            // \r moves cursor to beginning of line.
            // \n moves cursor down one line.
            return "Function IsMarkedComplete(row As Integer) As String\r\n" +
                   "\r\n" +
                   "  If IsEmpty(Cells(row, 8)) = False And IsEmpty(Cells(row, 9)) = False Then\r\n" +
                   "\r\n" +
                   "    IsMarkedComplete = \"True\"\r\n" +
                   "\r\n" +
                   "  ElseIf IsEmpty(Cells(row, 8)) = True And IsEmpty(Cells(row, 9)) = True Then\r\n" +
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
                   "Function GetTextBoxInfo() As String()\r\n" +
                   "\r\n" +
                   "  Dim shape as Shape\r\n" +
                   "  Dim infoArr(0 To 1) as String\r\n" +
                   "\r\n" +
                   "    For Each shape In Me.Shapes\r\n" +
                   "\r\n" +
                   "      If shape.Type = msoTextBox Then\r\n" +
                   "\r\n" +
                   "        If shape.TextFrame2.TextRange.Characters.Text Like \"*Component*\" Then\r\n" +
                   "\r\n" +
                   "          detailArr = Split(shape.TextFrame2.TextRange.Characters.Text, vbLf)\r\n" +
                   "          detailArr2 = Split(detailArr(0), \":\")\r\n" +
                   "          infoArr(0) = Trim(detailArr2(1)) ' Job Number\r\n" +
                   "          detailArr2 = Split(detailArr(1), \":\")\r\n" +
                   "          infoArr(1) = Trim(detailArr2(1)) ' Component\r\n" +
                   "\r\n" +
                   "          GetTextBoxInfo = infoArr\r\n" +
                   "\r\n" +
                   "        End If\r\n" +
                   "\r\n" +
                   "      End If\r\n" +
                   "\r\n" +
                   "    Next\r\n" +
                   "\r\n" +
                   "End Function\r\n" +
                   "\r\n" +
                   "Private Sub Worksheet_Change(ByVal Target As Range)\r\n" +
                   "\r\n" +
                   "  Dim infoArr() As String\r\n" +
                   "  Dim leftHeaderArr As Variant\r\n" +
                   "  infoArr = GetTextBoxInfo\r\n" +
                   "  leftHeaderArr = Split(Me.PageSetup.LeftHeader, \" \")\r\n" +
                   "\r\n" +
                   "  If Target.column = 8 Or Target.column = 9 Then\r\n" +
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
                   "      jobNumber:=infoArr(0), _\r\n" +
                   "      projectNumber:=CLng(leftHeaderArr(2)), _\r\n" +
                   "      component:=infoArr(1), _\r\n" +
                   "      taskID:=CInt(Cells(Target.row, 1).Value), _\r\n" +
                   "      initials:=Cells(Target.row, 8).Value, _\r\n" +
                   "      dateCompleted:=Cells(Target.row, 9).Value\r\n" +
                   "\r\n" +
                   "    ElseIf Completed = \"False\" Then\r\n" +
                   "\r\n" +
                   "      Database.SetTaskAsIncomplete _\r\n" +
                   "      jobNumber:=infoArr(0), _\r\n" +
                   "      projectNumber:=CLng(leftHeaderArr(2)), _\r\n" +
                   "      component:=infoArr(1), _\r\n" +
                   "      taskID:=CInt(Cells(Target.row, 1).Value)\r\n" +
                   "\r\n" +
                   "    End If\r\n" +
                   "\r\n" +
                   "  ElseIf Target.column = 7 Then\r\n" +
                   "\r\n" +
                   "    ThisWorkbook.Save\r\n" +
                   "\r\n" +
                   "    Database.SetNote _\r\n" +
                   "    jobNumber:=infoArr(0), _\r\n" +
                   "    projectNumber:=CLng(leftHeaderArr(2)), _\r\n" +
                   "    component:=infoArr(1), _\r\n" +
                   "    taskID:=CInt(Cells(Target.row, 1).Value), _\r\n" +
                   "    notes:=Cells(Target.row, 7).Value\r\n" +
                   "\r\n" +
                   "  End If\r\n" +
                   "\r\n" +
                   "End Sub\r\n"
                   ;
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
                            if (shape.TextFrame2.TextRange.Characters.Text.Contains("Component"))
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
                if (sheet.Cells[2, 2].value.ToString().Trim() == component && sheet.Index != 1)
                {
                    return sheet;
                }
            }

            return null;
        }

        public void InsertNewComponentSheets()
        {

        }

        public void UpdateExistingComponentSheet()
        {

        }
    }
}
