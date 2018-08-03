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
                programFinishHours: Convert.ToInt16(quoteWorksheet.Cells[23,8].value),
                programElectrodeHours: Convert.ToInt16(quoteWorksheet.Cells[24,8].value + quoteWorksheet.Cells[21, 8].value),
                cncRoughHours: Convert.ToInt16(quoteWorksheet.Cells[25,8].value),
                cncFinishHours: Convert.ToInt16(quoteWorksheet.Cells[26,8].value),
                grindFittingHours: Convert.ToInt16(quoteWorksheet.Cells[28, 8].value),
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
            VBIDE.VBComponent wsMod;

            string activePrinterString, dateTime;
            int r;

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

                        if(task.StartDate == null)
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
                        if(task.DateCompleted != null)
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

                CreateKanBanComponentSheets(pi, excelApp, wb);

                string initialDirectory = "";

                if(pi.KanBanWorkbookPath != "")
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

                return "";
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

        // TODO: Find an alternative to this method that does not use COM interop.
        // FreeSpire is limited to 200 rows and 5 sheets.
        // My current installation of DevExpress can only generate spreadsheets.  Loading and editing are unavailable.  Can add subscription for $500.


        private void CreateKanBanComponentSheets(ProjectInfo pi, Excel.Application excelApp, Excel.Workbook wb)
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
                    if(task.DateCompleted != null)
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

                if(component.Notes.Contains('\n'))
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

                    ws = wb.Sheets.Add(After: wb.Sheets[index]);

                    component = pi.ComponentList.Find(x => x.Name == componentName);

                    CreateKanBanComponentSheet(pi, component, wb, ws.Index);

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

                excelApp.Visible = true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);

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
            }
        }

        private Excel.Worksheet CreateKanBanComponentSheet(ProjectInfo pi, Component component, Excel.Workbook wb, int sheetIndex)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Worksheet ws = wb.Sheets[sheetIndex];
            Excel.Borders border;
            int r, n;
            string dateTime;

            Console.WriteLine($"{ws.Name} {wb.Name}");

            n = 2;
            
            try
            {
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
                    if (task.DateCompleted != null)
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

                return ws;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);

                Marshal.ReleaseComObject(ws);
                ws = null;

                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
                excelApp = null;

                return null;
            }


        }

        private Boolean SheetNExists(string sheetname, Excel.Workbook wb)
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
                   "  If IsEmpty(Cells(row, 9)) = False And IsEmpty(Cells(row, 10)) = False And IsEmpty(Cells(row, 11)) = False Then\r\n" +
                   "\r\n" +
                   "    IsMarkedComplete = \"True\"\r\n" +
                   "\r\n" +
                   "  ElseIf IsEmpty(Cells(row, 9)) = True And IsEmpty(Cells(row, 10)) = True And IsEmpty(Cells(row, 11)) = True Then\r\n" +
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
                   "  If Target.column >= 9 And Target.column <= 11 Then\r\n" +
                   "\r\n" +
                   "    Dim leftHeaderArr As Variant\r\n" +
                   "    Dim Completed As String\r\n" +
                   "\r\n" +
                   "    Completed = IsMarkedComplete(Target.row)\r\n" +
                   "\r\n" +
                   "    leftHeaderArr = Split(Me.PageSetup.LeftHeader, \" \")\r\n" +
                   "\r\n" +
                   "    If Completed = \"True\" Then\r\n" +
                   "\r\n" +
                   "      ThisWorkbook.Save\r\n" +
                   "\r\n" +
                   "      Database.SetTaskAsCompleted _\r\n" +
                   "      jobNumber:=Cells(Target.row, 1).Value, _\r\n" +
                   "      projectNumber:=CLng(leftHeaderArr(2)), _\r\n" +
                   "      component:=Cells(Target.row, 2).Value, _\r\n" +
                   "      taskID:=CInt(Cells(Target.row, 3).Value), _\r\n" +
                   "      initials:=Cells(Target.row, 10).Value, _\r\n" +
                   "      dateCompleted:=Cells(Target.row, 11).Value\r\n" +
                   "\r\n" +
                   "    ElseIf Completed = \"False\" Then\r\n" +
                   "\r\n" +
                   "      ThisWorkbook.Save\r\n" +
                   "\r\n" +
                   "      Database.SetTaskAsIncomplete _\r\n" +
                   "      jobNumber:=Cells(Target.row, 1).Value, _\r\n" +
                   "      projectNumber:=CLng(leftHeaderArr(2)), _\r\n" +
                   "      component:=Cells(Target.row, 2).Value, _\r\n" +
                   "      taskID:=CInt(Cells(Target.row, 3).Value)\r\n" +
                   "\r\n" +
                   "    End If\r\n" +
                   "\r\n" +
                   "  End If\r\n" +
                   "\r\n" +
                   "End Sub\r\n"
                   ;
        }

        public bool WorkbookHasMatchingComponent(Excel.Workbook workbook, string component)
        {
            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                if(sheet.Cells[2, 2].value.ToString().Trim() == component)
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
                if (sheet.Cells[2, 2].value.ToString().Trim() == component)
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
