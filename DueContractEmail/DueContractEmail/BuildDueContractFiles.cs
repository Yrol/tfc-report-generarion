using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Runtime.InteropServices;
using System.Threading;

namespace DueContractEmail
{
    class BuildDueContractFiles
    {
        private Microsoft.Office.Interop.Excel.Application excel = null;
        private Microsoft.Office.Interop.Excel.Workbook workBook = null;
        private Microsoft.Office.Interop.Excel.Worksheet workSheet = null;
        private Microsoft.Office.Interop.Excel.Range celLrangE = null;
        private Microsoft.Office.Interop.Excel.Font fontObj = null;
        private object m_objOpt = System.Reflection.Missing.Value;
        private DataTable tempBranchData;
        private Dictionary<string, string> fileEmailCollection;

        public BuildDueContractFiles(DataTable table, string currentPath)
        {
            Console.WriteLine("Building Excel files, please wait...");
            BranchData branchData = new BranchData();
            Dictionary<string, string> branchManagers = branchData.getBranchManagers();
            fileEmailCollection = new Dictionary<string, string>();
            Utils utils = new Utils();

            for (int count = 0; count < branchManagers.Count; count++)
            {
                //get branch manager details for branch manager collection
                var element = branchManagers.ElementAt(count);
                var branchCodeKey = (element.Key).ToUpper().ToString();
                var branchManagerEmail = element.Value;
                var fileName = "bm_" + utils.createFilename();

                //create temp branch data collection
                tempBranchData = new DataTable();
                tempBranchData.Columns.Add("Branch", typeof(string));
                tempBranchData.Columns.Add("Product", typeof(string));
                tempBranchData.Columns.Add("Contract No", typeof(string));
                tempBranchData.Columns.Add("Rental Amount", typeof(string));
                tempBranchData.Columns.Add("Rental No", typeof(string));
                tempBranchData.Columns.Add("Contract peroid", typeof(string));
                tempBranchData.Columns.Add("Debtor balance", typeof(string));
                tempBranchData.Columns.Add("Default Interest", typeof(string));
                tempBranchData.Columns.Add("No of days Over in arrears", typeof(string));
                tempBranchData.Columns.Add("Customer name", typeof(string));
                tempBranchData.Columns.Add("Address", typeof(string));
                tempBranchData.Columns.Add("Marketing officer", typeof(string));
                tempBranchData.Columns.Add("Collector name", typeof(string));

                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;
                workBook = excel.Workbooks.Add(Type.Missing);

                //collect excel object processes and add to a list
                int pid = -1;
                HandleRef hwnd = new HandleRef(excel, (IntPtr)excel.Hwnd);
                GetWindowThreadProcessId(hwnd, out pid);
                GlobalObjects.addPidsToList(pid);

                workSheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.ActiveSheet;
                workSheet.Name = "Contracts";
                workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[1, table.Columns.Count]].Merge();
                workSheet.Cells[1, 1] = "Due Contracts";
                workSheet.Cells.Font.Size = 15;

                foreach (DataRow datarow in table.Rows)
                {
                    string currentBranchCode = (datarow.Field<string>("Branch")).ToUpper().ToString();

                    if (branchCodeKey.Equals(currentBranchCode))
                    {
                        tempBranchData.ImportRow(datarow);
                    }
                }

                //create the branch specific excel file if data exists
                if (tempBranchData.Rows.Count > 0)
                {
                    try
                    {
                        int rowcount = 2;

                        //sort the data by product name
                        DataView dv = tempBranchData.DefaultView;
                        dv.Sort = "Product ASC";
                        DataTable sortedDT = dv.ToTable();

                        foreach (DataRow dr in sortedDT.Rows)
                        {
                            rowcount += 1;
                            for (int i = 1; i <= table.Columns.Count; i++)
                            {
                                if (rowcount == 3)
                                {
                                    workSheet.Cells[2, i] = table.Columns[i - 1].ColumnName;
                                    workSheet.Cells.Font.Color = System.Drawing.Color.Black;
                                }

                                workSheet.Cells[rowcount, i] = dr[i - 1].ToString();
                                workSheet.Cells[rowcount, i].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                                workSheet.Cells.Font.Size = 12;

                                if (rowcount > 3)
                                {
                                    if (i == table.Columns.Count)
                                    {
                                        if (rowcount % 2 == 0)
                                        {
                                            celLrangE = workSheet.Range[workSheet.Cells[rowcount, 1], workSheet.Cells[rowcount, table.Columns.Count]];
                                        }
                                    }
                                }
                            }
                        }

                        celLrangE = workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[rowcount, table.Columns.Count]];
                        celLrangE.EntireColumn.AutoFit();
                        Microsoft.Office.Interop.Excel.Borders border = celLrangE.Borders;
                        border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        border.Weight = 2d;
                        celLrangE = workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[2, table.Columns.Count]];

                        //bold and center title
                        celLrangE = workSheet.Range["A1", "M1"];
                        celLrangE.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                        fontObj = celLrangE.Font;
                        fontObj.Color = ConsoleColor.Red;

                        //bold header titles
                        celLrangE = workSheet.Range["A2", "M2"];
                        fontObj = celLrangE.Font;
                        fontObj.Bold = true;

                        //cell number format - Rentals
                        celLrangE = workSheet.Range["D3", "D"+rowcount];
                        celLrangE.NumberFormat = "0#.00";

                        //cell number format - Deter balance
                        celLrangE = workSheet.Range["G3", "G"+rowcount];
                        celLrangE.NumberFormat = "0#.00";

                        //cell number format - default interest
                        celLrangE = workSheet.Range["H3", "H"+rowcount];
                        celLrangE.NumberFormat = "0#.00";

                        //cell number format - No. of days over arrears
                        celLrangE = workSheet.Range["I3", "I"+rowcount];
                        celLrangE.NumberFormat = "0#.00";

                        workBook.SaveAs(currentPath + fileName, m_objOpt, m_objOpt,
                                m_objOpt, m_objOpt, m_objOpt, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                                m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt);
                        workBook.Close(false, m_objOpt, m_objOpt);
                        excel.Quit();
                        fileEmailCollection.Add(fileName, branchManagerEmail);
                        tempBranchData.Dispose();
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.ToString());
                    }
                    finally
                    {
                        utils.releaseObject(workBook);
                        utils.releaseObject(workSheet);
                        utils.releaseObject(excel);
                        utils.releaseObject(celLrangE);
                        utils.releaseObject(fontObj);
                    }
                }
            }

            Console.WriteLine("Creating excel files completed");

            //clear excel objects in a new thread using the pids in the pid list
            /*
            new Thread(() =>
            {
                Thread.CurrentThread.IsBackground = true;
                Console.WriteLine("Clear excel files BM thread");
                utils.killProcess(pidList, "EXCEL");
            }).Start();
            */

            //sending the emails
            if (fileEmailCollection.Count > 0)
            {
                SendEmail sendEmail = new SendEmail(fileEmailCollection, currentPath);
            }
        }


        //process references of the excel object
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int GetWindowThreadProcessId(HandleRef handle, out int processId);
    }
}
