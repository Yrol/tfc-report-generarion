using System;
using System.Collections.Generic;
using System.IO;
using System.Diagnostics;

namespace DueContractEmail
{
    class Utils
    {
        private const string directoryName = "com.tfc.duecontracts";
        public string path = "E:\\" + directoryName + "\\";

        public Utils() { }

        //get current time in a format
        public string createFilename()
        {
            long milliseconds = DateTime.Now.Ticks / TimeSpan.TicksPerMillisecond;
            return milliseconds.ToString();
        }

        //create directory to storr excel files
        public string createDirectory()
        {
            try
            {
                string dirName = path + this.createFilename();
                if (!Directory.Exists(dirName))
                {
                    DirectoryInfo dir = Directory.CreateDirectory(dirName);
                    return dirName + "\\";
                }
                Environment.Exit(0);
                return null;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                Environment.Exit(0);
                return null;
            }
        }

        //convert date
        public string formatDate(string dateString)
        {
            return DateTime.ParseExact(dateString, "yyyyMMdd", null).ToString("dd-MM-yyyy");
        }

        //check if attachement exists
        public bool isAttachementExists(string filePath)
        {
            if (File.Exists(filePath))
            {
                return true;
            }
            return false;
        }

        //return the file directory
        public string getFileSaveDir()
        {
            return path;
        }

        //create excel file
        public bool createExcelFile(string fname, string location)
        {
            return false;
        }

        //remove excel object from the memory
        public void killProcess(List<int> pidList, string processName)
        { 
            // to kill current process of excel
            Process[] allProcesses = Process.GetProcessesByName(processName);
            foreach (int pid in pidList)
            {
                foreach ( Process process in allProcesses)
                { 
                    if (process.Id == pid)
                    {
                        try
                        {
                            process.Kill();
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.ToString());
                        }
                    }
                }
            }
            allProcesses = null;
        }

        //formate the number
        public string doFormat(decimal myNumber)
        {
            var s = string.Format("{0:0.00}", myNumber);

            if (s.EndsWith("00"))
            {
                return ((int)myNumber).ToString();
            }
            else
            {
                return s;
            }
        }

        //release object (clear from memory)
        public void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Console.WriteLine("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }

        }

    }
}
