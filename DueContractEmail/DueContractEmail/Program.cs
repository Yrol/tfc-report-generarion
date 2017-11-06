using System;
using System.Data;
using System.Threading;
using System.ComponentModel;
using System.Collections.Generic;

namespace DueContractEmail
{
    class Program
    {

        //background worker variables
        private static BackgroundWorker worker = new BackgroundWorker();
        private static AutoResetEvent resetEvent = new AutoResetEvent(false);

        static void Main(string[] args)
        {
            Utils utils = new Utils();

            //fetch data from DB
            DataTable dataTable = new DataTable();
            RetrieveInformation retrieveInformation = new RetrieveInformation();
            dataTable = retrieveInformation.retrieveInformation();

            if (dataTable.Rows.Count > 0)
            {
                //create directory
                string currentPath = utils.createDirectory();
                if (currentPath.Equals(null)) { Environment.Exit(0); }
                Console.WriteLine("Files directory : " + currentPath);
                /*
                new Thread(() =>
                {
                    Thread.CurrentThread.IsBackground = false;
                    BuildDueContractFiles bm = new BuildDueContractFiles(dataTable, currentPath);
                    Console.WriteLine(dataTable.Rows.Count);
                }).Start();
                */
                BuildDueContractFiles bm = new BuildDueContractFiles(dataTable, currentPath);
                Console.WriteLine(dataTable.Rows.Count);
            }
            else
            {
                Console.WriteLine("No records available");
            }

            //Console.WriteLine(retreiveInformation.getDefaultInterest("NIT", "LEM", "0000000335"));

            //run the background worker to clear excel objects
            worker.DoWork += worker_DoWork;
            worker.RunWorkerCompleted += worker_RunWorkerCompleted;
            worker.ProgressChanged += worker_ProgressChanged;
            worker.WorkerReportsProgress = true;
            worker.RunWorkerAsync(GlobalObjects.getPidList());
            resetEvent.WaitOne();

        }

        static void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            Console.WriteLine(e.ProgressPercentage.ToString());
        }

        static void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            Console.WriteLine("Clearing the excel objects from memory, please wait...");
            Utils utils = new Utils();
            List<int> pids = (List<int>)e.Argument;
            try
            {
                utils.killProcess(pids, "EXCEL");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        static void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Console.WriteLine("Task completed..." + Environment.ExitCode);
            //Environment.Exit(Environment.ExitCode);
            System.Diagnostics.Process.GetCurrentProcess().Kill ();
        }
    }
}
