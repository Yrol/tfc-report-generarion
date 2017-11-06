using System;
using System.Data.OleDb;
using System.Data;
using System.Collections.Generic;

namespace DueContractEmail
{
    class RetrieveInformation
    {
        private DataTable table;
        private string todayDate;

        private string LEPRT_CTCODE_branch = null;
        private string LEPRT_PRODCT_product = null;
        private string LEPRT_CNTRNO_contractno = null;
        private decimal LEPRT_RENTAL_rentalamount = 0.0m;
        private decimal LEPRT_RENSNO_rentalno = 0.0m;

        private string XEPCS_PERIOD_period = null;
        private decimal XEPCS_OPNBAL_debtorbalance = 0.0m;
        private decimal XEPCS_NORNAR_arrears = 0.0m;
        private string XEPCS_ADDRES_address = null;
        private string XEPCS_MONAME_officername = null;
        private string XEPCS_NAMECH_customername = null;
        private string XEPCS_DCNAME_collector_name = null;

        private decimal LEPDI_default_interest = 0.0m;

        private decimal newDebtorBalance = 0.0m;
        private decimal daysArrears = 0.0m;
        private decimal defaultInterest = 0.0m;

        public DataTable retrieveInformation()
        {
            Utils utils = new Utils();
            todayDate = getCurrentDate();
            table = new DataTable();
            table.Columns.Add("Branch", typeof(string));
            table.Columns.Add("Product", typeof(string));
            table.Columns.Add("Contract No", typeof(string));
            table.Columns.Add("Rental Amount", typeof(string));
            table.Columns.Add("Rental No", typeof(string));
            table.Columns.Add("Contract peroid", typeof(string));
            table.Columns.Add("Debtor balance", typeof(string));
            table.Columns.Add("Default Interest", typeof(string));
            table.Columns.Add("No of days Over in arrears", typeof(string));
            table.Columns.Add("Customer name", typeof(string));
            table.Columns.Add("Address", typeof(string));
            table.Columns.Add("Marketing officer", typeof(string));
            table.Columns.Add("Collector name", typeof(string));

            Connection conn = new Connection();

            try
            {
                Console.WriteLine("Fetching data from the DB, please wait.......");
                OleDbConnection oledbConnection = new OleDbConnection();
                string SELECT = "SELECT " +
                                "LEPRT.CTCODE, " +
                                "LEPRT.PRODCT, " +
                                "LEPRT.CNTRNO, " +
                                "LEPRT.RENTAL, " +
                                "LEPRT.RENSNO, " +
                                "XEPCS.PERIOD, " +
                                "XEPCS.CLOBAL, " +
                                "XEPCS.NORNAR, " +
                                "XEPCS.NAMECH, " +
                                "XEPCS.ADDRES, " +
                                "XEPCS.MONAME, " +
                                "XEPCS.DCNAME " + 
                                "FROM LEPRT LEFT JOIN XEPCS ON LEPRT.CTCODE = XEPCS.CTCODE AND LEPRT.PRODCT = XEPCS.PRODCT AND LEPRT.CNTRNO = XEPCS.CNTRNO " +  
                                "WHERE LEPRT.RENDDT = " + todayDate + " AND (LEPRT.STATUS = 'A' OR LEPRT.STATUS = 'E')";

                oledbConnection.ConnectionString = conn.IMAS("DC@LENLIB");
                OleDbCommand myOledbCommand = new OleDbCommand(SELECT, oledbConnection);
                myOledbCommand.Connection.Open();
                OleDbDataReader myOledbDataReader = myOledbCommand.ExecuteReader();

                if (myOledbDataReader.HasRows && myOledbDataReader.FieldCount > 0)
                {
                    while (myOledbDataReader.Read())
                    {
                        //LEPRT table
                        LEPRT_CTCODE_branch = myOledbDataReader.GetString(0);
                        LEPRT_PRODCT_product = myOledbDataReader.GetString(1);
                        LEPRT_CNTRNO_contractno = myOledbDataReader.GetString(2);
                        LEPRT_RENTAL_rentalamount = myOledbDataReader.GetDecimal(3);
                        LEPRT_RENSNO_rentalno = myOledbDataReader.GetDecimal(4);

                        //XEPCS table
                        XEPCS_PERIOD_period = myOledbDataReader.GetValue(5).ToString();
                        XEPCS_OPNBAL_debtorbalance = myOledbDataReader.GetDecimal(6);

                        //get the default interest
                        //LEPDI_default_interest = getDefaultInterest(LEPRT_CTCODE_branch, LEPRT_PRODCT_product, LEPRT_CNTRNO_contractno);

                        //XEPCS table contunues...... 
                        XEPCS_NORNAR_arrears = myOledbDataReader.GetDecimal(7);
                        XEPCS_NAMECH_customername = myOledbDataReader.GetString(10);
                        XEPCS_ADDRES_address = myOledbDataReader.GetString(8);
                        XEPCS_MONAME_officername = myOledbDataReader.GetString(9);
                        XEPCS_DCNAME_collector_name = myOledbDataReader.GetString(11);

                        //calculate debtor balance and days of arrears
                        newDebtorBalance = XEPCS_OPNBAL_debtorbalance + LEPRT_RENTAL_rentalamount;
                        daysArrears = newDebtorBalance / LEPRT_RENTAL_rentalamount;

                        //formatted date
                        var formattedDaysArrears = utils.doFormat(daysArrears);
                        var formattedDebtorBalance = utils.doFormat(newDebtorBalance);

                        //add data to the table collection
                        table.Rows.Add(
                                       LEPRT_CTCODE_branch.ToString(),
                                       LEPRT_PRODCT_product.ToString(),
                                       LEPRT_CNTRNO_contractno.ToString(),
                                       LEPRT_RENTAL_rentalamount.ToString(),
                                       LEPRT_RENSNO_rentalno.ToString(),
                                       XEPCS_PERIOD_period.ToString(),
                                       formattedDebtorBalance.ToString(),
                                       LEPDI_default_interest.ToString(),
                                       formattedDaysArrears.ToString(),
                                       XEPCS_ADDRES_address.ToString(),
                                       XEPCS_MONAME_officername.ToString(),
                                       XEPCS_NAMECH_customername.ToString(),
                                       XEPCS_DCNAME_collector_name.ToString()
                                       );

                        LEPRT_CTCODE_branch = null;
                        LEPRT_PRODCT_product = null;
                        LEPRT_CNTRNO_contractno = null;
                        LEPRT_RENTAL_rentalamount = 0.0m;
                        LEPRT_RENSNO_rentalno = 0.0m;

                        XEPCS_PERIOD_period = null;
                        XEPCS_OPNBAL_debtorbalance = 0.0m;
                        XEPCS_NORNAR_arrears = 0.0m;
                        XEPCS_ADDRES_address = null;
                        XEPCS_MONAME_officername = null;
                        XEPCS_NAMECH_customername = null;
                        XEPCS_DCNAME_collector_name = null;

                        LEPDI_default_interest = 0.0m;

                        newDebtorBalance = 0.0m;
                        daysArrears = 0.0m;
                        defaultInterest = 0.0m;
                    }
                    myOledbCommand.Connection.Close();
                    myOledbCommand.Dispose();
                    Console.WriteLine("Data fetch completed with " + table.Rows.Count + " row(s)");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            return table;
        }

        //get default interest rate
        private decimal getDefaultInterest(string branchCode, string productCode, string contractNo)
        {
            defaultInterest = 0.0m;
            List<decimal> defualtInterstCollection = new List<decimal>();
            Connection conn = new Connection();
            try
            {
                OleDbConnection oledbConnection = new OleDbConnection();

                string SELECT = "SELECT " +
                                "DEFAMT, " +
                                "DEFPID, " +
                                "WAIAMT " +
                                "FROM LEPDI WHERE CTCODE ='" + branchCode + "' AND PRODCT ='" + productCode + "' AND CNTRNO='" + contractNo +"' AND STATUS = 'A'";

                oledbConnection.ConnectionString = conn.IMAS("DC@LENLIB");
                OleDbCommand myOledbCommand = new OleDbCommand(SELECT, oledbConnection);
                myOledbCommand.Connection.Open();
                OleDbDataReader myOledbDataReader = myOledbCommand.ExecuteReader();

                if (myOledbDataReader.HasRows && myOledbDataReader.FieldCount > 0)
                {
                    while (myOledbDataReader.Read())
                    {
                        decimal tempDefaultValue = myOledbDataReader.GetDecimal(0) - myOledbDataReader.GetDecimal(1) - myOledbDataReader.GetDecimal(2);
                        defualtInterstCollection.Add(tempDefaultValue);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }

            //add all the values at the end
            if (defualtInterstCollection.Count > 0)
            {
                foreach (decimal val in defualtInterstCollection)
                {
                    defaultInterest = defaultInterest + val;
                }
            } 
            return defaultInterest;
        }

        //return current date to match DB query
        private string getCurrentDate()
        {
            return DateTime.Now.ToString("yyyyMMdd");
        }
    }
}
