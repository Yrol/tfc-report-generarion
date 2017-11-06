using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DueContractEmail
{
    class Connection
    {
        public string IMAS(string Collection)
        {
            //Collection = DC@LENLIB
            string DataSource = "XXX.XXX.X.X";
            string dbUserName = "XXXXXXX";
            string dbPassword = "XXXXXXX";
            string ConnectionString = "Provider=IBMDA400;Data Source=" + DataSource + ";User Id=" + dbUserName + ";Password=" + dbPassword + ";Default Collection=" + Collection + ";";
            return ConnectionString;
        }
    }
}
