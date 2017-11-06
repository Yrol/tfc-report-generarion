using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DueContractEmail
{
    static class GlobalObjects
    {
        static List<int> pidList;

        static GlobalObjects()
        {
            pidList = new List<int>();
        }

        public static void addPidsToList(int pid)
        {
            pidList.Add(pid);
        }

        public static List<int> getPidList()
        {
            return pidList;
        }
    }
}
