using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DueContractEmail
{
    class BranchData
    {
        private Dictionary<string, string> branchManagers;
        private List<string> deliveryNotifications;

        public BranchData()
        {
            //add branch managers
            branchManagers = new Dictionary<string, string>();
            branchManagers.Add("AWG", "xxxxxxx@xxxxxxxxx.lk");
            branchManagers.Add("AMB", "xxxxxxx@xxxxxxxxx.lk");
            branchManagers.Add("AMP", "xxxxxxx@xxxxxxxxx.lk");
            branchManagers.Add("ANU", "xxxxxxx@xxxxxxxxx.lk");
            branchManagers.Add("AVI", "xxxxxxx@xxxxxxxxx.lk");
            branchManagers.Add("BAD", "xxxxxxx@xxxxxxxxx.lk");
            branchManagers.Add("BDR", "xxxxxxx@xxxxxxxxx.lk");
            branchManagers.Add("BAL", "xxxxxxx@xxxxxxxxx.lk");
            branchManagers.Add("BND", "xxxxxxx@xxxxxxxxx.lk");
            branchManagers.Add("BAT", "xxxxxxx@xxxxxxxxx.lk");
            branchManagers.Add("CHV", "xxxxxxx@xxxxxxxxx.lk");
            branchManagers.Add("CHI", "xxxxxxx@xxxxxxxxx.lk");
            branchManagers.Add("DAB", "xxxxxxx@xxxxxxxxx.lk");
            branchManagers.Add("NLW", "xxxxxxx@xxxxxxxxx.lk");
            branchManagers.Add("ELP", "xxxxxxx@xxxxxxxxx.lk");
            branchManagers.Add("EMB", "xxxxxxx@xxxxxxxxx.lk");
            branchManagers.Add("GLW", "xxxxxxx@xxxxxxxxx.lk");
            branchManagers.Add("GAL", "xxxxxxx@xxxxxxxxx.lk");
            branchManagers.Add("GAM", "xxxxxxx@xxxxxxxxx.lk");
            branchManagers.Add("HAT", "xxxxxxx@xxxxxxxxx.lk");
            branchManagers.Add("HIG", "xxxxxxx@xxxxxxxxx.lk");
            branchManagers.Add("HOM", "xxxxxxx@xxxxxxxxx.lk");
            branchManagers.Add("HOR", "xxxxxxx@xxxxxxxxx.lk");
            branchManagers.Add("RED", "xxxxxxx@xxxxxxxxx.lk");
            branchManagers.Add("JAE", "xxxxxxx@xxxxxxxxx.lk");
            branchManagers.Add("JFF", "xxxxxxx@xxxxxxxxx.lk");
            branchManagers.Add("KAD", "xxxxxxx@xxxxxxxxx.lk");
            branchManagers.Add("KAL", "xxxxxxx@xxxxxxxxx.lk");
            branchManagers.Add("KAN", "xxxxxxx@xxxxxxxxx.lk");


            //Delivery notification emails
            deliveryNotifications = new List<string>();
            deliveryNotifications.Add("yrolf@thefinance.lk");
        }

        public Dictionary<string, string> getBranchManagers()
        {
            return branchManagers;
        }

        public List<string> getDeliveryNotifications()
        {
            return deliveryNotifications;
        }
    }
}
