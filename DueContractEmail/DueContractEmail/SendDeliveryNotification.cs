using System;
using System.Collections.Generic;
using System.Net.Mime;
using System.Net;
using System.Linq;
using System.Net.Mail;

namespace DueContractEmail
{
    class SendDeliveryNotification
    {
        public SendDeliveryNotification(int emailCount, string fileType)
        {
            string current_date = DateTime.Now.ToString("dd/MM/yyyy");
            BranchData branchData = new BranchData();

            foreach (string emailAddress in branchData.getDeliveryNotifications())
            {
                try
                {
                    Console.WriteLine("Delivery email : " + emailAddress);
                    MailMessage message = new MailMessage("TFC Due Contracts Delivery Notification <tfcinfosys@thefinance.lk>", emailAddress, "Due contracts. ", "\nDear Sir /Madam,\n\n\nDelivery of " + fileType + " files completed with " + emailCount + " email(s) at " + current_date + ".\n\n\n\nBest regards,\n\nTFC Information Systems\n\nIT Division,\nThe Finance Company PLC,\nNo 55, Laurries Place,\nR.A.De Mel Mawatha,\nColombo -04\n\n\n0714053505-->Mr.Anuruddha Ranaweera (Head of IT)\n");
                    SmtpClient client = new SmtpClient("XXX.XXX.XXX.X");
                    client.Credentials = CredentialCache.DefaultNetworkCredentials;
                    client.Send(message);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.ToString());
                }
            }
        }
    }
}
