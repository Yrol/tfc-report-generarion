using System;
using System.Collections.Generic;
using System.Net.Mime;
using System.Net;
using System.Linq;
using System.Net.Mail;
using System.ComponentModel;
using System.Threading;

namespace DueContractEmail
{
    class SendEmail
    {
        public SendEmail(Dictionary<string, string> fileEmailCollection, string currentDirectory)
        {
            Console.WriteLine("Sending emails  please wait...");
            Utils utils = new Utils();
            string current_date = DateTime.Now.ToString("dd/MMMM/yyyy");

            for (int count = 0; count < fileEmailCollection.Count; count++)
            {
                //use this to send just one email, for testing only
                /*
                if (count > 0)
                {
                    break;
                }
                */

                var element = fileEmailCollection.ElementAt(count);
                var fileName = currentDirectory + element.Key + ".xlsx";
                var branchManagerEmail = element.Value;

                Console.WriteLine("BM:" + branchManagerEmail + " : " + fileName.ToString());
                if (utils.isAttachementExists(fileName))
                {
                    try
                    {
                        MailMessage message = new MailMessage("TFC Information Systems - Due Contracts <tfcinfosys@thefinance.lk>", branchManagerEmail, "Due contracts as of " + current_date, "\nDear Sir /Madam,\n\n\nHere with attached due contracts of your branch as at " + current_date + ".\n\n\n\nBest regards,\n\nTFC Information Systems\n\nIT Division,\nThe Finance Company PLC,\nNo 55, Laurries Place,\nR.A.De Mel Mawatha,\nColombo -04\n\n\n0714053505-->Mr.Anuruddha Ranaweera (Head of IT)\n");
                        // Create  the file attachment for this e-mail message.
                        Attachment data = new Attachment(fileName, MediaTypeNames.Application.Octet);
                        // Add time stamp information for the file.
                        ContentDisposition disposition = data.ContentDisposition;
                        disposition.CreationDate = System.IO.File.GetCreationTime(fileName);
                        disposition.ModificationDate = System.IO.File.GetLastWriteTime(fileName);
                        disposition.ReadDate = System.IO.File.GetLastAccessTime(fileName);
                        // Add the file attachment to this e-mail message.
                        message.Attachments.Add(data);

                        //Send the message.
                        SmtpClient client = new SmtpClient("XXX.XXX.XX.X");
                        // Add credentials if the SMTP server requires them.
                        client.Credentials = CredentialCache.DefaultNetworkCredentials;
                        //client.Send(message);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("Mail send exception: " + e.ToString());
                    }
                }
            }
            //send delivery notifications
            SendDeliveryNotification sendDeliveryNotifications = new SendDeliveryNotification(fileEmailCollection.Count, "Branch Managers");

            Console.WriteLine("Email send completed");//complete send

        }
    }
}
