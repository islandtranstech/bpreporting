using System;
using System.Collections.Generic;
using System.Text;
using System.Net;
using Microsoft.Office.Core;
using System.Net.Mail;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO; 

namespace AutomateBPReporting
{
    class Program
    {
        
        static List<List<object>> ExtractData(ISTCDataDataSet1.BillOfLadingDataTable table)
        {
            List<List<object>> data = new List<List<object>>();
            foreach (System.Data.DataRow row in table.Rows)
            {
                string billingCode = row["BillingCode"].ToString();
                string secondOrderNo = row["SecondOrderNo"].ToString();
                string orderNo = row["OrderNo"].ToString();
                string bolStatus = row["BOLStatus"].ToString();
                string bolId = row["BillOfLadingId"].ToString();
                DateTime? arrived = null;
                DateTime? left = null;
                DateTime? atCutomer = null;
                DateTime? finalized = null;
                DateTime? schedled = null;

                try
                {
                    if (row["ArrivedAtRackTime"] != null)
                    {
                        arrived = (DateTime)row["ArrivedAtRackTime"];
                    }
                }
                catch (Exception)
                {
                    
                }

                try
                {
                    if (row["LeftRackTime"] != null)
                    {
                        left = (DateTime)row["LeftRackTime"];
                    }
                }
                catch (Exception)
                {
                }
                try
                {
                    if (row["ArrivedAtCustomerTime"] != null)
                    {
                        atCutomer = (DateTime)row["ArrivedAtCustomerTime"];
                    }
                }
                catch (Exception)
                {
                }
                try
                {

                    if (row["FinalizedDeliveryTime"] != null)
                    {
                        finalized = (DateTime)row["FinalizedDeliveryTime"];
                    }
                }
                catch (Exception)
                {
                }
                try
                {

                    if (row["ScheduledDate"] != null)
                    {
                        schedled = (DateTime)row["ScheduledDate"];
                    }
                }
                catch (Exception)
                {
                }

                int rackTime = 0;
                if (left != null && arrived != null)
                {
                    TimeSpan? ts = left - arrived;
                    if (ts != null)
                    {
                        rackTime = (int) Math.Round(ts.Value.TotalMinutes);
                    }
                }

                int dropTime = 0;
                if (finalized != null && atCutomer != null)
                {
                    TimeSpan? ts = finalized - atCutomer;
                    if (ts != null)
                    {
                        dropTime = (int)Math.Round(ts.Value.TotalMinutes);
                    }
                }
  
                List<object> srow = new List<object>();
                srow.Add(bolStatus);
                
                srow.Add((arrived != null) ? arrived.Value.ToShortTimeString() : "");
                srow.Add((left != null) ? left.Value.ToShortTimeString() : "");
                srow.Add(rackTime);
                srow.Add((atCutomer != null) ? atCutomer.Value.ToShortTimeString() : "");
                srow.Add((finalized != null) ? finalized.Value.ToShortTimeString() : "");
                srow.Add(dropTime);
                srow.Add(secondOrderNo);
                srow.Add(orderNo);
                srow.Add(bolId);
                data.Add(srow);
            }
            return data;
        }

        static void Main(string[] args)
        {
            try
            {
                ExcelHelper eh = new ExcelHelper();

                // do babylon
                ISTCDataDataSet1TableAdapters.BillOfLadingTableAdapter bol = new ISTCDataDataSet1TableAdapters.BillOfLadingTableAdapter();

                DateTime today = DateTime.Now;
                DateTime reportDate = new DateTime(today.Year, today.Month, today.Day - 1);
                ISTCDataDataSet1.BillOfLadingDataTable table = bol.GetDataBPReport(25, reportDate);
                List<List<object>> writeRows = ExtractData(table);
                eh.WriteWorkSheet(writeRows, reportDate, "Babylon");
                
                // nj
                table = bol.GetDataBPReport(59, reportDate);
                writeRows = ExtractData(table);
                eh.WriteWorkSheet(writeRows, reportDate, "Jersey");

                // bk
                table = bol.GetDataBPReport(60, reportDate);
                writeRows = ExtractData(table);
                eh.WriteWorkSheet(writeRows, reportDate, "Brooklyn");


                string path = eh.SaveAs();
                Attachment attach = new Attachment(path);

                var fromAddress = new MailAddress("islandtranstech@gmail.com", "ITC Tech");
                var toAddress = new MailAddress("dfioretti@islandtrans.com", "Dave");
                const string fromPassword = "azaz09**";
                const string subject = "BP Email";
                const string body = "See Attached";


                TextReader sr = new StreamReader("d:/Reports/emails.txt");
                string emailTo = sr.ReadLine();

                var smtp = new SmtpClient
                {
                    Host = "smtp.gmail.com",
                    Port = 587,
                    EnableSsl = true,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    UseDefaultCredentials = false,
                    Credentials = new NetworkCredential(fromAddress.Address, fromPassword)
                };
                using (var message = new MailMessage(fromAddress, toAddress)
                {
                    Subject = subject,
                    Body = body

                })
                {
                    message.Attachments.Add(attach);
                    message.To.Add(emailTo);
                    smtp.Send(message);
                }
            }
            catch (Exception e)
            {
                TextWriter tw = new StreamWriter("D:/Reports/log.txt");

                // write a line of text to the file
                tw.WriteLine("ERRORS:");
                tw.WriteLine(e.Message + " " + e.StackTrace);

                tw.Close();
            }
        }
    }
}


/*
using System.Net;
using System.Net.Mail;

var fromAddress = new MailAddress("from@gmail.com", "From Name");
var toAddress = new MailAddress("to@example.com", "To Name");
const string fromPassword = "fromPassword";
const string subject = "Subject";
const string body = "Body";

var smtp = new SmtpClient
           {
               Host = "smtp.gmail.com",
               Port = 587,
               EnableSsl = true,
               DeliveryMethod = SmtpDeliveryMethod.Network,
               UseDefaultCredentials = false,
               Credentials = new NetworkCredential(fromAddress.Address, fromPassword)
           };
using (var message = new MailMessage(fromAddress, toAddress)
                     {
                         Subject = subject,
                         Body = body
                     })
{
    smtp.Send(message);
}
*/