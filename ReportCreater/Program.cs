using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net.Mail;
using System.Net;
using System.IO;
using Stimulsoft.Report;
using Stimulsoft.Report.Components;
using Stimulsoft.Base.Drawing;

namespace ReportCreater
{
    class Program
    {
        static void Main(string[] args)
        {
        }

        public static void CreateAndSendReport(string email)
        {
            //create report
            const string ReportsDirectory = "../../../Reports";

            var Report = new StiReport();
            var page = Report.Pages[0];
            var dataBand = new StiDataBand();
            dataBand.Height = 0.5;
            dataBand.Name = "DataBand";
            page.Components.Add(dataBand);
            dataBand.Components.Add(new StiText(new RectangleD(0, 0, 5, 0.5))
            {
                Text = DateTime.Now.ToString()
            });
            dataBand.Components.Add(new StiText(new RectangleD(3, 0, 5, 0.5))
            {
                Text = "Makishev E A"
            });

            Report.Render();

            var exportFilePath = $"{ReportsDirectory}/SP_Report{DateTime.Now.ToString("yyyy-dd-MM_HH-mm-ss")}.xls";
            Report.Render(false);
            Report.ExportDocument(StiExportFormat.Excel, exportFilePath);

            // send report
            try
            {
                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient("smtp.mail.ru");

                mail.From = new MailAddress("1esim1@mail.ru");
                mail.To.Add(email);
                mail.Subject = "Report";
                mail.Body = $"Hello i sended my first Stmulsoft report";
                Attachment reportFile = new Attachment(exportFilePath);
                mail.Attachments.Add(reportFile);

                SmtpServer.Port = 587;
                SmtpServer.Credentials = new System.Net.NetworkCredential("1esim1@mail.ru", "aidana05012008");
                SmtpServer.EnableSsl = true;

                SmtpServer.Send(mail);
            }
            catch
            {
                Console.WriteLine("Mail don't sended!");
            }
        }
    }
}

