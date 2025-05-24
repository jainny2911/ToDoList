using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Mail;
using ToDoList.Models;

namespace ToDoList.Controllers
{
    public class HomeController : Controller
    {
        private readonly IConfiguration _config;

        public HomeController(IConfiguration config)
        {
            _config = config;
        }

        [HttpGet]
        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public IActionResult SendEmail([FromBody] List<MailData> data)
        {
            try
            {
                if (data == null || data.Count == 0)
                    return BadRequest("No data received.");

                // Create Excel in memory
                using var workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("ToDo Data");

                worksheet.Cell(1, 1).Value = "ID";
                worksheet.Cell(1, 2).Value = "Description";
                worksheet.Cell(1, 3).Value = "Status";

                for (int i = 0; i < data.Count; i++)
                {
                    worksheet.Cell(i + 2, 1).Value = data[i].Id;
                    worksheet.Cell(i + 2, 2).Value = data[i].Description;
                    worksheet.Cell(i + 2, 3).Value = data[i].Status;
                }

                using var ms = new MemoryStream();
                workbook.SaveAs(ms);
                ms.Seek(0, SeekOrigin.Begin);
                var attachment = new Attachment(ms, "ToDoList.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

                // Get SMTP config from appsettings.json
                string host = _config["SmtpSettings:Host"];
                int port = int.Parse(_config["SmtpSettings:Port"]);
                bool enableSsl = bool.Parse(_config["SmtpSettings:EnableSsl"]);
                string username = _config["SmtpSettings:Username"];
                string password = _config["SmtpSettings:Password"];
                string from = _config["SmtpSettings:From"];
                string to = _config["SmtpSettings:To"];

                var smtpClient = new SmtpClient(host)
                {
                    Port = port,
                    Credentials = new NetworkCredential(username, password),
                    EnableSsl = enableSsl
                };

                var mail = new MailMessage
                {
                    From = new MailAddress(from),
                    Subject = "ToDo Tasks With Excel",
                    Body = "Please find the attached ToDo Excel file."
                };

                mail.To.Add(to);
                mail.Attachments.Add(attachment);

                smtpClient.Send(mail);
                return Ok("Email with Excel attachment sent.");
            }
            catch (Exception ex)
            {
                return BadRequest("Failed to send: " + ex.Message);
            }
        }
    }
}
