using System;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace Newsletter.Controllers
{
    public class HomeController : Controller
    {
        [ValidateInput(false)]
        [HttpPost]
        public async Task<ActionResult> SendNewsletter(HttpPostedFileBase recipient, string subject, string mailMessage, HttpPostedFileBase[] fileAttachments = null)
        {
            try
            {
                //SendEmail(subject, mailMessage, recipient, fileAttachments);
                ViewBag.Message = "Mail sent to recipients";

                // begin validations
                if (recipient == null)
                    throw new Exception("Recipient file must be uploaded");
                if (string.IsNullOrWhiteSpace(subject))
                    throw new Exception("Subject is required");
                if (string.IsNullOrWhiteSpace(mailMessage))
                    throw new Exception("Mail body is required");

                DataTable recipients = LoadExcelFile("recipient", "Recipients");
                if (recipients.Rows.Count == 0)
                    throw new Exception("At least one recipient is required.");
                if (!recipients.Columns.Contains("Email"))
                    throw new Exception("Column Email is required in the recipients file uploaded.");
                if (!recipients.Columns.Contains("Title"))
                    throw new Exception("Column Title is required in the recipients file uploaded.");
                if (!recipients.Columns.Contains("Name"))
                    throw new Exception("Column Name is required in the recipients file uploaded.");
                // end validations

                await DispatchEmailsAsync(recipients, subject, mailMessage, fileAttachments);

                return View();
            }
            catch (Exception ex)
            {
                ViewBag.Message = "Error: " + ex.Message;
                return View();
            }
        }

        private async Task<bool> DispatchEmailsAsync(DataTable recipients, string subject, string mailMessage, HttpPostedFileBase[] fileAttachments = null)
        {
            int noOfEmailsProcessed = 0, noOfEmailsPerHour = int.Parse(ConfigurationManager.AppSettings["NoOfEmailsPerHour"].ToString());
            DateTime dt = DateTime.Now;
            foreach (DataRow recipient in recipients.Rows)
            {
                if (IEIA.CommonUtil.RegexUtilities.IsValidEmail(recipient["Email"].ToString().Trim()))
                {
                    mailMessage = mailMessage.Replace("{0}", recipient["Title"].ToString().Trim()).Replace("{1}", recipient["Name"].ToString().Trim());
                    await SendEmailAsync(subject, mailMessage, recipient["Email"].ToString().Trim(), fileAttachments);

                    noOfEmailsProcessed++;
                }

                TimeSpan span = DateTime.Now.Subtract(dt);
                if (span.Minutes <= 60 && noOfEmailsProcessed >= noOfEmailsPerHour)
                {
                    System.Threading.Thread.Sleep(60 - span.Minutes);   // wait for the remaining time to 1 hour
                                                                        // reset the time and no of e-mails sent
                    dt = DateTime.Now;
                    noOfEmailsProcessed = 0;
                }
            }

            return true;
        }

        public ActionResult Index()
        {
            ViewBag.Message = string.Empty;
            return View();
        }

        public async Task<bool> SendEmailAsync(string subject, string mailBody, string recipients, HttpPostedFileBase[] fileAttachments = null)
        {
            try
            {
                SmtpClient smtpClient = new SmtpClient(ConfigurationManager.AppSettings["SmtpHost"], int.Parse(ConfigurationManager.AppSettings["SmtpPort"]));

                smtpClient.Credentials = new NetworkCredential(ConfigurationManager.AppSettings["EmailSenderUserName"], ConfigurationManager.AppSettings["EmailSenderPassword"]);
                smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtpClient.EnableSsl = bool.Parse(ConfigurationManager.AppSettings["SmtpSSL"]);
                MailMessage mail = new MailMessage();

                mail.From = new MailAddress(ConfigurationManager.AppSettings["EmailSenderEmail"], ConfigurationManager.AppSettings["EmailSenderName"]);

                foreach (string email in recipients.Split(new char[] { ',' }))
                {
                    mail.To.Add(new MailAddress(email));
                }

                mailBody = mailBody +
                            "<p><strong>Thank you for choosing IEI-ANCHOR PENSIONS.</strong></p>" +
                            "<img src=\"https://ieianchorpensions.com/images/logo2.png\" alt=\"IEIA-Logo\" />" +
                            "<p>https://twitter.com/ieiapensionmgrs<br/>" +
                                "https://www.facebook.com/anchorpensions/<br/>" +
                                "https://www.instagram.com/ieianchorpensions/<br/>" +
                                "https://plus.google.com/110719286991854279316/<br/>" +
                                "https://www.youtube.com/channel/UCyXZnwdv_jcXyNeFfBfhvEw<br/>" +
                                "https://www.linkedin.com/company-beta/16194771" +
                            "</p>" +
                            "<p>" +
                                "<strong>Disclaimer Notice</strong><br>" +
                                "The information contained in or accompanying this e - mail is intended for the use of the stated recipient and may contain information that is confidential and / or privileged.If the reader is not the intended recipient or the agent thereof, you are hereby notified that any dissemination or distribution of this e - mail is strictly prohibited. If you have received this email in error, please notify us immediately.Any views or options presented are solely those of the author and do not necessarily represent those of IEI - ANCHOR PENSION MANAGERS LIMITED" +
                            "</p>";

                if (fileAttachments != null)
                {
                    if (fileAttachments.Length > 0)
                    {
                        if (fileAttachments[0] != null)
                        {
                            foreach (var file in fileAttachments)
                            {
                                string fileName = Path.Combine(Server.MapPath("~/Content/Uploads/Newsletters"), file.FileName);
                                string fileExtension = Path.GetExtension(file.FileName);
                                if (System.IO.File.Exists(fileName))
                                {
                                    try
                                    {
                                        System.IO.File.Delete(fileName);
                                    }
                                    catch (Exception ex)
                                    {
                                        fileName = fileName.Replace(fileExtension, "_" + fileExtension);
                                        string exx = ex.Message;
                                    }
                                }

                                saveFile:
                                try
                                {
                                    file.SaveAs(fileName);
                                }
                                catch (Exception ex)
                                {
                                    if (ex.Message != null)
                                    {
                                        while (ex.Message.Contains("because it is being used by another process"))
                                        {
                                            fileName = fileName.Replace(fileExtension, "_" + fileExtension);
                                            goto saveFile;
                                        }
                                    }
                                }
                                file.SaveAs(fileName);
                                Attachment myAttachment = new Attachment(fileName);
                                mail.Attachments.Add(myAttachment);
                            }
                        }
                    }
                }

                mail.Subject = subject;
                mail.Body = mailBody;
                mail.IsBodyHtml = true;

                await smtpClient.SendMailAsync(mail);

                return true;
            }
            catch (Exception ex)
            {
                HomeController.LogError(ex, HttpContext.Server.MapPath("~/Error_Log.txt"));
                throw ex;
            }
        }

        public DataTable LoadExcelFile(string fileControlID, string returnedDataTableName)
        {
            DataTable table = new DataTable(returnedDataTableName);

            if (Request.Files[fileControlID].ContentLength > 0)
            {
                string fileExtension = Path.GetExtension(Request.Files[fileControlID].FileName);

                if (fileExtension == ".xls" || fileExtension == ".xlsx")
                {
                    string fileLocation = Path.Combine(Server.MapPath("~/Content/Uploads/Recipients"), Request.Files[fileControlID].FileName);
                    if (System.IO.File.Exists(fileLocation))
                    {
                        try
                        {
                            System.IO.File.Delete(fileLocation);
                        }
                        catch (Exception ex)
                        {
                            fileLocation = fileLocation.Replace(fileExtension, "_" + fileExtension);
                            string exx = ex.Message;
                        }
                    }

                    saveFile:
                    try
                    {
                        Request.Files[fileControlID].SaveAs(fileLocation);
                    }
                    catch (Exception ex)
                    {
                        if (ex.Message != null)
                        {
                            while (ex.Message.Contains("because it is being used by another process"))
                            {
                                fileLocation = fileLocation.Replace(fileExtension, "_" + fileExtension);
                                goto saveFile;
                            }
                        }
                    }

                    string excelConnectionString = string.Empty;

                    //connection String for xls file format.
                    if (fileExtension == ".xls")
                    {
                        excelConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileLocation + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
                    }
                    //connection String for xlsx file format.
                    else if (fileExtension == ".xlsx")
                    {
                        excelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileLocation + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
                    }

                    //Create Connection to Excel work book and add oledb namespace
                    OleDbConnection excelConnection = new OleDbConnection(excelConnectionString);
                    excelConnection.Open();
                    DataTable dt = new DataTable();

                    dt = excelConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    if (dt == null)
                    {
                        return null;
                    }

                    String[] excelSheets = new String[dt.Rows.Count];
                    int t = 0;

                    //excel data saves in temp file here.
                    foreach (DataRow row in dt.Rows)
                    {
                        excelSheets[t] = row["TABLE_NAME"].ToString();
                        t++;
                    }
                    OleDbConnection excelConnection1 = new OleDbConnection(excelConnectionString);

                    string query = string.Format("Select * from [{0}]", excelSheets[0]);
                    using (OleDbDataAdapter dataAdapter = new OleDbDataAdapter(query, excelConnection1))
                    {
                        dataAdapter.Fill(table);
                    }
                }
            }

            return table;
        }

        public static void LogError(Exception err, string path)
        {
            try
            {
                #region error details
                HttpException error = err as HttpException;
                Exception lastError = error;

                if (lastError.InnerException != null)
                    lastError = lastError.InnerException;

                string errorDateTime = string.Format("{0:dd/MM/yyyy hh:mm:ss}", DateTime.Now);
                string errorType = lastError.GetType().ToString();
                string errorMessage = lastError.Message;
                string errorStackTrace = lastError.StackTrace;
                string errorSource = lastError.Source;
                string errorMethod = lastError.TargetSite.Name;

                StringBuilder exception = new StringBuilder();

                exception.AppendLine(string.Format("Date/Time: {0}", errorDateTime));
                exception.AppendLine(string.Format("Error Type: {0}", errorType));
                exception.AppendLine(string.Format("Methord: {0}", errorMethod));
                exception.AppendLine(string.Format("Source: {0}", errorSource));
                exception.AppendLine(string.Format("Message: {0}", errorMessage));
                exception.AppendLine(string.Format("Stack Trace: {0}", errorStackTrace));
                #endregion

                #region save the error in a file
                using (FileStream fs = new FileStream(path, FileMode.Append, FileAccess.ReadWrite, FileShare.ReadWrite))
                {
                    StreamWriter sr = new StreamWriter(fs);

                    sr.WriteLine(exception);
                }
                #endregion                
            }
            catch (Exception ex)
            {
                string error = ex.Message;
            }
        }
    }
}