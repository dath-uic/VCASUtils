using VCASPdfUtil.Models;
using RazorEngine.Templating;
using SelectPdf;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace VCASPdfUtil.Processer
{
    public class EmailManager
    {
        private readonly string FormsTemplateFolderPath = string.Format("{0}\\FormTemplates", Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName);
        public string SendEmailWithAttachment<T>(string attachmentTemplate, object obj, EmailMetadataModel emailModel)
        {
            //Prepare attachment before delegating call to SendEmail()
            if (emailModel.HasAttachment)
            {
                T customModel = (T)obj;
                byte[] pdf = null;

                var templateService = new TemplateService();
                var formHtml = templateService.Parse(System.IO.File.ReadAllText(Path.Combine(FormsTemplateFolderPath, attachmentTemplate)), customModel, null, null);

                HtmlToPdf converter = new HtmlToPdf();
                SelectPdf.GlobalProperties.LicenseKey = System.Configuration.ConfigurationSettings.AppSettings.Get("SelectPdfLicenseKey");
                // set converter options
                converter.Options.MarginLeft = 0;
                converter.Options.MarginTop = 10;
                converter.Options.MarginBottom = 10;
                converter.Options.MarginRight = 0;
                converter.Options.WebPageWidth = 780;
                converter.Options.AutoFitWidth = HtmlToPdfPageFitMode.NoAdjustment;

                // create a new pdf document converting an url
                PdfDocument doc = converter.ConvertHtmlString(formHtml);

                // save pdf document
                pdf = doc.Save();

                // close pdf document
                doc.Close();

                Attachment emailAttachment = new Attachment(new MemoryStream(pdf), emailModel.AttachmentNameWithoutExtn + ".pdf", "application/pdf");
                emailModel.EmailAttachment = emailAttachment;
            }
            string emailException = SendEmail(emailModel);
            return emailException;
        }
        public string SendEmail(EmailMetadataModel emailMetadata)
        {
            string isEnableSendEmail = System.Configuration.ConfigurationSettings.AppSettings.Get("IsEnableSendEmail");
            bool failed = false;
            string emailException = null;

            if (!string.IsNullOrEmpty(isEnableSendEmail) && isEnableSendEmail.ToLower().Equals("true"))
            {
                string smtpServerAddress = ConfigurationSettings.AppSettings.Get("SMTPServer");
                var systemAdminEmails = ConfigurationSettings.AppSettings.Get("SystemAdminEmail");

                SmtpClient client = new SmtpClient(smtpServerAddress, 25);
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.EnableSsl = false;
                client.UseDefaultCredentials = false;

                int tryAgain = Convert.ToInt32(ConfigurationSettings.AppSettings.Get("EmailRetryAttemptsCount"));

                string emailSendWaitInterval = ConfigurationSettings.AppSettings.Get("EmailSendInterval");
                int numEmailSendInterval = Convert.ToInt32(emailSendWaitInterval);

                List<string> tempList = new List<string>();
                do
                {
                    try
                    {
                        failed = false;

                        MailMessage message = new MailMessage();
                        message.From = new MailAddress(emailMetadata.From);

                        while (emailMetadata.To.Count > 0)
                        {
                            message.To.Clear();
                            message.Bcc.Clear();
                            message.To.Add(new MailAddress(emailMetadata.To[0]));

                            if (emailMetadata.EnableAdminBcc)
                            {
                                if (!string.IsNullOrEmpty(systemAdminEmails))
                                {
                                    foreach (var address in systemAdminEmails.Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries))
                                        message.Bcc.Add(new MailAddress(address));
                                }
                            }

                            if (emailMetadata.Bcc != null)
                            {
                                foreach (var address in emailMetadata.Bcc)
                                    message.Bcc.Add(new MailAddress(address));
                            }

                            tempList.Add(emailMetadata.To[0]);

                            emailMetadata.To.RemoveAt(0);
                            var emailHtmlBody = emailMetadata.EmailBody;

                            message.IsBodyHtml = true;
                            message.Subject = emailMetadata.Subject;
                            message.Body = emailHtmlBody;

                            if (emailMetadata.HasAttachment)
                                message.Attachments.Add(emailMetadata.EmailAttachment);

                            client.Send(message);
                            tryAgain = Convert.ToInt32(ConfigurationSettings.AppSettings.Get("EmailRetryAttemptsCount"));
                            System.Threading.Thread.Sleep(numEmailSendInterval);
                            tempList.Clear();
                            emailException = null;
                        }
                    }
                    catch (Exception ex)
                    {
                        failed = true;
                        tryAgain--;
                        emailMetadata.To.AddRange(tempList);
                        tempList.Clear();
                        emailException = ex.Message.ToString();
                        System.Threading.Thread.Sleep(numEmailSendInterval);
                    }
                } while (failed && tryAgain != 0);
            }

            return emailException;
        }
    }
}
