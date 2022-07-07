using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Web;
using SelectPdf;
using Box.V2;
using Box.V2.Config;
using Box.V2.Auth;
using Box.V2.Models;
using Box.V2.Exceptions;
using System.Collections.Concurrent;
using Box.V2.JWTAuth;
using VCASPdfUtil.Processer;
using VCASPdfUtil.Models;
using System.Configuration;
using RazorEngine.Templating;

namespace VCASPdfUtil
{
    class Program
    {
        public static readonly string FormsTemplateFolderPath = string.Format("{0}\\FormTemplates", Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName);
        static void Main(string[] args)
        {

            try
            {
                CombinePDFUsingSelectPDF();
                SendCompletionEmail();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex);
            }

        }


        //selectPDF
        public static void CombinePDFUsingSelectPDF()
        {
            SelectPdf.GlobalProperties.LicenseKey = System.Configuration.ConfigurationSettings.AppSettings.Get("SelectPdfLicenseKey");
            Console.WriteLine("--->>" + "Begin Combining Pdf's");
            DateTime dt = DateTime.Now;
            string currentMonth = dt.Month.ToString();
            string currentYear = dt.Year.ToString();

            // create a new pdf document
            PdfDocument doc = new PdfDocument();
            string buildingStandardsLocation = System.Configuration.ConfigurationSettings.AppSettings.Get("BuildingStandards");
            DirectoryInfo bldgStandards = new DirectoryInfo(buildingStandardsLocation);

            DirectoryInfo[] folderArray = bldgStandards.GetDirectories().OrderBy(m => m.Name).ToArray();
            List<FileInfo> combinedFiles = new List<FileInfo>();
            foreach (DirectoryInfo folder in folderArray)
            {
                Console.WriteLine("--->>" + folder);
                doc = new PdfDocument();
                FileInfo[] filePaths = folder.GetFiles("*.pdf").OrderBy(m => m.Name).ToArray();
                PdfDocument doc1 = new PdfDocument();
                foreach (FileInfo file in filePaths)
                {
                    doc1 = new PdfDocument(file.FullName);
                    doc.Append(doc1);

                    //doc1.Close();
                }

                doc.Save(folder.FullName + "\\" + folder +"_MERGED_"+ DateTime.Now.ToString("dd") + DateTime.Now.ToString("yyyy") + ".pdf");
                //<folder_name>_MERGED_072022

            }

            PdfDocument combinedDoc = new PdfDocument();
            FileInfo[] filePathsInBldsFolder = bldgStandards.GetFiles("*.pdf").OrderBy(m => m.Name).ToArray();

            foreach (FileInfo file in filePathsInBldsFolder)
            {
                PdfDocument combinedDoc1 = new PdfDocument(file.FullName);
                combinedDoc.Append(combinedDoc1);
            }

            foreach (DirectoryInfo folder in folderArray)
            {
                FileInfo[] filePaths = folder.GetFiles(folder + "_MERGED_"+ DateTime.Now.ToString("dd") + DateTime.Now.ToString("yyyy") + ".pdf");

                combinedFiles.Add(filePaths[0]);
            }

            foreach (FileInfo file in combinedFiles)
            {
                PdfDocument combinedDoc1 = new PdfDocument(file.FullName);
                combinedDoc.Append(combinedDoc1);
            }

            combinedDoc.Save(buildingStandardsLocation + "\\" + "BUILDING_STANDARDS_MERGED_"+ DateTime.Now.ToString("dd") + DateTime.Now.ToString("yyyy") + ".pdf");
            //BUILDING_STANDARDS_MERGED_072022.pdf

            Console.WriteLine("{DONE}");
        }

        public static void SendCompletionEmail()
        {
            EmailManager emailManager = new EmailManager();
            NotificationModel notificationRecord = new NotificationModel();

            EmailMetadataModel emailMetadata = new EmailMetadataModel();
            string toAddress = ConfigurationSettings.AppSettings.Get("ToEmailAddress");
            emailMetadata.To = toAddress.Split(';').ToList();
            emailMetadata.From = ConfigurationSettings.AppSettings.Get("FromEmailAddress");
            emailMetadata.Subject = ConfigurationSettings.AppSettings.Get("CombiningPdfCompletedEmailSubject");
            emailMetadata.Subject +=  DateTime.Now.ToString("dd") + "/" + DateTime.Now.ToString("yyyy") ;
            emailMetadata.EmailBody = GetEmailBody(notificationRecord, "CombiningCompletedEmailTemplate.txt");
            emailMetadata.HasAttachment = false;

            string emailException = emailManager.SendEmail(emailMetadata);
        }

        public static string GetEmailBody<T>(T notificationRecords, string template)
        {
            var templateService = new TemplateService();
            var formHtml = templateService.Parse(System.IO.File.ReadAllText(Path.Combine(FormsTemplateFolderPath, template)), notificationRecords, null, null);

            return formHtml;
        }
    }

}
