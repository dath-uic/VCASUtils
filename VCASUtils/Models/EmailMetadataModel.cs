using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace VCASPdfUtil.Models
{
    public class EmailMetadataModel
    {
        public string From { get; set; }
        public List<string> To { get; set; }
        public List<string> Cc { get; set; }
        public List<string> Bcc { get; set; }
        public string Subject { get; set; }
        public string EmailBody { get; set; }
        public string AttachmentNameWithoutExtn { get; set; }
        public bool HasAttachment { get; set; } = false;
        public bool EnableAdminBcc { get; set; } = true;
        public Attachment EmailAttachment { get; set; }
    }
}
