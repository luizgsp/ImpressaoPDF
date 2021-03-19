using GeracaoColetasPDF.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace GeracaoColetasPDF.Services
{
    class SendMail
    {
        public string SellerName { get; set; }
        public string EmailAccount { get; set; }
        public string Password { get; set; }
        public string EmailBody { get; set; }
        public string EmailCustomer { get; set; }
        public string Subject { get; set; }
        public string TargetFile { get; set; }
        public Priorities Priority { get; set; }

        public SendMail(string emailAccount, string emailBody, string emailCustomer, string subject, string targetFile,
            string sellerName, string password, Priorities priority)
        {
            SellerName = sellerName;
            EmailAccount = emailAccount;
            Password = password;
            EmailBody = emailBody;
            EmailCustomer = emailCustomer;
            Subject = subject;
            TargetFile = targetFile;
            Priority = priority;
        }

        public bool Mailing()
        {
            try
            {
                Config config = new Config();

                MailMessage mensagemEmail = new MailMessage(); //(ContaEmail, EmailCli, Assunto, enviaMensagem);
                mensagemEmail.Sender = new MailAddress(EmailAccount, "Minas Ferramentas - " + SellerName);
                mensagemEmail.From = new MailAddress(EmailAccount, "Minas Ferramentas - " + SellerName);
                mensagemEmail.To.Add(new MailAddress(EmailCustomer));
                mensagemEmail.Subject = Subject;
                mensagemEmail.Body = EmailBody;
                mensagemEmail.IsBodyHtml = true;
                mensagemEmail.Priority = (MailPriority) Priority;
                Attachment anexo = new Attachment(TargetFile);
                mensagemEmail.Attachments.Add(anexo);
                SmtpClient client = new SmtpClient();
                client.Host = config.SmtpServer;
                client.Port = config.SmtpPort;
                client.EnableSsl = false;
                client.UseDefaultCredentials = false;
                client.Credentials = new NetworkCredential(EmailAccount, Password);
                // envia a mensagem
                client.Send(mensagemEmail);
                return true;
            }
            catch (Exception e)
            {
                throw new Exception("Erro no envio da Mensagem:" + e.Message);
            }
            
        }
    }
}
