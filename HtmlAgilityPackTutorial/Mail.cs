using System;
using System.Net.Mail;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HtmlAgilityPackTutorial
{
    public class Mail
    {
        SmtpClient smtp;
        MailMessage ePosta;
        public string alici, konu, icerik, attachKonum;

        public Mail()
        {
            smtp = new SmtpClient();
            smtp.Credentials = new System.Net.NetworkCredential("ibrahim.okuyucu@outlook.com", "22La+3482");
            smtp.Port = 587;
            smtp.Host = "smtp-mail.outlook.com";
            smtp.EnableSsl = true;
        }

        public void Gonder()
        {
            Task.Factory.StartNew(() =>
            {
                try
                {
                    ePosta = new MailMessage();
                    ePosta.From = new MailAddress("ibrahim.okuyucu@outlook.com", "Test Bayi Portalı");
                    ePosta.To.Add(alici);
                    ePosta.Attachments.Add(new Attachment(attachKonum));
                    ePosta.Subject = konu;
                    ePosta.Body = icerik;

                    //smtp.SendAsync(ePosta, (object)ePosta);
                    if (!string.IsNullOrEmpty(alici))
                    {
                        smtp.Send(ePosta);
                    }
                }
                catch (Exception Ex)
                {
                    throw;
                }
            });
        }
    }
}