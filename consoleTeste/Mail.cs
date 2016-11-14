using System;
using System.Net;
using System.Net.Mail;

namespace consoleTeste
{
    public static class Mail
    {
        public static bool EnviaEmail(string destinatarios, string assunto, string mensagem, out string exception)
        {
            try
            {
                var sendermail = new MailAddress("leandrofire@live.com", "Conta de Serviço");
                var password = "DlA685947";
                SmtpClient smtp = new SmtpClient
                {
                    Host = "smtp.live.com",
                    Port = 587,
                    EnableSsl = true,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    UseDefaultCredentials = false,
                    Credentials = new NetworkCredential(sendermail.Address, password)
                };
                var mess = new MailMessage();
                string[] arrDestinatarios = destinatarios.Split(';');
                foreach (string destinatario in arrDestinatarios)
                {
                    mess.To.Add(new MailAddress(destinatario)); 
                }
                string[] arrDestinatarioscopia = destinatarios.Split(';');
                foreach (string cc in arrDestinatarioscopia)
                {
                    mess.CC.Add(new MailAddress(cc)); //Adding Multiple CC email Id  
                }
                mess.From = sendermail;
                mess.Subject = assunto;
                mess.Body = mensagem;
                mess.IsBodyHtml = true;
                
                    smtp.Send(mess);
                
                exception = "Sucesso";
                return true;
            }
            catch (Exception e)
            {
                exception = e.ToString();
                return false;
            }
        }
    }
}

