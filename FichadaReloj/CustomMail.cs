using System;
using System.Configuration;
using System.Net.Mail;

namespace FichadaRelojService
{
    public class CustomMail
    {
        #region Properties

        public MailType Type { get; set; }

        public string Subject { get; set; }

        public string Content { get; set; }

        #endregion

        #region Constructors

        public CustomMail(MailType type, string subject, string content)
        {
            string emisor = ConfigurationManager.AppSettings["emisor"];
            this.Type = type;
            this.Subject = subject;
            
            this.Content = string.Format("{0} <br/> {1}{2}", content, "Mensaje enviado desde el emisor: ", emisor);
        }

        #endregion

        #region Public Methods

        public void Send()
        {
            var mailSettings = MailSettings.GetInstance();

            if (!this.CanSendErrorMail(mailSettings)) return;

            try
            {
                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient(mailSettings.Smtp);

                mail.From = new MailAddress(mailSettings.SupportMail);
                mail.To.Add(mailSettings.AdministratorMail);
                mail.Subject = this.Subject;
                mail.Body = this.Content;
                mail.IsBodyHtml = true;

                SmtpServer.Port = mailSettings.SmtpPort;
                SmtpServer.Credentials = new System.Net.NetworkCredential(mailSettings.SupportMail, mailSettings.SupportMailPassword);
                SmtpServer.EnableSsl = false;

                SmtpServer.Send(mail);
                Logger.GetInstance().AddLog(true, "MailService", string.Format("Se ha enviado un mail con el asunto: {0}", this.Subject));
            }
            catch (Exception ex)
            {
                Logger.GetInstance().AddLog(false, "MailService", string.Format("No se pudo enviar el mail. Error: {0}", ex));
            }
        }


        #endregion

        #region Private Methods

        private bool CanSendErrorMail(MailSettings mailSettings)
        {
            if (this.Type == MailType.Error)
            {
                if (DateTime.Now.Date > mailSettings.ActualDate.Date)
                {
                    mailSettings.ActualDate = DateTime.Now;
                    mailSettings.ErrorMailsSentToday = 0;
                }

                if (mailSettings.ErrorMailsSentToday > mailSettings.ErrorMailsSentDailyLimit)
                {
                    Logger.GetInstance().AddLog(false, "Mail.Send()", "Se llegó al límite diario de mails de error, no se pueden enviar más.");
                    return false;
                }

                if (mailSettings.LastErrorMailSentDate != null)
                {
                    double minutesDifference = DateTime.Now.TimeOfDay.TotalMinutes - mailSettings.LastErrorMailSentDate.TimeOfDay.TotalMinutes;
                    if (minutesDifference < 5)
                    {
                        Logger.GetInstance().AddLog(false, "Mail.Send()", "Deben pasar 5 minutos entre los distintos mails de error para no llenar la casilla.");
                        return false;
                    }
                    else
                    {
                        mailSettings.LastErrorMailSentDate = DateTime.Now;
                    }
                }
                else
                {
                    mailSettings.LastErrorMailSentDate = DateTime.Now;
                }
            }

            return true;
        }

        #endregion
    }

    public enum MailType
    {
        Information = 1,

        Error = 2
    }
}
