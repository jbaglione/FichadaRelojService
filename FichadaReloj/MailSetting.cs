using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FichadaRelojService
{
    public class MailSettings
    {
        #region Properties

        /// <summary>
        /// Instancia del Singleton
        /// </summary>
        private static MailSettings _instance;

        /// <summary>
        /// Para soportar multithreading
        /// </summary>
        private static object syncLock = new object();

        /// <summary>
        /// Obtiene la casilla de mail de soporte
        /// según lo especificado en el app.config.
        /// </summary>
        public string SupportMail = ConfigurationManager.AppSettings["supportMail"];

        /// <summary>
        /// Obtiene la casilla de mail de soporte
        /// según lo especificado en el app.config.
        /// </summary>
        public string AdministratorMail = ConfigurationManager.AppSettings["administratorMail"];

        /// <summary>
        /// Obtiene el password de mail de soporte
        /// según lo especificado en el app.config.
        /// </summary>
        public string SupportMailPassword = ConfigurationManager.AppSettings["supportMailPassword"];

        /// <summary>
        /// Obtiene la dirección smtp
        /// según lo especificado en el app.config.
        /// </summary>
        public string Smtp = ConfigurationManager.AppSettings["smtp"];

        /// <summary>
        /// Obtiene el puerto smtp
        /// según lo especificado en el app.config.
        /// </summary>
        public int SmtpPort = Convert.ToInt32(ConfigurationManager.AppSettings["smtpPort"]);

        /// <summary>
        /// Devuelve la maxima cantidad de mails que se pueden enviar en el día 
        /// </summary>
        public int ErrorMailsSentDailyLimit = Convert.ToInt32(ConfigurationManager.AppSettings["ErrorMailsSentDailyLimit"]);

        /// <summary>
        /// Devuelve y setea la cantidad de mails de error enviados en el día
        /// </summary>
        public int ErrorMailsSentToday { get; set; }

        /// <summary>
        /// Fecha que sirve para ver si actualizo mails de error enviados en el día
        /// </summary>
        public DateTime ActualDate { get; set; }

        /// <summary>
        /// Fecha/Hora del último mail de error enviado
        /// </summary>
        public DateTime LastErrorMailSentDate { get; set; }

        #endregion

        #region Constructor

        protected MailSettings()
        {
            this.ErrorMailsSentToday = 0;
            this.ActualDate = DateTime.Now;
            this.LastErrorMailSentDate = DateTime.Now.AddMinutes(-5);
        }

        #endregion

        #region Methods

        public static MailSettings GetInstance()
        {
            if (_instance == null)
            {
                lock (syncLock)
                {
                    if (_instance == null)
                    {
                        _instance = new MailSettings();
                    }
                }
            }

            return _instance;
        }

        #endregion
    }
}
