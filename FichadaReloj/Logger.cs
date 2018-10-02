using ShamanClases;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FichadaRelojService
{
    public class Logger
    {
        #region Properties

        /// <summary>
        /// Instancia del Logger
        /// </summary>
        private static Logger _instance;

        /// <summary>
        /// Para multithreading
        /// </summary>
        private static object syncLock = new object();

        #endregion

        #region Constructor

        protected Logger()
        {

        }

        #endregion

        #region Methods

        public static Logger GetInstance()
        {
            if (_instance == null)
            {
                lock (syncLock)
                {
                    if (_instance == null)
                    {
                        _instance = new Logger();
                    }
                }
            }

            return _instance;
        }

        public void AddLog(bool rdo, string logProcedure, string logDescription)
        {
            string path;

            path = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            path = path + "\\" + modFechas.DateToSql(DateTime.Now).Replace("-", "_") + ".log";

            if (!File.Exists(path))
            {
                using (StreamWriter sw = File.CreateText(path))
                {
                    sw.WriteLine("Log " + DateTime.Now.Date);
                }
            }

            using (StreamWriter sw = File.AppendText(path))
            {
                string rdoStr = rdo ? "Ok" : "Error";
                sw.WriteLine(DateTime.Now.Hour.ToString("00") + ":" + DateTime.Now.Minute.ToString("00") + "\t" + rdoStr + "\t" + logProcedure + "\t" + logDescription);
            }
        }

        #endregion
    }
}
