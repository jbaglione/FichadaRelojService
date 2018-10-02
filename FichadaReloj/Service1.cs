using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.IO;
using System.Net;
using ShamanClases.PanelC;
using System.Reflection;
using System.Configuration;
using ShamanClases;
using zkemkeeper;
using System.Net.Mail;

namespace FichadaRelojService
{
    public partial class Service1 : ServiceBase
    {

        Timer t = new Timer();
        private Conexion shamanConexion = new Conexion();
        string m_exePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

        string dBServer1 = ConfigurationManager.AppSettings["DBServer1"];
        string dBServer2 = ConfigurationManager.AppSettings["DBServer2"];
        string dBServer3 = ConfigurationManager.AppSettings["DBServer3"];
        string timePool = ConfigurationManager.AppSettings["TimePool"];
        DataTable dtRelojes;

        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            t.Elapsed += delegate { ElapsedHandler(); };
            t.Interval = 60000;
            t.Start();
            Logger.GetInstance().AddLog(true, "OnStart", "Servicio inicializado.");
            CustomMail mail = new CustomMail(MailType.Information, "Se ha inicializado el servicio de Fichada", "Servicio de Fichada");
            mail.Send();
            LeerRelojes();
        }

        protected override void OnPause()
        {
            Logger.GetInstance().AddLog(true, "OnPause", "Se ejecutó el método OnPause, el servicio deja de estar activo.");
            CustomMail mail = new CustomMail(MailType.Information, "Se ha pausado el servicio de Fichada", "Servicio de Fichada");
            mail.Send();
            t.Stop();
        }

        protected override void OnContinue()
        {
            Logger.GetInstance().AddLog(true, "OnPause", "Se ejecutó el método OnContinue, el servicio vuelve a estar activo.");
            t.Start();
        }

        protected override void OnStop()
        {
            Logger.GetInstance().AddLog(true, "OnStop", "Se ejecutó el método OnStop, el servicio deja de estar activo.");
            CustomMail mail = new CustomMail(MailType.Information, "Se ha detenido el servicio de Fichada", "Servicio de Fichada");
            mail.Send();
            t.Stop();
        }

        public void ElapsedHandler()
        {
            ProcesarRelojes();
        }

        private void ProcesarRelojes()
        {
            try
            {
                int vIdx;
                //--> Proceso relojes
                if (dtRelojes.Rows.Count > 0)
                {
                    Logger.GetInstance().AddLog(true, "ProcesarRelojes()", "Procesando " + dtRelojes.Rows.Count + " relojes");

                    for (vIdx = 0; vIdx <= dtRelojes.Rows.Count - 1; vIdx++)
                    {
                        DataRow dtRow = dtRelojes.Rows[vIdx];
                        this.clkZKSoft(Convert.ToInt32(dtRow["ID"]), 1, dtRow["Descripcion"].ToString(), dtRow["DireccionIP"].ToString(), Convert.ToInt32(dtRow["Puerto"]), 0);
                        // dtRow["DireccionIP"] = Pilar: "192.168.5.240"; Salta: "192.168.4.240"
                        //switch (Convert.ToInt32(dtRow["DriverId"])) //Pueden existir relojes con distintos drivers, y se conectan diferente.
                        //{ case 0:... case 1:...}
                    }
                }
                else
                    Logger.GetInstance().AddLog(true, "ProcesarRelojes()", "No se encontraron relojes, reinicie el servicio");

                if (modCache.cnnsCache.Count > 0)
                    modDeclares.ShamanSession.Cerrar(modDeclares.ShamanSession.PID);
            }
            catch (Exception ex)
            {
                Logger.GetInstance().AddLog(false, "ProcesarRelojes()", ex.Message);
            }
        }

        private void LeerRelojes()
        {
            try
            {
                if (this.ConnectServer(dBServer1))
                {
                    RelojesIngresos objRelojes = new RelojesIngresos();
                    Logger.GetInstance().AddLog(true, "LeerRelojes()", string.Format("Buscando relojes en ", dBServer1));

                    dtRelojes = objRelojes.GetAll();
                    modDeclares.ShamanSession.Cerrar(modDeclares.ShamanSession.PID);

                    Logger.GetInstance().AddLog(true, "LeerRelojes()", dtRelojes.Rows.Count > 0?
                        "Se encontraron " + dtRelojes.Rows.Count + " relojes":
                        "No se encontraron relojes");

                    if (modCache.cnnsCache.Count > 0)
                        modDeclares.ShamanSession.Cerrar(modDeclares.ShamanSession.PID);
                }
            }
            catch (Exception ex)
            {
                Logger.GetInstance().AddLog(false, "LeerRelojes()", ex.Message);
            }
        }

        private bool ConnectServer(string pCnnStr)
        {
            bool connectServer = false;

            try
            {
                string[] vArgs = pCnnStr.Split(';');
                int vIdx;

                if (vArgs.Length >= 5)
                {
                    for (vIdx = 0; vIdx <= vArgs.Length - 1; vIdx++)
                        vArgs[vIdx] = modConvert.Parcer(vArgs[vIdx], "=", 1);
                    if (shamanConexion.Iniciar(vArgs[0], Convert.ToInt32(vArgs[1]), vArgs[2], "EMERGENCIAS", "JAVIER", 1, false, vArgs[3], vArgs[4]))
                    {
                        Logger.GetInstance().AddLog(true, "ConnectServer()", "Cache ConnectionString" + modCache.cnnsCache["DefaultStatic"].ConnectionString);
                        connectServer = true;
                    }
                    else
                        Logger.GetInstance().AddLog(true, "ConnectServer()", "No se pudo conectar a " + pCnnStr);
                }
                return connectServer;
            }
            catch (Exception ex)
            {
                Logger.GetInstance().AddLog(false, "ConnectServer()", pCnnStr + " - " + ex.Message);
            }

            return connectServer;
        }

        private bool clkZKSoft(int pRid, int pNro, string pDes, string pDir, int pPor, long pPsw)
        {
            bool clkZKSoft = false;

            try
            {
                string sdwEnrollNumber = "";
                int idwVerifyMode;
                int idwInOutMode;
                int idwYear;
                int idwMonth;
                int idwDay;
                int idwHour;
                int idwMinute;
                int idwSecond;
                int idwWorkcode = 0;
                string vFic;
                bool vClean = false;
                CZKEM Reloj = new CZKEM();
                DevOps devolucionOperacion = new DevOps();

                RelojesIngresos objFichada = new RelojesIngresos();
                List<RelojResponse> relojResponse = new List<RelojResponse>();
                //DevOps devolucionOperacion = new DevOps();

                //if (pDir == "192.168.0.241")
                //    pDir = "192.168.5.111"; // Reloj Pilar
                // If pDir = "192.168.0.241" Then pDir = "200.49.156.125"
                // If pDir = "200.85.127.22" Then pDir = "192.168.5.125"
                // If pDir = "192.168.4.240" Then
                // pDir = "661705e2a569.sn.mynetname.net"
                // pPor = 64370
                // End If
                //FuncionPrueba(pRid);
                Logger.GetInstance().AddLog(true, "clkZKSoft()", "Conectandose " + pDir + ":" + pPor);
                if (Reloj.Connect_Net(pDir, pPor))/* && false)*/
                {
                    
                    Logger.GetInstance().AddLog(true, "clkZKSoft()", "Conectado a " + pDir + ":" + pPor);
                    Reloj.EnableDevice(pNro, false);

                    // ----> Leo Datos
                    if (Reloj.ReadGeneralLogData(pNro))
                    {

                        // SSR_GetGeneralLogData
                        // ----> Leo Datos
                        while (Reloj.SSR_GetGeneralLogData(pNro, out sdwEnrollNumber, out idwVerifyMode, out idwInOutMode, out idwYear, out idwMonth, out idwDay, out idwHour, out idwMinute, out idwSecond, ref idwWorkcode))
                        {
                            if (!vClean)
                                Logger.GetInstance().AddLog(true, "clkZKSoft()", "Procesando registros de " + pDir + ":" + pPor);

                            vClean = true;
                            //Logger.GetInstance().AddLog(true, "clkZKSoft()", "Revisando Legajo: " + sdwEnrollNumber);

                            if (idwYear == DateTime.Now.Year)
                            {
                                RelojResponse relojResponseItem = new RelojResponse();
                                vFic = string.Format(idwYear.ToString("0000")) + "-" + string.Format(idwMonth.ToString("00")) + "-" + string.Format(idwDay.ToString("00")) + " " + String.Format(idwHour.ToString("00")) + ":" + String.Format(idwMinute.ToString("00")) + ":" + String.Format(idwSecond.ToString("00"));
                                //Logger.GetInstance().AddLog(true, "clkZKSoft()", "Fecha del Registro: " + vFic);
                                relojResponseItem.Fich = vFic;
                                relojResponseItem.Nro = pNro;
                                relojResponseItem.SdwEnrollNumber = sdwEnrollNumber;
                                relojResponseItem.IdwVerifyMode = idwVerifyMode;
                                relojResponseItem.IdwInOutMode = idwInOutMode;
                                relojResponseItem.IdwWorkcode = idwWorkcode;

                                relojResponse.Add(relojResponseItem);
                            }
                        }
                        
                        if (relojResponse.Count > 0)
                        {
                            if (ConnectServer(dBServer1))
                            {
                                Logger.GetInstance().AddLog(true, "clkZKSoft()", "Envios a Server1: " + dBServer1);
                                Logger.GetInstance().AddLog(true, "clkZKSoft()", "Cantidad de Registros: " + relojResponse.Count);

                                foreach (RelojResponse item in relojResponse)
                                {
                                    Logger.GetInstance().AddLog(true, "clkZKSoft()", "Revisando Legajo: " + item.SdwEnrollNumber);
                                    Logger.GetInstance().AddLog(true, "clkZKSoft()", "Fecha del Registro: " + item.Fich);
                                    try
                                    {
                                        devolucionOperacion = objFichada.SetFichada(pRid, item.SdwEnrollNumber, item.Fich, "CLOCK");

                                        Logger.GetInstance().AddLog(true, "clkZKSoft()", "SetFichada CacheDebug: " + devolucionOperacion.CacheDebug);

                                        if (devolucionOperacion.Resultado)
                                            Logger.GetInstance().AddLog(true, "clkZKSoft()", "SetFichada OK");
                                        else
                                            Logger.GetInstance().AddLog(false, "clkZKSoft()", "SetFichada Error: " + devolucionOperacion.DescripcionError);
                                    }
                                    catch (Exception ex)
                                    {
                                        Logger.GetInstance().AddLog(false, "clkZKSoft()", "Excepción en SetFichada: " + ex.Message);
                                    }
                                }
                                modDeclares.ShamanSession.Cerrar(modDeclares.ShamanSession.PID);
                            }

                            if (ConnectServer(dBServer2))
                            {
                                Logger.GetInstance().AddLog(true, "clkZKSoft()", "Envios a Server2: " + dBServer2);
                                Logger.GetInstance().AddLog(true, "clkZKSoft()", "Cantidad de Registros: " + relojResponse.Count);
                                foreach (RelojResponse item in relojResponse)
                                {
                                    Logger.GetInstance().AddLog(true, "clkZKSoft()", "Revisando Legajo: " + item.SdwEnrollNumber);
                                    Logger.GetInstance().AddLog(true, "clkZKSoft()", "Fecha del Registro: " + item.Fich);
                                    try
                                    {
                                        devolucionOperacion = objFichada.SetFichada(pRid, item.SdwEnrollNumber, item.Fich, "CLOCK");
                                        Logger.GetInstance().AddLog(true, "clkZKSoft()", "SetFichada CacheDebug: " + devolucionOperacion.CacheDebug);

                                        if (devolucionOperacion.Resultado)
                                            Logger.GetInstance().AddLog(true, "clkZKSoft()", "SetFichada OK");
                                        else
                                            Logger.GetInstance().AddLog(false, "clkZKSoft()", "SetFichada Error: " + devolucionOperacion.DescripcionError);
                                    }
                                    catch (Exception ex)
                                    {
                                        Logger.GetInstance().AddLog(false, "clkZKSoft()", "Excepción en SetFichada: " + ex.Message);
                                    }
                                }
                                modDeclares.ShamanSession.Cerrar(modDeclares.ShamanSession.PID);
                            }

                            if (ConnectServer(dBServer3))
                            {
                                Logger.GetInstance().AddLog(true, "clkZKSoft()", "Envios a Server3: " + dBServer3);
                                Logger.GetInstance().AddLog(true, "clkZKSoft()", "Cantidad de Registros: " + relojResponse.Count);
                                foreach (RelojResponse item in relojResponse)
                                {
                                    Logger.GetInstance().AddLog(true, "clkZKSoft()", "Revisando Legajo: " + item.SdwEnrollNumber);
                                    Logger.GetInstance().AddLog(true, "clkZKSoft()", "Fecha del Registro: " + item.Fich);
                                    try
                                    {
                                        devolucionOperacion = objFichada.SetFichada(pRid, item.SdwEnrollNumber, item.Fich, "CLOCK");
                                        Logger.GetInstance().AddLog(true, "clkZKSoft()", "SetFichada CacheDebug: " + devolucionOperacion.CacheDebug);

                                        if (devolucionOperacion.Resultado)
                                            Logger.GetInstance().AddLog(true, "clkZKSoft()", "SetFichada OK");
                                        else
                                            Logger.GetInstance().AddLog(false, "clkZKSoft()", "SetFichada Error: " + devolucionOperacion.DescripcionError);
                                    }
                                    catch (Exception ex)
                                    {
                                        Logger.GetInstance().AddLog(false, "clkZKSoft()", "Excepción en SetFichada: " + ex.Message);
                                    }
                                }
                                modDeclares.ShamanSession.Cerrar(modDeclares.ShamanSession.PID);
                            }
                        }

                        if (vClean)
                        {
                            Logger.GetInstance().AddLog(true, "clkZKSoft()", "Vaciar Reloj " + pNro + " Ip: " + pDir + ":" + pPor);
                            if (Reloj.ClearGLog(pNro))
                            {
                                Reloj.RefreshData(pNro);
                                Logger.GetInstance().AddLog(true, "clkZKSoft()", "Se vació RelojId " + pNro + " Ip: " + pDir + ":" + pPor);
                            }
                            else
                            {
                                int idwErrorCode = 0;
                                Reloj.GetLastError(idwErrorCode);
                                Logger.GetInstance().AddLog(false, "clkZKSoft()", "Error al vaciar " + pDir + ":" + pPor + " " + idwErrorCode);
                            }

                            vClean = false;
                        }
                        else
                            Logger.GetInstance().AddLog(true, "clkZKSoft()", "No hay fichadas en " + pDir + ":" + pPor);
                    }
                    else
                        Logger.GetInstance().AddLog(true, "clkZKSoft()", "No hay fichadas en " + pDir + ":" + pPor);

                    Logger.GetInstance().AddLog(true, "clkZKSoft()", "Desconectar Reloj " + pDir + ":" + pPor);
                    Reloj.Disconnect();

                    clkZKSoft = true;
                }
                else
                    Logger.GetInstance().AddLog(true, "clkZKSoft()", "Sin conexión a " + pDir + ":" + pPor);
            }
            catch (Exception ex)
            {
                Logger.GetInstance().AddLog(false, "clkZKSoft()", ex.Message);
            }
            return clkZKSoft;
        }

        private void FuncionPrueba(int pRid)
        {
            

            if (ConnectServer(dBServer1))
            {
                Logger.GetInstance().AddLog(true, "FuncionPrueba()", "pRid: " + pRid.ToString());
                Logger.GetInstance().AddLog(true, "FuncionPrueba()", "Envios a Server1: " + dBServer1);

                FuncionPrueba2(1, "9109");

                modDeclares.ShamanSession.Cerrar(modDeclares.ShamanSession.PID);
            }
        }

        private static DevOps FuncionPrueba2(int pRid, string SdwEnrollNumber)
        {
            RelojesIngresos objFichada = new RelojesIngresos();
            DevOps devolucionOperacion = objFichada.SetFichada(pRid, SdwEnrollNumber, "2018-07-10 15:13:14", "CLOCK");
            Logger.GetInstance().AddLog(true, "FuncionPrueba2()", "SetFichada CacheDebug: " + devolucionOperacion.CacheDebug);

            if (devolucionOperacion.Resultado)
                Logger.GetInstance().AddLog(true, "FuncionPrueba2()", "SetFichada OK");
            else
                Logger.GetInstance().AddLog(false, "FuncionPrueba2()", "SetFichada Error: " + devolucionOperacion.DescripcionError);
            return devolucionOperacion;
        }
    }
}
