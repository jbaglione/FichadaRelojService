using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FichadaRelojService
{
    public class RelojResponse
    {
        private int rrNro;
        private string rrSdwEnrollNumber;
        private int rrIdwVerifyMode;
        private int rrIdwInOutMode;
        private int rrIdwWorkcode;
        private string rrFich;

        // = Format(idwYear, "0000") & "-" & Format(idwMonth, "00") & "-" & Format(idwDay, "00") & " " & Format(idwHour, "00") & ":" & Format(idwMinute, "00") & ":" & Format(idwSecond, "00")
        public int Nro
        {
            get
            {
                return rrNro;
            }
            set
            {
                rrNro = value;
            }
        }
        public string SdwEnrollNumber
        {
            get
            {
                return rrSdwEnrollNumber;
            }
            set
            {
                rrSdwEnrollNumber = value;
            }
        }
        public int IdwVerifyMode
        {
            get
            {
                return rrIdwVerifyMode;
            }
            set
            {
                rrIdwVerifyMode = value;
            }
        }
        public int IdwInOutMode
        {
            get
            {
                return rrIdwInOutMode;
            }
            set
            {
                rrIdwInOutMode = value;
            }
        }
        public int IdwWorkcode
        {
            get
            {
                return rrIdwWorkcode;
            }
            set
            {
                rrIdwWorkcode = value;
            }
        }
        public string Fich
        {
            get
            {
                return rrFich;
            }
            set
            {
                rrFich = value;
            }
        }
    }
}
