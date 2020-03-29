using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.NetworkInformation;
using System.Net;
using System.Diagnostics;
using System.Timers;
using System.Diagnostics.PerformanceData;
using System.Threading;
using System.Windows.Forms;

namespace RPTEC_ACF
{ 
    class Contador
    {
        public System.Timers.Timer aTimer = new System.Timers.Timer();
        public Boolean TimerOn;
        public bool StartTimer()
        {

            try
            {
                aTimer.Interval = 60000; //sesenta segundos
                aTimer.Enabled = true;
                if (aTimer.Enabled == true)
                {
                    TimerOn = true;
                }
                return true;
            }
            catch(Exception e)
            {
                Logs.Log("Contador.StartTimer", e.Message.ToString());
                return false;
            }   
            //aTimer.Elapsed += new ElapsedEventHandler(OnTimedEvent);
            
        }

        public bool StopTimer()
        {
            aTimer.Enabled = false;            
            return false;
        }

    }
}
