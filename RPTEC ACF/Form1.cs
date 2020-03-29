using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using EAGetMail;
using System.IO;
using System.Timers;
using System.Diagnostics;

namespace RPTEC_ACF
{
    public partial class Form1 : Form
    {
        public bool TimerON;
        int second = 0;
        //public bool running;

        ////BackgroundWorker que ayuda a realizar el trabajo en segundo plano, no afectanod la UI
       // public BackgroundWorker backworker = new BackgroundWorker();
        //delega la carga de trabajo en un hilo-seguro , de una manera segura
        //private delegate void ();

        public Form1()
        {
            InitializeComponent();
            //timer.Interval = 1000;
            //timer.Start();

        }


        private void backworker_DoWork(object sender, DoWorkEventArgs e)
        {
            Depurador();
        }


        private void Form1_Load_1(object sender, EventArgs e)
        {

            lbProcessState.Text = "Promagra Iniciado";
            lbProcessState.ForeColor = System.Drawing.Color.DarkOrange;

        }

        //private void Form1_Load(object sender, EventArgs e)
        //{
        //    //timer.Interval = 1000;
        //    //timer.Start();

        //    lbProcessState.Text = "CORRIENDO";
        //    lbProcessState.ForeColor = System.Drawing.Color.Green;
        //}


        //private void timer_Tick_1(object sender, EventArgs e)
        //{
        //    txtTimer.Text = DateTime.Now.ToString();
        //    second = second + 1;
        //    if (second >= 10)
        //    {
        //        timer.Stop();
        //        MessageBox.Show("Exiting from Timer....");
        //    }
        //}
        private void timer_Tick(object sender, EventArgs e)
        {
            txtTimer.Text = DateTime.Now.ToString();
            second = second + 1;
            if (second >= 10)
            {
                timer.Stop();
                MessageBox.Show("Exiting from Timer....");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            TimerON = true;        
            
           // bool Waiting = false;

            //if (TimerON == true)
            //{
            //lbProcessState.Text = "CORRIENDO";
            lbProcessState.Text = "DEPURANDO...";
            lbProcessState.ForeColor = System.Drawing.Color.Green;



            //backworker.RunWorkerAsync();

            //this.timer.Enabled = true;
            ////this.timer.Tick += new System.EventHandler(this.timer_Tick_1);
            //timer.Interval = 1000;
            //timer.Start();
            //timer_Tick_1(timer.Start(),EventArgs e);

            //}
            //sw.Start();
            //while (TimerON == true)
            //{
            //Stopwatch sw = new Stopwatch();
            //sw.Start();
            //while (sw.Elapsed < TimeSpan.FromSeconds(5))
            //{
            //    //Waiting = true;
            //    txtTimer.Text = (sw.Elapsed).ToString();                    
            //}

            //Waiting = false;  ///termina la espera

            //while (chkboxCiclo.Checked == true && TimerON == true)

            //Render();
            // Application.DoEvents();

            //{

            while (TimerON == true)
            {

                lbProcessState.Text = "DEPURANDO...";
                lbProcessState.ForeColor = System.Drawing.Color.Green;

                ReceiveMail ObjRecibeMail = new ReceiveMail();
                ObjRecibeMail.GetCorreo();
                Application.DoEvents();


                //Stopwatch sw = new Stopwatch();
                //sw.Start();
                //while (sw.Elapsed < TimeSpan.FromSeconds(60))
                //{
                //    //Waiting = true;

                //    //TimeSpan elapsed = GetElapsedTime();
                //    //txtTimer.Text = (sw.Elapsed).ToString("@m\:ss\.ff");
                //    // + sw.Elapsed.Minutes.ToString() + sw.Elapsed.Seconds.ToString();
                //    TimeSpan timeSpan = sw.Elapsed;
                //    String horas = timeSpan.Hours.ToString();
                //    String minutos = timeSpan.Minutes.ToString();
                //    String segundos = timeSpan.Seconds.ToString();
                //    String milisegundos = timeSpan.Milliseconds.ToString();
                //    txtTimer.Text = horas + minutos + segundos;
                //    //TimeSpan Hora= TimeSpan.Hours, TimeSpan.Minutes, TimeSpan.Seconds, TimeSpan.Milliseconds

                //    // ts.ToString("mm\\:ss\\.ff")
                //    Application.DoEvents();
                //}


            }
            //}

            //if (TimerON == true)
            //{
            //ReceiveMail ObjRecibeMail = new ReceiveMail();
            //ObjRecibeMail.GetCorreo();
            //}

            //lbProcessState.Text = "DEPURADO TERMINADO";
            //lbProcessState.ForeColor = System.Drawing.Color.Blue;

            //}

}

        private void Depurador()
        {
            while (chkboxCiclo.Checked == true && TimerON == true)
            {

                while (TimerON == true)
                {
                    Stopwatch sw = new Stopwatch();
                    sw.Start();
                    while (sw.Elapsed < TimeSpan.FromSeconds(90))
                    {
                        //Waiting = true;
                        txtTimer.Text = (sw.Elapsed).ToString();
                    }

                    lbProcessState.Text = "DEPURANDO...";
                    lbProcessState.ForeColor = System.Drawing.Color.Green;

                    ReceiveMail ObjRecibeMail = new ReceiveMail();
                    ObjRecibeMail.GetCorreo();
                }
            }


        }
        #region COMENTADO
        ////while (running==false)
        ////{
        ////    ReceiveMail ObjRecibeMail = new ReceiveMail();
        ////    ObjRecibeMail.GetCorreo();
        //}
        //r.Enabled = false;



        ////objeto Recibe Mail
        //            Contador Cn = new Contador();
        //TimerON=Cn.StartTimer();

        //if (TimerON == true)
        //{
        //    lbProcessState.Text = "CORRIENDO";
        //    lbProcessState.ForeColor = System.Drawing.Color.Green;
        //}


        //while (TimerON == true)
        //{
        //    ReceiveMail ObjRecibeMail = new ReceiveMail();
        //    ObjRecibeMail.GetCorreo();
        //}          

        ////}
        //private void timer_Elapsed(object sender, ElapsedEventArgs e)
        //{
        //    running = false;
        //}

        #endregion 

        private void button2_Click(object sender, EventArgs e)
        {
            ////DETIENE EL TIMER DE GET CORREO
            //Contador Cn = new Contador();
            //TimerON =Cn.StopTimer();
            TimerON = false;


            if (TimerON == false)
            {
                lbProcessState.Text = "DETENIDO";
                lbProcessState.ForeColor = System.Drawing.Color.Red;
                Application.DoEvents();
            }
        }

        private void chkboxCiclo_CheckedChanged(object sender, EventArgs e)
        {

            DepurarCorreo();
           //while( chkboxCiclo.Checked==true && TimerON==true)
           // {

           //     while (TimerON == true)
           //     {
           //         Stopwatch sw = new Stopwatch();
           //         sw.Start();
           //         while (sw.Elapsed < TimeSpan.FromSeconds(5))
           //         {
           //             //Waiting = true;
           //             txtTimer.Text = (sw.Elapsed).ToString();
           //         }

           //         lbProcessState.Text = "DEPURANDO...";
           //         lbProcessState.ForeColor = System.Drawing.Color.Green;

           //         ReceiveMail ObjRecibeMail = new ReceiveMail();
           //          ObjRecibeMail.GetCorreo();
           //     }
           // }

        }

        private void DepurarCorreo()
        {
            while (chkboxCiclo.Checked == true && TimerON == true)
            {

                while (TimerON == true)
                {
                    Stopwatch sw = new Stopwatch();
                    sw.Start();
                    while (sw.Elapsed < TimeSpan.FromSeconds(5))
                    {
                        //Waiting = true;
                        txtTimer.Text = (sw.Elapsed).ToString();
                    }

                    lbProcessState.Text = "DEPURANDO...";
                    lbProcessState.ForeColor = System.Drawing.Color.Green;

                    ReceiveMail ObjRecibeMail = new ReceiveMail();
                    ObjRecibeMail.GetCorreo();
                }
            }



        }
    }
}
