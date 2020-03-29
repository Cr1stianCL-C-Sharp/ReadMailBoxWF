
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System;
using System.IO;

namespace RPTEC_ACF
{
    class Logs
    {

        //public static void Log()
        //{
            public static DateTime Today = DateTime.Now;
            //public System.Globalization.CultureInfo cur = new
            //public System.Globalization.CultureInfo("en-US");
            //public string Stoday = Today.ToString("yyyyMMdd", cur);
            //public string LogDirectory = Directory.GetCurrentDirectory();
            //LogDirectory = String.Format("{0}\\Log\\" + Stoday + "\\", LogDirectory);//+"Log_" + Stoday +".txt", LogDirectory);
            //public String FilePath = (LogDirectory + "Log_" + Stoday + ".txt");
            //using (StreamReader r = File.OpenText("log.txt"))
            //{
            //    DumpLog(r);
            //}
        //}

        public static void Log(string logMessage ,string e)
        {
        FileManager Fm = new FileManager();
        System.Globalization.CultureInfo cur = new
        System.Globalization.CultureInfo("en-US");
        string Stoday = Today.ToString("yyyyMMdd", cur);
        string LogDirectory = Directory.GetCurrentDirectory();
        LogDirectory = String.Format("{0}\\Log\\" + Stoday + "\\", LogDirectory);//+"Log_" + Stoday +".txt", LogDirectory);
        String FilePath = (LogDirectory + "Log_" + Stoday + ".txt");

            //if (Directory.Exists(LogDirectory))
            //{
            //    Directory.CreateDirectory(LogDirectory);
            //Fm.ScanFilePath(LogDirectory);
            /////Fm.ScanFilePath(PathHist);
            if (Fm.ScanFilePath(LogDirectory) == true)
            {
                using (StreamWriter w = File.AppendText(FilePath))

                {
                    if (File.Exists(FilePath))
                    {
                        
                        w.Write("\r\nLog Entry : ");
                        w.WriteLine("{0} {1}", DateTime.Now.ToLongTimeString(),
                            DateTime.Now.ToLongDateString());
                        w.WriteLine("  :");
                        w.WriteLine("  :{0} {1}", logMessage, e);
                        w.WriteLine("-------------------------------");
                        w.Flush();
                        w.Close();
                        w.Dispose();
                    }   
                }

            }      


            //StreamWriter w =new StreamWriter(FilePath);
            //////if (!Directory.Exists(LogDirectory))
            //////{
            //////    Directory.CreateDirectory(LogDirectory);
            //////    if (!File.Exists(FilePath))
            //////    {
            //////        File.CreateText(FilePath);
            //////        w = File.CreateText(FilePath);

            //////        w.Flush();
            //////        w.Close();
            //////        w.Dispose();

            //////        using (StreamWriter sw = new StreamWriter(FilePath)
            //////        using (w = File.AppendText(FilePath))
            //////        {
            //////            string logMessage = string.Empty;
            //////            Exception e;
            //////            Log(logMessage, e, w);
            //////            Log("Test2", w);
            //////            w.Write("\r\nLog Entry : ");
            //////            w.WriteLine("{0} {1}", DateTime.Now.ToLongTimeString(),
            //////                DateTime.Now.ToLongDateString());
            //////            w.WriteLine("  :");
            //////            w.WriteLine("  :{0} {1}", logMessage, e);
            //////            w.WriteLine("-------------------------------");
            //////            w.Flush();
            //////            w.Close();
            //////            w.Dispose();


            //////            objFileLogger.Flush();
            //////            objFileLogger.Close();
            //////            objFileLogger.Dispose()
            //////        }

            //////    }
            //////}

            //w.Close();
            //StreamWriter w = new StreamWriter();
            //StreamWriter w =new StreamWriter(FilePath);





        }

        //public static void DumpLog(StreamReader r)
        //{
        //    string line;
        //    while ((line = r.ReadLine()) != null)
        //    {
        //        Console.WriteLine(line);
        //    }
        }
    }
