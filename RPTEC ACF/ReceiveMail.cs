using System;
using EAGetMail;
using System.IO;
using System.Threading.Tasks;
using System.Threading;
using System.Collections.Generic;
using System.Net.NetworkInformation;
using System.Security.Permissions;

namespace RPTEC_ACF
{
    class ReceiveMail
    {

        public Boolean FileSaved = false;
        public DateTime Today = DateTime.Now;
        //public MailClient PoClient;
        public int MailCounting = 0;
        public int MailPurging = 0;
        public int MailProcessed = 0;
        public Thread Work1;
        public Thread work2;
        public static Boolean DownloadedMail = false;
        public String FileName = "";




        //public Logs lg = new Logs();

        //public MailClient oClient = new MailClient("TryIt");        
        //public       MailInfo[] infos = oClient.GetMailInfos();


        public Boolean GetCorreo()
        {
            #region anterior
            // Crea un directorio y carpeta llamada mailbox
            // para guardar los correos en ella
            //string curpath = Directory.GetCurrentDirectory();
            //string mailbox = String.Format("{0}\\inbox", curpath);
            //string FilePath = "";
            //DateTime Today = DateTime.Now;

            //FilterMail ObjFil = new FilterMail();

            ////// si la carpeta no existe la crea(Mailbox)
            ////if (!Directory.Exists(mailbox))
            ////{
            ////    Directory.CreateDirectory(mailbox);
            ////}
            ////////"depuradormail@acfcapital.cl", "depu1606"
            //MailServer oServer = new MailServer("pop.acf",
            //            "depuradormail@acfcapital.cl", "depu1606", ServerProtocol.Pop3);
            ////MailClient oClient = new MailClient("TryIt");
            ////MailClient oClient = new MailClient("TryIt");
            //MailClient oClient = new MailClient("TryIt");
            //PoClient = oClient;

            //// activa la conexion SSL 
            //oServer.SSLConnection = true;
            //// oServer.SSLConnection = true;
            //// oServer.Port = 995;

            //// Setea el puerto 995 con seguridad SSL (certificado digital)
            //oServer.Port = 995;
            ////oServer.

            //try
            //{
            //    // Create a folder named "inbox" under current directory
            //    // to save the email retrieved.
            //    string curpath = Directory.GetCurrentDirectory();
            //    string mailbox = String.Format("{0}\\MailBox", curpath);

            //    // If the folder is not existed, create it.
            //    if (!Directory.Exists(mailbox))
            //    {
            //        Directory.CreateDirectory(mailbox);
            //    }                
            //        // Leave a copy of message on server.
            //        bool leaveCopy = true;
            //        // Download emails to this local folder
            //        string downloadFolder = mailbox;
            //        // Send request to EAGetMail Service, then EAGetMail Service retrieves email 
            //        // in background and this method returns immediately.
            //        oClient.GetMailsByQueue(oServer, downloadFolder, leaveCopy);

            //    //String FoldRouteAdj = @"\\10.177.1.230\mailacf$\ADJUNTOS"; //// ruta de adjuntos rptec
            //    //String SavePath = @"\\10.177.1.230\mailacf$\RPETC_ACF" + "\\";


            //    // Get all *.eml files in specified folder and parse it one by one.
            //    string[] files = Directory.GetFiles(mailbox, "*.eml");
            //    for (int i = 0; i < files.Length; i++)
            //    {
            //        //ParseEmail(files[i]);
            //        FilterMail fl = new FilterMail();
            //        fl.PendingProcess();

            //        //ObjFil.Filter(oMail, i);
            //}
            //Mail oMail;

            //oClient.Connect(oServer);
            //MailInfo[] infos = oClient.GetMailInfos();
            //int CountMsg = infos.Length;
            //for (int i = 0; i < CountMsg; i++)
            //{
            //item = i;
            //MailInfo info = infos[i];

            //// Receive email from POP3 server
            //oMail = oClient.GetMail(info);

            //ReceiveMail RV = new ReceiveMail();
            //SendMail SM = new SendMail();

            ////ruta donde se guardará el mensaje.
            //RV.SaveMail(oMail, i, ref SavePath);
            ////RV.SaveAttach(oMail, ref attname, FoldRouteAdj);

            //ObjFil.Filter(oMail, i);

            //FilterMail fl = new FilterMail();
            //fl.PendingProcess();
            //PendingProcess();
            //SaveMail(oMail,fileName,i);
            // Mark email as deleted from POP3 server.
            // oClient.Delete(info);

            //}

            //for (int i = 0; i < CountMsg; i++)
            //{
            //    //oClient.Connect(oServer);
            //    MailInfo info = infos[i];
            //    oClient.Delete(info);
            //}

            // Delete method just mark the email as deleted, 
            // Quit method pure the emails from server exactly.
            //oClient.Quit();

            //FilterMail fl = new FilterMail();
            //fl.PendingProcess();

            // Quit and pure emails marked as deleted from POP3 server.
            // oClient.Quit();

            ///preuba de log
            ////Logs.Log("ReceiveMail.GetCorreo", "");
            //string PathDirectory=Directory.GetCurrentDirectory();
            //PathDirectory = PathDirectory + @"\" + FilePath.;

            ////valida si existe el mail guardado si es asi elimina del servidor el correo(oClient.Quit)
            //if (FileSaved == true )
            //{
            //    // Quit and pure emails marked as deleted from POP3 server.
            //    oClient.Quit();
            //}
            //else
            //{
            //    ////No se guardo el mail ERROR*****
            //}

            #endregion 

            try
            {

                //DownloadMailsService();
                WaitingDownload();
                //ParseMailsFromInbox();
                //    ///donwload mail here:
                //    ThreadStart TDownload = DownloadMailsService;
                //    Thread Work1 = new Thread(TDownload);

                //    //whatcher here:
                //    ThreadStart TWaiting = WaitingDownload;
                //    Thread Work2 = new Thread(TWaiting);               


                //while (DownloadedMail == false)
                //{
                //    Thread.Sleep(10000); ///espera un minuto o mas bien 10 segundos
                //}
                //parsing here:
                if (DownloadedMail == true)
                {
                    ParseMailsFromInbox();
                }



                return true;
            } catch (Exception ep)
            {
                //Message contains the information returned by mail server
                Logs.Log("Error en Procesos GetCorreo: {0}", ep.Message.ToString());

                return false;
            }

        }

        private string GetLocalPathInbox()
        {
            string curpath = Directory.GetCurrentDirectory();

            FileManager Rcsv = new FileManager();
            System.Globalization.CultureInfo cur = new
                         System.Globalization.CultureInfo("en-US");
            //////System.Globalization.CultureInfo("en-US");

            string Sday = Today.ToString("yyyyMMdd", cur);
            string Mailbox = String.Format("{0}\\MailBox");//\\{ 1}", curpath, Sday);
            Rcsv.ScanFilePath(Mailbox);

            return Mailbox;
        }
        private void WaitingDownload()
        {

            try
            {
                String Path = GetLocalPathInbox();
                String FileName = "uidl.txt";
                String PathFile = string.Format("{0}\\{1}", Path, FileName);

                DateTime lastModified;
                DateTime Seconds;
                System.DateTime now = DateTime.Now;

                if (File.Exists(PathFile)) {
                    lastModified = File.GetLastWriteTime(Path);
                    Seconds = lastModified.AddSeconds(30);

                    while (now < Seconds)
                    {
                        Thread.Sleep(2000);
                        lastModified = File.GetLastWriteTime(Path);
                        Seconds = lastModified.AddSeconds(30);
                        now = DateTime.Now;

                    }

                    if (now > lastModified)
                        DownloadedMail = true;
                }
                else
                {


                }

                //DateTime lastModified = File.GetLastWriteTime(Path);
                //DateTime Seconds = lastModified.AddSeconds(30);
                //DateTime SumMinutes= lastModified.Second +


                //String Waitingdate = lastModified.ToString();
                //String Modifieddate =lastModified.ToString();
                //if (lastModified < Seconds && now > lastModified)

                //FileSystemWatcher watcher = new FileSystemWatcher();
                //watcher.IncludeSubdirectories = true;
                //watcher.Path = GetLocalPathInbox();

                //watcher.NotifyFilter = NotifyFilters.Attributes |
                //NotifyFilters.CreationTime |
                //NotifyFilters.DirectoryName |
                //NotifyFilters.FileName |
                //NotifyFilters.LastAccess;
                //// NotifyFilters.LastWrite |
                ////NotifyFilters.Security |
                ////NotifyFilters.Size;
                //watcher.Filter = "*.txt";

                //watcher.Changed += new FileSystemEventHandler(OnChanged);
                ////watcher.Error += new ErrorEventHandler(OnError);

                //watcher.EnableRaisingEvents = true;


            }
            catch (Exception ex)
            {
                Logs.Log("ReceiveMail.WaitingDownload, Error: ", ex.Message.ToString());

                DownloadedMail = false;
            }


        }

        //////public static void OnChanged(object source, FileSystemEventArgs e)
        //////{

        //////    List<string> AllChanges = new List<string>();
        //////    String Changes = String.Format("{0}, Con Ruta {1} Ha sido {2}", e.Name, e.FullPath, e.ChangeType);
        //////    Logs.Log(String.Format("Cambios en el documento:{0}", Changes), "- Datos Registrados");

        //////    AllChanges.Add(Changes.ToString());

        //////    if (Changes == "")
        //////    {
        //////        DownloadedMail = true;
        //////    }


        //////}



        private void DownloadMailsService()
        {
            try
            {
                Logs.Log("DownloadMailsService", "Tratando de descargar Correos - Iniciando Servicio...");

                FilterMail ObjFil = new FilterMail();
                MailServer oServer = new MailServer("pop.acf",
                            "depuradormail@acfcapital.cl", "depu1606", ServerProtocol.Pop3);

                MailClient oClient = new MailClient("TryIt");
                oServer.SSLConnection = true;
                oServer.Port = 995;

                string curpath = Directory.GetCurrentDirectory();
                string mailbox = String.Format("{0}\\MailBox\\", curpath);



                //FileManager Rcsv = new FileManager();
                //Rutas Para Guardar los Archivos de respaldo y Grabacion del mensaje
                //string PathHistoric = @"\\10.177.1.230\MAILACF$\RPETC_ACF"; //+ SToday;
                //string PathHistoric = SavePath;
                //System.Globalization.CultureInfo cur = new
                //             System.Globalization.CultureInfo("en-US");
                ////string sdate = Today.ToString("yyyyMMddHHmmss", cur);
                //string Sday = Today.ToString("yyyyMMdd", cur);

                ///crea la ruta del dia en que se recibio
                //string InboxDailyPath = String.Format(mailbox + Sday);
                //Rcsv.ScanFilePath(InboxDailyPath);                             

                bool leaveCopy = false;
                //string downloadFolder = InboxDailyPath;
                string downloadFolder = mailbox;
                oClient.GetMailsByQueue(oServer, downloadFolder, leaveCopy);
                Logs.Log("DownloadMailsService", "Se genero el proceso de Descarga Correctamente...");
                //return true;    

            }
            catch (Exception ep)
            {
                Logs.Log("Error Al Instanciar Servicio de Descarga de Correos: ", ep.Message.ToString());
                //return false;              
            }

        }

        private Boolean DoPing(string IPString)
        {
            Ping myPing = new Ping();
            int NoNet = 0;

            byte[] buffer = new byte[32];
            int timeout = 1000;
            PingOptions pingOptions = new PingOptions();

            try
            {
                NetworkInterface[] nics = NetworkInterface.GetAllNetworkInterfaces();

                foreach (NetworkInterface nic in nics)
                {
                    if (
                        (nic.NetworkInterfaceType != NetworkInterfaceType.Loopback && nic.NetworkInterfaceType != NetworkInterfaceType.Tunnel) &&
                        nic.OperationalStatus == OperationalStatus.Up)
                    {

                        PingReply reply = myPing.Send(IPString, timeout, buffer, pingOptions);
                        while (reply.Status != IPStatus.Success)
                        {
                            Thread.Sleep(5000);
                            NoNet++;

                            Logs.Log("Receive Mail.DoPing Error: ", "De Conexion Con La Direccion: " + IPString);
                            if (NoNet == 10)
                            {
                                SendMail sm = new SendMail();
                                string Asunto = "DEPURADOR MAIL ACF - Error de Red";
                                String BodyMail = "Se Perdio Conectividad con el NAS " + IPString + @"\MailAcf ,por mas de 50 segundos Favor Revisar";
                                sm.SenderMailAlert("Crosas@acfcapital.cl", Asunto, BodyMail);
                            }

                        }

                        //networkIsAvailable = true;
                        
                        if (reply.Status == IPStatus.Success)
                        {
                            // //presumably online
                            return true;
                        }
                        //else
                        //{
                            
                        //    Logs.Log("Receive Mail.DoPing Error: ", "De Conexion Con La Direccion: " + IPString);
                        //    //if (NoNet == 10)
                        //    //{
                        //    //    SendMail sm = new SendMail();
                        //    //    string Asunto = "DEPURADOR MAIL ACF - Error de Red";
                        //    //    String BodyMail = "Se Perdio Conectividad con el NAS "+ IPString + @"\MailAcf ,por mas de 50 segundos Favor Revisar";
                        //    //    sm.SenderMailAlert("Crosas@acfcapital.cl", Asunto, BodyMail);
                        //    //}
                        //    //Thread.Sleep(10000);
                        //    return false;
                        //}
                    }
                    //else
                    //{
                    //    NoNet++;
                    //    Logs.Log("Receive Mail.DoPing Error: ", "De Conexion - No hay red local : "+ NoNet + "Reintentos para la LAN:" + nic.ToString());                        
                    //    if (NoNet == 10)
                    //    {
                    //        SendMail sm = new SendMail();
                    //        string Asunto = "DEPURADOR MAIL ACF - Error de Red";
                    //        String BodyMail = "Se Perdio Conectividad con el NAS " + IPString + @"\MailAcf ,por mas de 50 segundos Favor Revisar";
                    //        sm.SenderMailAlert("Crosas@acfcapital.cl", Asunto, BodyMail);
                    //    }
                    //    Thread.Sleep(10000);
                        
                    //    return false;
                    //}
                }

                //SendMail sm = new SendMail();
                //string Asunto = "DEPURADOR MAIL ACF - Error de Red";
                //String BodyMail = "Se Perdio Conectividad con el NAS " + IPString + @"\MailAcf ,por mas de 50 segundos Favor Revisar";
                //sm.SenderMailAlert("Crosas@acfcapital.cl", Asunto, BodyMail);

                return false;

            }
            catch (Exception ex)
            {
                //if()
                Logs.Log("Receive Mail.DoPing Error: ", ex.Message.ToString());
                return false;
            }            
        }

        private Boolean ParseMailsFromInbox()
        {
            try
            {
                Logs.Log("ReceiveMail.ParseMailsFromInbox", "Tratando de Parsear Correos - Iniciando....");

                //List<string> Files = new List<string>();
                Mail oMail = new Mail("TryIt");
                //String MailPath = @""+Itm;
                /*oMail.LoadOMSG(Itm);  */  ////carga el msg como un objeto
                //oMail.Load(Itm, true);

                System.Globalization.CultureInfo cur = new
                             System.Globalization.CultureInfo("en-US");
                //string sdate = Today.ToString("yyyyMMddHHmmss", cur);
                string Sday = Today.ToString("yyyyMMdd", cur);

                string curpath = Directory.GetCurrentDirectory();
                string Mailbox = String.Format("{0}\\MailBox\\{1}", curpath, Sday);

                FileManager Rcsv = new FileManager();
                Rcsv.ScanFilePath(Mailbox);

                /////Get all *.eml files in specified folder and parse it one by one.
                string[] files = Directory.GetFiles(Mailbox, "*.eml");

                //foreach (string file in Directory.EnumerateFiles(Mailbox, "*.eml"))
                //{
                //    String Paths = Path.GetFullPath(file);
                //    Files.Add(Paths);
                //    //oMail.Load(file, true);
                //}

                FilterMail fl = new FilterMail();
                int Count = files.Length;
                MailCounting = files.Length;
                Logs.Log("Se contaron: " + MailCounting, " Correos disponibles en Inbox");
                for (int i = 0; i < Count; i++)
                {
                    String FileRoute =files[i];
                    String fname = Path.GetFileName(FileRoute);
                    oMail.Load(files[i], true);
                    fl.Filter(oMail, i, fname);

                    String Tipe = (fl.IsType);

                    String Origen = files[i];
                    String Destino = GetDestino(oMail, Tipe);

                    MoveMail(Origen, Destino); 
                    //if(Tipe==1 || Tipe == 2)
                    //    MailProcessed++;

                    //if (Tipe == 3 || Tipe == 4)
                    //{
                    //    MailProcessed++;
                    //}
                    //else
                    //{
                    //    MailPurging++;
                    //}
                        



                }

                Logs.Log("Procesados: " + MailProcessed, "de: "+ MailCounting +"Descartados: "+ MailPurging);


                //foreach (string itm in Files)
                //{

                //    oMail.Load(itm, true);
                //    fl.Filter(oMail,itm);


                //}

                //PendingLoadMail(Files);

                //int Count = files.Length;
                //for (int i = 0; i < Count; i++)
                //{
                //    //ParseEmail(files[i]);
                //    FilterMail fl = new FilterMail();
                //    fl.ProcessMail(files);
                //    fl.PendingProcess();
                //    //ObjFil.Filter(oMail, i);
                //}

                //Logs.Log("ReceiveMail.ParseMailsFromInbox", "Correos Parseados Correctamente. Cantidad: " + Count);
                return true;
            }
            catch (Exception e)
            {
                Logs.Log("ReceiveMail.ParseMailsFromInbox", "Error en Parsing Mails Desde Inbox: " + e.Message.ToString());
                return false;
            }

        }

        private String GetDestino(Mail oMail, String Tipe)
        {

            try
            {

                //Rutas Para Guardar los Archivos de respaldo y Grabacion del mensaje
                string Destino = @"\\10.177.1.220\MAILACF\";  /// RPETC_ACF"; //+ SToday;
                //string PathHistoric = SavePath;
                System.Globalization.CultureInfo cur = new
                             System.Globalization.CultureInfo("en-US");
                string sdate = Today.ToString("yyyyMMddHHmmss", cur);
                string Sday = Today.ToString("yyyyMMdd", cur);

                FilterMail fl = new FilterMail();
                String Ofrom = (oMail.From.ToString());
                //ReceiveMail rv = new ReceiveMail();
                fl.ShopEmail(ref Ofrom);


                switch (Tipe)
                {
                    case "1":    //////-  --1 = RPTC                            
                        Destino = Destino + @"RPETC";
                        MailProcessed++;
                        break;
                    case "2":   //////-- - 2 = XML
                        Destino = Destino + @"XML\" + Ofrom;
                        MailProcessed++;
                        break;
                    case "3":  //////-- - 3 = DICOM                                  
                        Destino = Destino + @"DICOM";
                        MailProcessed++;
                        break;
                    case "4":  //////-- - 4 = RESPALDOS
                        Destino = Destino + @"RESPALDOS\" + Ofrom;
                        MailProcessed++;
                        break;
                    default:  //// descartados
                        Destino = Destino + @"DESCARTADOS";
                        MailPurging++;
                        break;
                }

                Destino = String.Format("{0}\\{1}", Destino, Sday);

                /////crea la ruta del dia en que se recibio
                //string PathHist = String.Format(PathHistoric + Sday,
                //PathHistoric, sdate, Today.Millisecond.ToString("d3"), i);

                ///crea la ruta del mensaje de respaaldo a guardar

                //string HistFile = String.Format("{0}" + Sday + "\\{1}{2}{3}.msg",
                //           PathHistoric, sdate, Today.Millisecond.ToString("d3"), i);

                //determina si el directorio existe , sino lo crea.
                FileManager Rcsv = new FileManager();
                Rcsv.ScanFilePath(Destino);
                //sv.ScanFilePath(PathHist);

                return Destino;

            } catch (Exception ep)
            {
                Logs.Log("ReceiveMail.GetDestino", ep.Message.ToString());
                return "";
            }

            //return "";
        }

        private Boolean MoveMail(string sourceFile, string destinationFile)
        {
            //string sourceFile = @"C:\Users\Public\public\test.txt";
            //string destinationFile = @"C:\Users\Public\private\test.txt";
            try
            {
                //bool MoveIt = false;
                //int NoNet = 0;
                String Server = "10.177.1.220";
                //while (DoPing(Server) == false)// && MoveIt == false)
                //{
                //DoPing(Server);
                //while (DoPing(Server) == false)
                //{
                //    Thread.Sleep(3000);
                //    NoNet++;
                //    if (NoNet == 20)
                //    {
                //        SendMail sm = new SendMail();
                //        string Asunto = "DEPURADOR MAIL ACF - Error de Red";
                //        String BodyMail = "Se Perdio Conectividad con el NAS 10.177.1.220\\MailAcf ,por mas de 50 segundos Favor Revisar";
                //        sm.SenderMailAlert("Crosas@acfcapital.cl", Asunto, BodyMail);
                //    }
                //}


                //do
                //{
                //    string fullPathInternal = Path.GetFullPath(sourceFile);
                //    new FileIOPermission(FileIOPermissionAccess.Write | FileIOPermissionAccess.Read, new string[] { fullPathInternal }, false, false).Demand();
                //    string dst = Path.GetFullPath(destinationFile);
                if (DoPing(Server) == true)
                {
                    String Filename = Path.GetFileName(sourceFile);

                    String InsertingPathFile = destinationFile + "\\" + Filename;
                    if (!File.Exists(InsertingPathFile))
                    {
                        //int slen = (sourceFile).Length;
                        //int dlen = (destinationFile).Length;
                        //C:\Users\crosas\Desktop\copiaaaa

                        //destinationFile = @"\\10.177.1.220\Public\HOLAA\" + Filename;
                        File.Move(sourceFile, InsertingPathFile);
                        //File.Copy(sourceFile, destinationFile, true);

                        //File.Delete(sourceFile);
                        ////MoveIt = true;
                        return true;
                    }
                    else
                    {
                        File.Delete(destinationFile);
                        Logs.Log("Archivo: " + sourceFile, "ya existia, y fue reemplazado - Ruta: " + destinationFile);
                        return true;
                    }

                }
                    
                //}
                //while (DoPing(Server) == false);


               

                //}//else{

                //    Thread.Sleep(5000);
               // }
                //{


                // To move an entire directory. To programmatically modify or combine
                // path strings, use the System.IO.Path class.
                //System.IO.Directory.Move(@"C:\Users\Public\public\test\", @"C:\Users\Public\private");

                return false;

            }
            catch(IOException ep)
            {
                Logs.Log("Error: ", ep.Message.ToString());
                return false;
            }
            // To move a file or folder to a new location:
           
        }
        
        public Boolean SaveMail(Mail oMail ,int i,ref String SavePath)
        {
            ////rea un directorio y carpeta llamada mailbox
            ////para guardar los correos en ella
            string curpath = Directory.GetCurrentDirectory();
            string mailbox = String.Format("{0}\\inbox", curpath);
            string FilePath = "";
            //teTime Today = DateTime.Now;


            //FilterMail ObjFil = new FilterMail();

            // si la carpeta no existe la crea(Mailbox)
            if (!Directory.Exists(mailbox))
            {
                Directory.CreateDirectory(mailbox);
            }
            try
            {

            FileManager Rcsv = new FileManager();
                //Rutas Para Guardar los Archivos de respaldo y Grabacion del mensaje
                //string PathHistoric = @"\\10.177.1.230\MAILACF$\RPETC_ACF"; //+ SToday;
            string PathHistoric = SavePath;
            System.Globalization.CultureInfo cur = new
                         System.Globalization.CultureInfo("en-US");
            string sdate = Today.ToString("yyyyMMddHHmmss", cur);
            string Sday = Today.ToString("yyyyMMdd", cur);

                   //System.Globalization.CultureInfo cur = new
                        //System.Globalization.CultureInfo("en-US");
                    //string sdate = Today.ToString("yyyyMMddHHmmss", cur);
                    string fileName = String.Format("{0}\\{1}{2}{3}.msg",
                        mailbox, sdate, Today.Millisecond.ToString("d3"), i);
                    FilePath = fileName;


                ///crea la ruta del dia en que se recibio
            string PathHist = String.Format(PathHistoric + Sday,
            PathHistoric, sdate, Today.Millisecond.ToString("d3"), i);

                ///crea la ruta del mensaje de respaaldo a guardar

            string HistFile = String.Format("{0}"+Sday+"\\{1}{2}{3}.msg",
                       PathHistoric, sdate, Today.Millisecond.ToString("d3"), i);

                //determina si el directorio existe , sino lo crea.
            Rcsv.ScanFilePath(PathHistoric);
            Rcsv.ScanFilePath(PathHist);

            if (fileName != null)
                {
                    //oMail.SaveAs(fileName, true);guardar como eml
                    oMail.SaveAsOMSG(fileName, true,true);
                }        
            if(HistFile != null)
                {
                    //oMail.SaveAs(HistFile, true, true);guardar como eml
                    oMail.SaveAsOMSG(HistFile, true,true);
                    SavePath = HistFile;

                }
            if (fileName != null && PathHistoric != null)
                {
                    //GUARDO EXITOSAMENTE

                    //Logs.Log("ReceiveMail.SaveMail", "Guardo Exitosamente el MSG");
                    /////-------- delarar salida de servidor: Marcado de correos descargados y cerraar conexion
                    //  PoClient
                    
                    return true;
                }
                else
                {
                    Logs.Log("ReceiveMail.SaveMail", "ERROR- No se logró Guardar el MSG");
                    return false;
                }
                
                
            }
            catch (Exception ep)
            {
                Logs.Log("ReceiveMail.SaveMail", ep.Message.ToString());
                return false;                
                //throw;
            }
            
            //return true;
        }



    }//////termino de Clase


}  ////termino de namespace



//var client = new Pop3Client();
//ServicePointManager.ServerCertificateValidationCallback =

//       delegate (object s

//           , X509Certificate certificate

//           , X509Chain chain

//           , SslPolicyErrors sslPolicyErrors)

//       { return true; };

//    ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(ValidateServerCertificate);
//    client.Connect("pop.acf", 995, true);
//client.Authenticate("pruebasacf@acfcapital.cl", "prue2504");
//var count = client.GetMessageCount();
//Message message = client.GetMessage(count);
//Console.WriteLine(message.Headers.Subject);
//    return true;
//}


//public static bool ValidateServerCertificate(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
//{
//    if (sslPolicyErrors == SslPolicyErrors.None)
//        return true;
//    else
//    {
//        if (System.Windows.Forms.MessageBox.Show("The server certificate is not valid.\nAccept?", "Certificate Validation", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
//            return true;
//        else
//            return false;