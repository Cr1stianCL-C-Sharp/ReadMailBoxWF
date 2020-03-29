using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EAGetMail;
using System.IO;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.Configuration;
using System.Data;

namespace RPTEC_ACF
{
    class FilterMail
    {
        public DateTime Today = DateTime.Now;        

        public void Filter(Mail oMail,int i)
        {
           try
            {             
                String oFrom = (oMail.From).ToString();
                String xMail = oFrom;
                string Tipe_Mail = string.Empty;
                StringBuilder sb = new StringBuilder();
                SqlCommand cmd = new SqlCommand();
                SqlDataReader reader;

                string connString = ConfigurationManager.ConnectionStrings["myConnectionString"].ConnectionString;
                SqlConnection Connex = new SqlConnection(connString);

                ShopEmail(ref xMail);

                sb.Append("if exists(select dc_mail from tb_filtro_mail where dc_mail = ");
                sb.Append("'" + xMail + "')");
                sb.Append("begin select dn_tipo from tb_filtro_mail where dc_mail =");
                sb.Append("'" + xMail + "'end");
            
                    Connex.Open();
                if (Connex != null && Connex.State == ConnectionState.Open)
                {
                    string query= sb.ToString();
                    cmd.CommandText = query;
                    cmd.Connection = Connex;
                    reader = cmd.ExecuteReader();
                    reader.Read();

                        if (reader.HasRows)
                        {                                
                           Tipe_Mail = reader["dn_tipo"].ToString();             
                        }
                 }
                 else
                 {
                     //error al log
                 }
                    Connex.Close();

                switch (Tipe_Mail)
                {
                    case "1":    //////-  --1 = RPTC
                        RptcFilter(oMail, i);    /////VALIDADO
                                                
                        break;
                    case "2":   //////-- - 2 = XML
                        XmlFilter(oMail, i);     ////VALIDADO
                        
                        break;
                    case "3":  //////-- - 3 = BODY -- demasiado variable no se tomo en cuenta
                        //BodyFilter(oMail, i);
                        //Saved = true;
                        DicomFilter(oMail, i);                        
                        break;
                    case "4":  //////-- - 4 = BACKUP
                        BackUpFilter(oMail, i);   /////VALIDADO
                        
                        break;
                    
                    default:
                        DiscardFilter(oMail, i);
                        break;
                }

              // PendingProcess();
                
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.Filter", e.Message.ToString());
            }

            #region commented
            //Connex.Open();
            //if (Connex != null && Connex.State == ConnectionState.Open)
            //{

            //    //if exists(
            //    //select dc_mail from tb_filtro_mail where dc_mail = 'enviodte@servifactura.cl')
            //    //begin select dn_tipo from tb_filtro_mail where dc_mail = 'enviodte@servifactura.cl'
            //    //end
            //    sb.Append("if exists(select dc_mail from tb_filtro_mail where dc_mail = ");
            //    sb.Append("'" + xMail + "')");
            //    sb.Append("begin select dn_tipo from tb_filtro_mail where dc_mail =");
            //    sb.Append("'" + xMail + "'end");

            //    cmd.CommandText = sb.ToString();
            //    //DataSet ds = new DataSet();
            //    cmd.Connection = Connex;
            //    reader = cmd.ExecuteReader();
            //    reader.Read();

            //Connex.Open();
            //adapter.SelectCommand = new SqlCommand(cmd.CommandText, cmd.Connection);
            //adapter.Fill(ds);
            //Connex.Close();
            //for (i = 0; i <= ds.Tables[0].Rows.Count - 1; i++)
            //{
            //MessageBox.Show(ds.Tables[0].Rows[1].ItemArray[1].ToString());
            //}

            //if (reader.HasRows)
            //        {
            //            string tf = reader["dn_tipo"].ToString();

            //        }
            //     }
            //        else
            //            {
            //                //error al log

            //            }


            //    Connex.Close();

            //}
            //catch (Exception)
            //{
            //    throw;
            //}///termino del try/catch
            #endregion

            #region Comentarizado


            //  String oFrom = (oMail.From).ToString();
            //  ShopODestino(ref oFrom);
            //  bool Saved = false;

            //  if (oFrom == "crosas")  /////  || oFrom== "crosas"
            //  {
            //      XmlFilter(oMail,i);
            //      Saved = true;
            //  }
            //  if (oFrom == "tcuello")  /////  || oFrom== "crosas"//tcuello@arayahermanos.cl
            //  {
            //      XmlFilter(oMail, i);
            //      Saved = true;
            //  }


            //  if (oFrom == "administracion")  /////  || oFrom== "crosas"//administracion@labotec.cl
            //  {
            //      XmlFilter(oMail, i);
            //      Saved = true;
            //  }

            //  if (oFrom == "zuniga")  /////  || oFrom== "crosas"  //zuniga@decomuebles.cl
            //  {
            //      XmlFilter(oMail, i);
            //      Saved = true;
            //  }

            //  if (oFrom == "zuniga")  /////  || oFrom== "crosas" //dte@bsale.cl
            //  {
            //      XmlFilter(oMail, i);
            //      Saved = true;
            //  }              

            //  //dte_dilorsa@empresasorion.cl        
            //  if (oFrom == "dte_dilorsa")  /////  || oFrom== "crosas"//administracion@labotec.cl
            //  {
            //      XmlFilter(oMail, i);
            //      Saved = true;
            //  }

            //  ////sii@crosan.cl
            //  if (oFrom == "sii")  /////  || oFrom== "crosas"   ////sii@crosan.cl
            //  {
            //      XmlFilter(oMail, i);
            //      Saved = true;
            //  }

            //  //simtexx@simtexx.cl
            //  if (oFrom == "simtexx")  /////  || oFrom== "crosas"//simtexx@simtexx.cl
            //  {
            //      XmlFilter(oMail, i);
            //      Saved = true;
            //  }


            //  ////si es de servifactura  --- enviodte  -- enviodte@servifactura.cl
            //  if (oFrom == "enviodte")  /////  || oFrom== "crosas"
            //  {                    
            //    ServiFacturaFilter(oMail, i);
            //      Saved = true;
            //  }

            //  // Si es RPTC
            //  if (oFrom == "rpetc" )  /////  || oFrom== "crosas"
            //  {
            //  RptcFilter(oMail, i);
            //      Saved = true;            
            //  }
            //  // SI ES DE dte@move-up.cl
            //  if (oFrom == "dte")  /////  || oFrom== "crosas"  //dte@move-up.cl
            //  {
            //      RptcFilter(oMail, i);
            //      Saved = true;
            //  }
            //  //si es  DeptoControlGestion
            //  if (oFrom == "DeptoControlGestion") /////|| oFrom == "crosas"
            //  {
            //    DeptoControlGestionFilter(oMail, i);
            //      Saved = true;
            //  }
            //  //si es dicom
            //  if (oFrom == "informes.equifax")  //////informes.equifax@equifax.cl
            //  {
            // DicomFilter(oMail, i);
            //      Saved = true;
            //  }

            //  //// si no entra a ninguno de los if lo graba igual en los descartados
            //  if (Saved == false) 
            //  { 
            // DiscardFilter(oMail, i);
            //  }

            //  /////-------- delarar salida de servidor: Marcado de correos descargados y cerraar conexion
            ////  PoClient
            #endregion

        }
        ////private void ServiFacturaFilter(Mail oMail,int i)
        ////{
        ////    String BodyAtt = String.Empty;
        ////    String[] attname = new string[30];
        ////    String[] StrBodyAtt = new string[30];
        ////    String FoldRouteAdj = @"\\10.177.1.230\mailacf$\ADJUNTOS"; //// ruta de adjuntos rptec
        ////    String SavePath = @"\\10.177.1.230\mailacf$\SERVIFACTURA" + "\\";  ////debe tener (\) al final

        ////    ReceiveMail RV = new ReceiveMail();
        ////    RV.SaveMail(oMail, i, SavePath);
        ////    SaveAttach(oMail, ref attname, FoldRouteAdj);
        ////    ReadAttach(ref attname,ref StrBodyAtt);
        ////    //GetXmlBody(BodyAtt);

        ////    /////SendMail         


        ////}
        //private static string _FormatHtmlTag(string src)
        //{
        //    src = src.Replace(">", "&gt;");
        //    src = src.Replace("<", "&lt;");
        //    return src;
        //}
        //private void BodyFilter(Mail oMail,int i)
        //{

        //    //BodyTextFormat btf = new BodyTextFormat();
        //    //oMail.DecodeTNEF();
        //    //string html = (oMail.HtmlBody);
        //    String oRut = string.Empty;
        //    //string html5 = btf.ToString;

        //        // Parse html body
        //    string html = oMail.HtmlBody;
        //    StringBuilder hdr = new StringBuilder();

        //    // Parse sender
        //    hdr.Append("<font face=\"Courier New,Arial\" size=2>");
        //    hdr.Append("<b>From:</b> " + _FormatHtmlTag(oMail.From.ToString()) + "<br>");

        //    // Parse to
        //    MailAddress[] addrs = oMail.To;
        //    int count = addrs.Length;
        //    if (count > 0)
        //    {
        //        hdr.Append("<b>To:</b> ");
        //        for (int k = 0; i < count; i++)
        //        {
        //            hdr.Append(_FormatHtmlTag(addrs[i].ToString()));
        //            if (k < count - 1)
        //            {
        //                hdr.Append(";");
        //            }
        //        }
        //        hdr.Append("<br>");
        //    }

        //    //String BodyEmail = string.Empty;
        //    String oFrom = (oMail.From).ToString();
        //    String oBody = (oMail.Content).ToString();
        //    String NomEmp = string.Empty;
        //    String Executive = string.Empty;
        //    ShopODestino(ref oFrom);
        //    string[] xRut = new string[5];
        //    int j=0; 
        //    oFrom = "facturaselectronicas";
        //    ReadBodyEmail(ref oBody,ref oRut, oFrom);
        //    xRut[j] = oRut;
        //    GetExecutive(oMail,Executive, xRut,NomEmp);   
            
        //}
        private void XmlFilter(Mail oMail,int i)
        {
            String WhoReceive=string.Empty, Subject = string.Empty, Body = string.Empty, Executive = string.Empty;
            String[] attname = new string[30];            
            String[] StrBodyAtt = new string[30];
            String[] RutEmi = new string[30];
            String[] Datos = new string[21];
            
            String oFrom = (oMail.From).ToString();
            //String RutEmi = string.Empty;
            ShopODestino(ref oFrom);
            String SavePath = @"\\10.177.1.230\mailacf$\XML" + "\\"+ oFrom + "\\";
            String FoldRouteAdj = @"\\10.177.1.230\mailacf$\ADJUNTOS"; //// ruta de adjuntos
            String NomEmp = string.Empty;
            ReceiveMail RV = new ReceiveMail();
            SendMail SM = new SendMail();
            String State = string.Empty;

            try
            {

                RV.SaveMail(oMail, i, ref SavePath);
                SaveAttach(oMail, ref attname, FoldRouteAdj);

                if (ValidateFileXml(attname) == true)
                {

                    ReadAttach(ref attname, ref StrBodyAtt);


                    if (!StrBodyAtt[0].Equals(""))
                    {
                        GetXmlBody(ref StrBodyAtt);
                        Body = StrBodyAtt[0];
                        GetXmlInfo(Body,ref Datos);
                        SaveXmlInfo(ref Datos);
                        //GetRutEmisorXML (StrBodyAtt, RutEmi);
                        //GetState(StrBodyAtt, ref State);
                        GetExecutive(oMail, ref Executive, RutEmi, ref NomEmp);
                        //SaveXmlInfo

                        Body = StrBodyAtt[0];
                        Subject = "Cesión Electronica: " + RutEmi[0] + " " + NomEmp;// + " ";// + State;
                        //SM.SenderMail(Executive, Subject, Body);
                        if (Executive != "" && NomEmp != "")
                        {
                            SM.SenderMail(Executive, Subject, Body); //envia el correo al ejecutivo , 
                        }
                        else
                        {
                            CopyMailToPendingFolder(SavePath);// Si no tiene ejecutivo  lo copia a la carpeta de pendientes,

                        }
                    }
                    
                }
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.XmlFilter", e.Message.ToString());
            }
            #region comentado1              


            //if (ValidateFileTxt(attname) == true)
            //{
            //    ReadAttach(ref attname, ref StrBodyAtt);
            //    GetRpetcInfo(BodyAtt, DatRptc);
            //    SaveRptcInfo(DatRptc);
            //    GetExecutive(oMail, ref Executive, xrut, ref NomEmp);

            //    if (Executive != "" && NomEmp != "")
            //    {
            //        SM.SenderMail(Executive, Subject, BodyAtt);
            //    }
            //    else
            //    {
            //        CopyMailToPendingFolder(SavePath);

            //    }
            //}
            //else
            //{

            //    ///log donde diga que el valor que se trato de parsear era xml

            //}
            #endregion

        }
        private void GetState(String[] BodyAtt,ref String State)
        {
           try
            {
                for (int i = 0; i < BodyAtt.Length; i++)
                {

                    if (BodyAtt[i] != null)
                    {
                        string strStart = "Estado del Envio       :";
                        string strEnd = "Detalle de Cesion Electronica de Credito";
                        String Boddy = BodyAtt[i];

                        int Start, End;
                        if (Boddy.Contains(strStart))
                        {
                            Start = Boddy.IndexOf(strStart, 0) + strStart.Length;
                            End = Boddy.IndexOf(strEnd, 0);
                            //End = Start + 13;
                            Boddy = Boddy.Substring(Start, End - Start);
                            State = (Boddy).Trim();
                            // return BodyAtt;
                        }
                        else
                        {
                            //return "";
                            break;
                        }
                    }
                    else
                    {
                        break;
                    }
                }
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetState", e.Message.ToString());
            }

        }
        /////////////////////////////metodos xml  GetTipoDteXML

        /// <summary>
        /// 
        /// </summary>
        /// <param name="BodyAtt"></param>
        /// <returns></returns>
        /// 
        private String GetTipoDteXML(String BodyAtt)
        {
            try
            {
                string strStart = "TipoDTE:";
                string strEnd = "RUTEmisor:";
                //String Boddy = BodyAtt[i];

                int Start, End;
                if (BodyAtt.Contains(strStart))
                {
                    Start = BodyAtt.IndexOf(strStart, 0) + strStart.Length;
                    End = BodyAtt.IndexOf(strEnd, 0);
                    //End = Start + 13;
                    BodyAtt = BodyAtt.Substring(Start, End - Start);
                    BodyAtt = (BodyAtt).Trim();
                    return BodyAtt;
                }
                else
                {
                    return "";
                }

            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetRutEmisorXML", e.Message.ToString());
                return "";
            }
        }

        private String GetRutEmisorXML(String BodyAtt)
        {
            try
            {                
                        string strStart = "Emisor:";
                        string strEnd = "RUTReceptor:";
                        //String Boddy = BodyAtt[i];

                        int Start, End;
                        if (BodyAtt.Contains(strStart))
                        {
                            Start = BodyAtt.IndexOf(strStart, 0) + strStart.Length;
                            End = BodyAtt.IndexOf(strEnd, 0);
                            //End = Start + 13;
                            BodyAtt = BodyAtt.Substring(Start, End - Start);
                            BodyAtt = (BodyAtt).Trim();
                            return BodyAtt;
                        }
                        else
                        {
                            return "";                            
                        }                     
                
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetRutEmisorXML", e.Message.ToString());
                return "";
            }            
        }
        private String GetRutRecepXML(String BodyAtt)
        {

            try
            {
                

                        string strStart = "RUTReceptor:";
                        string strEnd = "Folio:";
                        //String Boddy = BodyAtt[i];

                        int Start, End;
                        if (BodyAtt.Contains(strStart))
                        {
                            Start = BodyAtt.IndexOf(strStart, 0) + strStart.Length;
                            End = BodyAtt.IndexOf(strEnd, 0);
                            //End = Start + 13;
                            BodyAtt = BodyAtt.Substring(Start, End - Start);
                            BodyAtt = (BodyAtt).Trim();
                            return BodyAtt;
                        }
                        else
                        {
                            return "";                            
                        }                   
                    
                
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetRutRecepXML", e.Message.ToString());
                return "";
            }
        }
        private String GetFolioXML(String BodyAtt)
        {
            try
            {
                        string strStart = "Folio:";
                        string strEnd = "FchEmis:";
                        /////String Boddy = BodyAtt[i];

                        int Start, End;
                        if (BodyAtt.Contains(strStart))
                        {
                            Start = BodyAtt.IndexOf(strStart, 0) + strStart.Length;
                            End = BodyAtt.IndexOf(strEnd, 0);
                            //End = Start + 13;
                            BodyAtt = BodyAtt.Substring(Start, End - Start);
                            BodyAtt= (BodyAtt).Trim();
                            return BodyAtt;
                        }
                        else
                        {
                        return "";
                        }

            }catch (Exception e)
            {
                Logs.Log("FilterMail.GetFolioXML", e.Message.ToString());
                return "";
            }
        }
        private String GetFechEmisionXML(String BodyAtt)
        {
            try
            {
                        string strStart = "FchEmis:";
                        string strEnd = "MntTotal:";
                        ////String Boddy = BodyAtt[i];

                        int Start, End;
                        if (BodyAtt.Contains(strStart))
                        {
                            Start = BodyAtt.IndexOf(strStart, 0) + strStart.Length;
                            End = BodyAtt.IndexOf(strEnd, 0);
                            //End = Start + 13;
                            BodyAtt = BodyAtt.Substring(Start, End - Start);
                            BodyAtt = (BodyAtt).Trim();
                            return BodyAtt;
                        }
                        else
                        {
                            return "";                            
                        }
                   
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetFechEmisionXML", e.Message.ToString());
                return "";
            }
        }
        private String GetMontoXML(String BodyAtt)
        {
            try
            {
                        string strStart = "MntTotal:";
                        string strEnd = "Cedente:";
                       

                        int Start, End;
                        if (BodyAtt.Contains(strStart))
                        {
                            Start = BodyAtt.IndexOf(strStart, 0) + strStart.Length;
                            End = BodyAtt.IndexOf(strEnd, 0);
                            //End = Start + 13;
                            BodyAtt = BodyAtt.Substring(Start, End - Start);
                            BodyAtt = (BodyAtt).Trim();
                            return BodyAtt;
                        }
                        else
                        {
                            return "";                            
                        }
              
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetMontoXML", e.Message.ToString());
                return "";
            }
        }
        private String GetRutCedentXML(String BodyAtt)
        {
            try
            {               
                        string strStart = "RUT:";
                        string strEnd = "RazonSocial:";                        

                        int Start, End;
                        if (BodyAtt.Contains(strStart))
                        {
                            Start = BodyAtt.IndexOf(strStart, 0) + strStart.Length;
                            End = BodyAtt.IndexOf(strEnd, 0);
                            //End = Start + 13;
                            BodyAtt = BodyAtt.Substring(Start, End - Start);
                            BodyAtt = (BodyAtt).Trim();
                            return BodyAtt;
                        }
                        else
                        {
                            return "";
                        }                  
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetRutCedentXML", e.Message.ToString());
                return "";
            }
        }

        private String GetRznSocialXML(String BodyAtt)
        {
            try
            {
                        string strStart = "RazonSocial:";
                        string strEnd = "Direccion:";                        

                        int Start, End;
                        if (BodyAtt.Contains(strStart))
                        {
                            Start = BodyAtt.IndexOf(strStart, 0) + strStart.Length;
                            End = BodyAtt.IndexOf(strEnd, 0);
                             //End = Start + 13;
                            BodyAtt = BodyAtt.Substring(Start, End - Start);
                            BodyAtt = (BodyAtt).Trim();
                            return BodyAtt;
                        }
                        else
                        {
                            return "";
                           
                        }                  
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetRznSocialXML", e.Message.ToString());
                return "";
            }
        }
        /// <summary>
        /// //////////////////// GetMailCedent
        /// </summary>
        /// <param name="BodyAtt"></param>
        /// <returns></returns>
        /// 
        //private String GetMailCedent(String BodyAtt)
        //{
        //    try
        //    {
        //        string strStart1 = "Cedente:";
        //        string strStart2 = "eMail:";
        //        string strEnd = "RUTAutorizado:";

        //        int Start, End;
        //        if (BodyAtt.Contains(strStart1))
        //        {
        //            Start = BodyAtt.IndexOf(strStart1, 0) + strStart1.Length;
        //            End = BodyAtt.IndexOf(strEnd, 0);                    
        //            BodyAtt = BodyAtt.Substring(Start, End - Start);
        //            BodyAtt = (BodyAtt).Trim();
        //            return BodyAtt;
        //        }
        //        else
        //        {
        //            return "";
        //        }
        //    }
        //    catch (Exception e)
        //    {
        //        Logs.Log("FilterMail.GetRutEmi", e.Message.ToString());
        //        return "";
        //    }
        //}

        private String GetDireccCedentXML(String BodyAtt)
        {
            try
            {                
                        string strStart = "Direccion:";
                        string strEnd = "eMail:";                      

                        int Start, End;
                        if (BodyAtt.Contains(strStart))
                        {
                            Start = BodyAtt.IndexOf(strStart, 0) + strStart.Length;
                            End = BodyAtt.IndexOf(strEnd, 0);
                            //End = Start + 13;
                            BodyAtt = BodyAtt.Substring(Start, End - Start);
                            BodyAtt = (BodyAtt).Trim();
                            return BodyAtt;
                        }
                        else
                        {
                            return "";                            
                        }                  
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetRutEmi", e.Message.ToString());
                return "";
            }
        }

        private String GetEmailCedentXML(String BodyAtt)
        {
            try
            {                
                        string strStart = "eMail:";
                        string strEnd = "RUTAutorizado:";
                        
                        int Start, End;
                        if (BodyAtt.Contains(strStart))
                        {
                            Start = BodyAtt.IndexOf(strStart, 0) + strStart.Length;
                            End = BodyAtt.IndexOf(strEnd, 0);
                            //End = Start + 13;
                            BodyAtt = BodyAtt.Substring(Start, End - Start);
                            BodyAtt = (BodyAtt).Trim();
                            return BodyAtt;
                        }
                        else
                        {
                            return "";
                          
                        }
                 
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetEmailCedentXML", e.Message.ToString());
                return "";
            }
        }
        private String GetRutAutoXML(String BodyAtt)
        {
            try
            {               
                        string strStart1 = "RUTAutorizado:";
                        string strStart2 = "RUT:";
                        string strEnd = "Nombre:";
                       

                        int Start1, Start2, End;
                        if (BodyAtt.Contains(strStart2))
                        {
                            Start1 = BodyAtt.IndexOf(strStart1, 0) + strStart1.Length;
                            Start2 = BodyAtt.IndexOf(strStart2, Start1) + strStart2.Length;
                            End = BodyAtt.IndexOf(strEnd, Start1);
                            //End = Start + 13;
                            BodyAtt = BodyAtt.Substring(Start2, End - Start2);
                            BodyAtt = (BodyAtt).Trim();
                            return BodyAtt;
                        }
                        else
                        {
                            return "";                           
                        }                
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetRutAutoXML", e.Message.ToString());
                return ""; 
            }
        }
        private String GetNomAutoXML(String BodyAtt)
        {
            try
            {                
                       
                        string strStart = "Nombre:";
                        string strEnd = "DeclaracionJurada:";                       

                        int Start, End;
                        if (BodyAtt.Contains(strStart))
                        {
                            Start = BodyAtt.IndexOf(strStart, 0) + strStart.Length;
                            End = BodyAtt.IndexOf(strEnd, 0);
                            //End = Start + 13;
                            BodyAtt = BodyAtt.Substring(Start, End - Start);
                            BodyAtt = (BodyAtt).Trim();
                            return BodyAtt;
                        }
                        else
                        {
                            return "";                           
                        }                  
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetNomAutoXML", e.Message.ToString());
                return "";
            }
        }
        private String GetRutCesionarioXML(String BodyAtt)
        {
            try
            {
              
                        string strStart1 = "Cesionario:";
                        string strStart2 = "RUT:";
                        string strEnd = "RazonSocial:";                     
                                        
                        int Start1, Start2, End;
                        if (BodyAtt.Contains(strStart1))
                        {
                            Start1 = BodyAtt.IndexOf(strStart1, 0) + strStart1.Length;
                            Start2 = BodyAtt.IndexOf(strStart2, Start1) + strStart2.Length;
                            End = BodyAtt.IndexOf(strEnd, Start1);
                            //End = Start + 13;
                            BodyAtt = BodyAtt.Substring(Start2, End - Start2);
                            BodyAtt = (BodyAtt).Trim();
                            return BodyAtt;
                        }
                        else
                        {
                            return "";                          
                        }                
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetRutCesionarioXML", e.Message.ToString());
                return "";
            }
        }
        private String GetRznSocCesionarioXML(String BodyAtt)
        {
            try
            {                
                        string strStart1 = "Cesionario:";
                        string strStart2 = "RazonSocial:";
                        string strEnd = "Direccion:";                        

                        int Start1,Start2, End;
                        if (BodyAtt.Contains(strStart1))
                        {
                            Start1 = BodyAtt.IndexOf(strStart1, 0) + strStart1.Length;
                            Start2 = BodyAtt.IndexOf(strStart2, Start1) + strStart2.Length;
                            End = BodyAtt.IndexOf(strEnd, Start1);                    
                            BodyAtt = BodyAtt.Substring(Start2, End - Start2);
                            BodyAtt = (BodyAtt).Trim();
                            return BodyAtt;
                        }
                        else
                        {
                    return "";
                        }
              
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetRznSocCesionarioXML", e.Message.ToString());
                return "";  
            }
        }

       

        private String GetEmailCesionarioXML(String BodyAtt)
        {
            try
            {

                string strStart1 = "Cesionario:";
                string strStart2 = "eMail:";
                string strEnd = "MontoCesion:";

                int Start1,Start2, End;
                if (BodyAtt.Contains(strStart1))
                {
                    Start1 = BodyAtt.IndexOf(strStart1, 0) + strStart1.Length;
                    Start2 = BodyAtt.IndexOf(strStart2, Start1) + strStart2.Length;
                    End = BodyAtt.IndexOf(strEnd, Start1);
                    BodyAtt = BodyAtt.Substring(Start2, End - Start2);
                    BodyAtt = (BodyAtt).Trim();
                    return BodyAtt;
                }
                else
                {
                    return "";
                }
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetDireccCesionarioXML", e.Message.ToString());
                return "";
            }
        }
        private String GetDireccCesionarioXML(String BodyAtt)
        {
            try
            {
                
                        string strStart1 = "Cesionario:";
                        string strStart2 = "Direccion:";
                        string strEnd = "eMail:";                       

                        int Start1,Start2, End;
                        if (BodyAtt.Contains(strStart1))
                        {
                            Start1 = BodyAtt.IndexOf(strStart1, 0) + strStart1.Length;
                            Start2 = BodyAtt.IndexOf(strStart2, Start1) + strStart2.Length;
                            End = BodyAtt.IndexOf(strEnd, Start1);                    
                            BodyAtt = BodyAtt.Substring(Start2, End - Start2);
                            BodyAtt = (BodyAtt).Trim();
                            return BodyAtt;
                        }
                        else
                        {
                            return "";                          
                        }                   
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetDireccCesionarioXML", e.Message.ToString());
                return "";
            }
        }

        //private String GetMailCesionarioXML(String BodyAtt)
        //{
        //    try
        //    {
        //                string strStart1 = "Cesionario:";
        //                string strStart2 = "eMail:";
        //                string strEnd = "MontoCesion:";                       

        //                int Start, End;
        //                if (BodyAtt.Contains(strStart1))
        //                {
        //                    Start = BodyAtt.IndexOf(strStart1, 0) + strStart1.Length;
        //                    End = BodyAtt.IndexOf(strEnd, 0);                    
        //                    BodyAtt = BodyAtt.Substring(Start, End - Start);
        //                    BodyAtt = (BodyAtt).Trim();
        //                    return BodyAtt;
        //                }
        //                else
        //                {
        //                    return "";                           
        //                }                 
        //    }
        //    catch (Exception e)
        //    {
        //        Logs.Log("FilterMail.GetDireccCesionarioXML", e.Message.ToString());
        //        return "";
        //    }
        //}

        private String GetMontoCesionarioXML(String BodyAtt)
        {
            try
            {                
                        string strStart1 = "Cesionario:";
                        string strStart2 = "MontoCesion:";
                        string strEnd = "UltimoVencimiento:";                        

                        int Start1,Start2, End;
                        if (BodyAtt.Contains(strStart1))
                        {
                            Start1 = BodyAtt.IndexOf(strStart1, 0) + strStart1.Length;
                            Start2 = BodyAtt.IndexOf(strStart2, Start1) + strStart2.Length;
                            End = BodyAtt.IndexOf(strEnd, Start1);
                            BodyAtt = BodyAtt.Substring(Start2, End - Start2);
                            BodyAtt = (BodyAtt).Trim();
                            return BodyAtt;
                        }
                        else
                        {
                            return "";
                            
                        }                
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetDireccCesionarioXML", e.Message.ToString());
                return "";
            }
        }

        private String GetUltimoVencXML(String BodyAtt)
        {
            try
            {               
                      
                        string strStart = "UltimoVencimiento:";
                        string strEnd = "OtrasCondiciones";                       

                        int Start, End;
                        if (BodyAtt.Contains(strStart))
                        {
                            Start = BodyAtt.IndexOf(strStart, 0) + strStart.Length;
                            End = BodyAtt.IndexOf(strEnd, 0);
                            //End = Start + 13;
                            BodyAtt = BodyAtt.Substring(Start, End - Start);
                            BodyAtt = (BodyAtt).Trim();
                            return BodyAtt;
                        }
                        else
                        {
                            return "";                            
                        }                   
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetDireccCesionarioXML", e.Message.ToString());
                return "";
            }
        }

        private String GetMailDeudorXML(String BodyAtt)
        {
            try
            {
                        
                        string strStart = "eMailDeudor:";
                        string strEnd = "TmstCesion:";                       

                        int Start, End;
                        if (BodyAtt.Contains(strStart))
                        {
                            Start = BodyAtt.IndexOf(strStart, 0) + strStart.Length;
                            End = BodyAtt.IndexOf(strEnd, 0);
                            //End = Start + 13;
                            BodyAtt = BodyAtt.Substring(Start, End - Start);
                            BodyAtt = (BodyAtt).Trim();
                            return BodyAtt;
                        }
                        else
                        {
                            return "";                           
                        }
                  
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetDireccCesionarioXML", e.Message.ToString());
                return "";
            }
        }

        private String GetTmstCesionXML(String BodyAtt)
        {
            try
            {
                       
                        string strStart = "TmstCesion:";
                        //string strEnd = "TmstCesion:";
                        
                        int Start, End;
                        if (BodyAtt.Contains(strStart))
                        {
                            Start = BodyAtt.IndexOf(strStart, 0) + strStart.Length;
                            End = BodyAtt.IndexOf(strStart, 0) + strStart.Length+19;
                            //End = Start + 13;
                            BodyAtt = BodyAtt.Substring(Start, End - Start);
                            BodyAtt = (BodyAtt).Trim();
                            return BodyAtt;
                        }
                        else
                        {
                            return "";
                        }

            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetDireccCesionarioXML", e.Message.ToString());
                return "";
            }
        }
        private void GetExecutive(Mail oMail,ref String Executive,String[] xrut,ref String NomEmp)
        {
            try{

                SqlCommand cmd = new SqlCommand();
                SqlDataReader reader;

                string connString = ConfigurationManager.ConnectionStrings["myConnectionString"].ConnectionString;
                SqlConnection Connex = new SqlConnection(connString);

                Connex.Open();

                for (int i = 0; i < xrut.Length; i++)
                {
                    if (xrut[i] != null)
                    { 
                            if (Connex != null && Connex.State == ConnectionState.Open)
                            {

                            cmd.CommandText = "select top 1 dg_nombre_cliente,dbo.fnObtMailUsuario(dc_ejecutivo) as Mail  from BD_FACTORING.dbo.tb_operacion_factor_enc where dc_rut_cliente = '" + xrut[i] + "' order by df_creacion desc";
                            cmd.Connection = Connex;
                            reader = cmd.ExecuteReader();
                            reader.Read();

                                 if (reader.HasRows)
                                 {
                                 string NomClient = reader["dg_nombre_cliente"].ToString();
                                 string EmailExec = reader["Mail"].ToString();


                                NomEmp = NomClient;
                                Executive = EmailExec;
                                 }
                            }
                            else
                             {
                                     //error al log

                             }
                    }
                    else
                    {
                       break;   /// rompe el ciclo al ser null
                    }
                    
                }

                Connex.Close();

            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetExecutive", e.Message.ToString());
            }

        }                
        private void DicomFilter(Mail oMail, int i)
        {
            String[] attname = new string[10];
            String[] Datos = new string[10];
            //String FoldRouteAdj = @"\\10.177.1.230\mailacf$\ADJUNTOS"; //// ruta de adjuntos rptec
            String FoldRouteAdj = @"\\10.177.1.230\mailacf$\DICOM\PDF"; //// ruta de adjuntos rptec
            String SavePath = @"\\10.177.1.230\mailacf$\DICOM\";  ////debe tener (\) al final

            try
            {
                ReceiveMail RV = new ReceiveMail();
                RV.SaveMail(oMail, i, ref SavePath);
                SaveAttach(oMail, ref attname, FoldRouteAdj);
                GetDicomInfo(attname);
                //SaveDicomInfo(ref Datos);
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.DicomFilter", e.Message.ToString());
            }            
        }
        public void PendingProcess()
        {
            List<string> Files = new List<string>();
            String folderPath = @"\\10.177.1.230\mailacf$\RPETC_ACF\PENDIENTES";

            try
            {
                foreach (string file in Directory.EnumerateFiles(folderPath, "*.msg"))
                {
                    String Paths = Path.GetFullPath(file);
                    Files.Add(Paths);
                }
                PendingLoadMail(Files);
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.PendingProcess", e.Message.ToString());
            }           
            #region other
            //byte[] msgBytes = null;

            //foreach (string item in Files)
            //{               
            //    using (BinaryReader br = new BinaryReader(File.OpenRead("@"+item)))
            //    {
            //        FileInfo fi = new FileInfo("@" + item);
            //        msgBytes = br.ReadBytes((int)fi.Length);
            //    }
            //    MailMessage msg = new MailMessage();
            //    msg.LoadMessage(msgBytes);
            //    String Subject = (msg.Subject);
            //    String Body = (msg.BodyHtmlText);
            //    String From = (msg.From).ToString();            
            #endregion
        }
        private void PendingLoadMail(List<string> Files)
        {

            try
            {
                foreach (String Itm in Files)
                {
                    String FoldRoute = @"\\10.177.1.230\mailacf$\ADJUNTOS";
                    Mail oMail = new Mail("TryIt");
                    //String MailPath = @""+Itm;
                    oMail.LoadOMSG(Itm);    ////carga el msg como un objeto
                    String[] attname = new string[10];
                    String[] StrBodyAtt = new string[10];
                    String[] DatRptc = new string[23]; // soporta los 22 parametros de getrptsinfo
                    String[] xrut = new string[10];

                    String Executive = String.Empty;
                    String NomEmp = String.Empty;
                    String BodyAtt = String.Empty;

                    #region RTF Outlook Format Parsing And Decoding
                    //Attachment[] atts = oMail.Attachments;
                    //int count = atts.Length;
                    //string tempFolder = "c:\\temp";
                    //if (!System.IO.Directory.Exists(tempFolder))
                    //    System.IO.Directory.CreateDirectory(tempFolder);
                    //for (int i = 0; i < count; i++)
                    //{
                    //    Attachment att = atts[i];
                    //    //this attachment is in OUTLOOK RTF format (TNEF), decode it here.
                    //    if (String.Compare(att.Name, "winmail.dat") == 0)
                    //    {
                    //        Attachment[] tatts = null;
                    //        try
                    //        {
                    //            tatts = Mail.ParseTNEF(att.Content, true);
                    //        }
                    //        catch (Exception)
                    //        {
                    //            //MessageBox.Show(e.Message.ToString());
                    //            continue;
                    //        }
                    //        int y = tatts.Length;
                    //        for (int x = 0; x < y; x++)
                    //        {
                    //            Attachment tatt = tatts[x];
                    //            string tattname = String.Format("{0}\\{1}", tempFolder, tatt.Name);
                    //            tatt.SaveAs(tattname, true);
                    //        }
                    //        continue;
                    //    }
                    //    string attname = String.Format("{0}\\{1}", tempFolder, att.Name);
                    //    att.SaveAs(attname, true);

                    //    ///llamar

                    //}

                    #endregion
                    StringBuilder hdr = new StringBuilder();
                    Attachment[] atts = oMail.Attachments;
                    int count = atts.Length;
                    //string tempFolder = "c:\\temp";
                    if (!System.IO.Directory.Exists(FoldRoute))
                        Directory.CreateDirectory(FoldRoute);

                    for (int i = 0; i < count; i++)
                    {
                        Attachment att = atts[i];
                        attname[i] = String.Format("{0}\\{1}", FoldRoute, att.Name);
                        att.SaveAs(attname[i], true);
                    }
                    attname = attname.Where(c => c != null).ToArray(); /// limpia los valores nulos

                    string tx = ".txt";
                    string xml = ".xml";
                    String State = string.Empty;
                    //String  = string.Empty;
                    String[] RutEmi = new string[10];
                    String[] Datos = new string[30];
                    String Subject = String.Empty;

                    int sup = attname.Length;
                    for (int k = 0; k < sup; k++)
                    //foreach (String s in attname)
                    {
                        //if (attname[k]==("*.txt"))
                        //if (attname.Contains(txt))

                        if (attname[k] != null)
                        {
                            if (attname[k].IndexOf(tx, 0) != -1)// || attname[k].IndexOf(xml, 0) != -1)
                            {

                                ReadAttach(ref attname, ref StrBodyAtt);
                                StrBodyAtt = StrBodyAtt.Where(c => c != null).ToArray(); /// limpia los valores nulos

                                int value = StrBodyAtt.Length;
                                for (int i = 0; i < value; i++)
                                {
                                    if (StrBodyAtt[i] != null)
                                    {

                                        BodyAtt = StrBodyAtt[i];
                                        GetRpetcInfo(BodyAtt, DatRptc);

                                    }

                                }
                                SendMail sn = new SendMail();

                                ValidateDates(ref DatRptc);  /////ordena las fechas de los rptcs para que pasen bien en el insert del sp
                                SaveRptcInfo(DatRptc);
                                //GetState(StrBodyAtt,ref State);

                                //string Subject = "Cesión Electronica: " + DatRptc[0] + " " + DatRptc[8];
                                //string Subject = "Cesión Electronica " + DatRptc[0] + " " + DatRptc[3] + " " + State;
                                xrut[0] = DatRptc[7];
                                GetExecutive(oMail, ref Executive, xrut, ref NomEmp);
                                Subject = "Cesión Electronica: " + DatRptc[0] + " " + NomEmp;
                               // sn.SenderMail(Executive, Subject, BodyAtt);  ///clase que envia el correo

                                String fileName = System.IO.Path.GetFullPath(Itm.ToString());
                                //
                                if (sn.SenderMail(Executive, Subject, BodyAtt)==true)
                                {
                                    PendingDeleteLoadMail(fileName);  ///elimina de los pendientes si es que se envio correo a ejecutivo
                                }else
                                {
                                    Logs.Log("FilterMail.PendingLoadMail", "No se elimino porque aun no tiene Ejecutivo Asignado");
                                }
                                

                            }
                            else if (attname[k].IndexOf(xml, 0) != -1)
                            {
                                ReadAttach(ref attname, ref StrBodyAtt);


                                int value = StrBodyAtt.Length;
                                for (int i = 0; i < value; i++)
                                {
                                    if (StrBodyAtt[i] != null)
                                    {
                                        //GetXmlBody(ref StrBodyAtt);

                                        ////GetRutEmisorXML(StrBodyAtt, RutEmi);

                                        //GetExecutive(oMail, ref Executive, RutEmi, ref NomEmp);


                                        if (!StrBodyAtt[0].Equals(""))
                                        {
                                            GetXmlBody(ref StrBodyAtt);
                                            BodyAtt = StrBodyAtt[0];
                                            GetXmlInfo(BodyAtt, ref DatRptc);
                                            SaveXmlInfo(ref DatRptc);
                                            //GetRutEmisorXML (StrBodyAtt, RutEmi);
                                            //GetState(StrBodyAtt, ref State);
                                            GetExecutive(oMail, ref Executive, RutEmi, ref NomEmp);
                                            //SaveXmlInfo

                                            BodyAtt = StrBodyAtt[0];
                                            Subject = "Cesión Electronica: " + RutEmi[0] + " " + NomEmp;// + " ";// + State;
                                                                                                        //SM.SenderMail(Executive, Subject, Body);
                                            SendMail SM = new SendMail();
                                            if (Executive != "" && NomEmp != "")
                                            {
                                                SM.SenderMail(Executive, Subject, BodyAtt); //envia el correo al ejecutivo , 
                                            }
                                            //else
                                            //{
                                            //    CopyMailToPendingFolder(SavePath);// Si no tiene ejecutivo  lo copia a la carpeta de pendientes,

                                            //}
                                        }


                                        //Body = StrBodyAtt[0];
                                        //Subject = "Cesión Electronica " + RutEmi[0] + " " + NomEmp + " " + State;
                                        //SM.SenderMail(Executive, Subject, Body);

                                    }

                                }
                                SaveRptcInfo(DatRptc);
                                GetExecutive(oMail, ref Executive, xrut, ref NomEmp);

                            }

                        }

                    }
                    #region Comentarios utiles
                    //}

                    //Attachment[] atts = oMail.Attachments;
                    //int count = atts.Length;
                    //String[] attname = new String[30];

                    ////if (count > 0)
                    ////{
                    ////    if (!Directory.Exists(FoldRoute))
                    ////        Directory.CreateDirectory(FoldRoute);

                    ////    hdr.Append("<b>Attachments:</b>");
                    ////    for (int i = 0; i < count; i++)
                    ////    {
                    ////        Attachment att = atts[i];

                    ////        attname[i] = String.Format("{0}\\{1}", FoldRoute, att.Name);
                    ////        att.SaveAs(attname[i], true);
                    ////        hdr.Append(String.Format("<a href=\"{0}\" target=\"_blank\">{1}</a> ",
                    ////                attname, att.Name));

                    ////        //if (att.ContentID.Length > 0)
                    ////        //{

                    ////        //    //// Show embedded images.
                    ////        //    html = html.Replace("cid:" + att.ContentID, attname[n]);
                    ////        //}
                    ////        //else if (String.Compare(att.ContentType, 0, "image/", 0,
                    ////        //            "image/".Length, true) == 0)
                    ////        //{

                    ////        //    ////// show attached images.
                    ////        //    html = html + String.Format("<hr><img src=\"{0}\">", attname);
                    ////        //}
                    ////    }
                    ////}

                    // public void PendingParseAttachment()
                    //// {
                    //     Mail oMail = new Mail("TryIt");
                    //     oMail.Load("c:\\test.eml", false);

                    //// Decode winmail.dat (TNEF) automatically
                    //oMail.DecodeTNEF();

                    #endregion
                }

            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.PendingLoadMail", e.Message.ToString());
            }

            
        }
        private void PendingDeleteLoadMail(string fileName)
        {

            if (File.Exists(fileName))
            {
                System.IO.File.Delete(fileName);

            }
            else
            {
                Logs.Log("FilterMail:PendingDeleteLoadMail", "Error - No se encontro Ruta de eliminacion");

            } 

        }


        private void DiscardFilter(Mail oMail, int i)
        {
            try
            {
                String SavePath = @"\\10.177.1.230\mailacf$\DESCARTADOS" + "\\";   ////debe tener (\) al final
                ReceiveMail RV = new ReceiveMail();
                RV.SaveMail(oMail, i, ref SavePath);
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.DiscardFilter", e.Message.ToString());
            }

                        
        }


        private void GetXmlBody(ref String[] BodyAtt)
        {
            int largo = BodyAtt.Length;
            string[] stringSeparators = new string[] { "\n" };
            string strStart = "<IdDTE>";
            string SearchEnd = "</DocumentoCesion>";
            //int i = 0;

            try
            {
                for (int i = 0; i < largo; i++)
                {
                    if (BodyAtt[i] != null)
                    {
                        /// eliminar los tags de los XML
                        //Line = Environment.NewLine;
                        //String Vacio = string.Empty;
                        int Start, End;
                        String SBodyAtt = BodyAtt[i];
                        Start = SBodyAtt.IndexOf(strStart, 0) + strStart.Length;
                        End = SBodyAtt.IndexOf(SearchEnd, 0); ////+ SearchEnd.Length;
                        String NewBody = SBodyAtt.Substring(Start, End - Start);

                        NewBody = Regex.Replace(NewBody, "</.*?>", string.Empty);
                        NewBody = NewBody.Replace("<", System.Environment.NewLine);
                        NewBody = Regex.Replace(NewBody, ">", ":");

                        BodyAtt[i] = NewBody;

                        i++;   //incremento de i                
                    }
                    else if (BodyAtt[i] == null)
                    {
                        break;
                    }
                }
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetXmlBody", e.Message.ToString());
            }     
        }              
        private void GetDicomInfo(String[] attname)
        {
            string strStart = "DP__";
            string xRut = "";
            string strEnd = "_";
            string strEnd2 = ".PDF";
            //int n = 0;
            
            string[] HoraFinal=new string[3];
            string[] FechaFinal = new string[3];
            String[] Datos;
            String[] All= new String[5];
            String[] Datos2;
            String[] NoNull;
            string[] stringSeparators = new string[] { "_" };
            string[] gn= new string[] { "_" };

            // 
            NoNull = attname.Where(c => c != null).ToArray();
            String Filename = System.IO.Path.GetFileName(NoNull[0].ToString());
            String FileRoute = NoNull[0];
            
            try
            {

                int ln = NoNull.Length;
                for(int i=0; i<ln; i++)
                {
                    int Start, End, End2;
                    if (Filename.Contains(strStart)&& Filename.Contains(strEnd2))
                    {
                        Start = NoNull[0].IndexOf(strStart, 0) + strStart.Length;
                        End = NoNull[0].IndexOf(strEnd, Start);
                        xRut = NoNull[0].Substring(Start, End - Start);
                        End2 = NoNull[0].IndexOf(strEnd2, Start);
                        NoNull[0] = NoNull[0].Substring(End, End2 - End);
                        Datos = NoNull[0].Split(stringSeparators, 2, StringSplitOptions.RemoveEmptyEntries);
                        Datos2 = Datos[1].Split(gn, 2, StringSplitOptions.RemoveEmptyEntries);

                        string xfecha = Datos[0];
                        string xhora = Datos2[0];
                        string id = Datos2[1];

                        FechaFinal[0] = xfecha.Substring(0, 4);
                        FechaFinal[1] = xfecha.Substring(4, 2);
                        FechaFinal[2] = xfecha.Substring(6, 2);

                        xfecha = FechaFinal[2] + "-" + FechaFinal[1] + "-" + FechaFinal[0];

                        HoraFinal[0] = xhora.Substring(0, 2);
                        HoraFinal[1] = xhora.Substring(2, 2);
                        HoraFinal[2] = xhora.Substring(4, 2);
                        xhora = HoraFinal[0] + ":" + HoraFinal[1] + ":" + HoraFinal[2];
                        All[0] = xRut;
                        All[1] = xfecha;
                        All[2] = xhora;
                        All[3] = id;
                        All[4] = FileRoute;

                        SaveDicomInfo(All);
                        ////DP__91835754_20140618_092355_0001941732.PDF

                    }

                }
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetDicomInfo", e.Message.ToString());
            }               

        }
        private void RptcFilter(Mail oMail,int i)
        {
            String[] attname = new String[10];
            String[] DatRptc = new String[23];
            String[] StrBodyAtt = new String[10];
            String Executive = String.Empty;
            String NomEmp = String.Empty;
            String Subject = String.Empty;
            String[] xrut = new String[10];
            String vc = string.Empty;
            String BodyAtt = String.Empty;
            String FoldRouteAdj = @"\\10.177.1.230\mailacf$\ADJUNTOS"; //// ruta de adjuntos rptec
            String SavePath = @"\\10.177.1.230\mailacf$\RPETC_ACF" + "\\";  ////debe tener (\) al final
            ReceiveMail RV = new ReceiveMail();
            SendMail SM = new SendMail();

            try
            {
                RV.SaveMail(oMail, i, ref SavePath);
                SaveAttach(oMail, ref attname, FoldRouteAdj);
                ////validar si es rptc


                if (ValidateFileTxt(attname) == true)
                {
                    ReadAttach(ref attname, ref StrBodyAtt);
                    BodyAtt = StrBodyAtt[0];
                    GetRpetcInfo(BodyAtt, DatRptc);

                    ValidateDates(ref DatRptc);
                    if (AskFileExistOnTableRptc(DatRptc) == false)
                    {
                        SaveRptcInfo(DatRptc);
                        xrut[0] = DatRptc[7];
                        GetExecutive(oMail, ref Executive, xrut, ref NomEmp);
                        Subject = "Cesión Electronica: " + xrut[0] + " " + NomEmp;// + " ";// + State;
                                                                                  //Executive = string.Empty;
                        if (Executive != "" && NomEmp != "")
                        {
                            SM.SenderMail(Executive, Subject, BodyAtt); //envia el correo al ejecutivo , 
                        }
                        else
                        {
                            CopyMailToPendingFolder(SavePath);// Si no tiene ejecutivo  lo copia a la carpeta de pendientes,

                        }


                    }
                    else
                    {
                        Logs.Log("RptcFilter.AskFileExistOnTable", "Se trato de guardar un valor que ya existe en la BD");
                    }


                }
                else
                {
                    ///log donde diga que el valor que se trato de parsear era xml
                }  

            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.RptcFilter", e.Message.ToString());
            }

        }

        private void ValidateDates(ref String[]datos)
        {
            String Date4 = datos[2];
            String Date1 = datos[11];
            String Date2 = datos[18];
            String Date3 = datos[21];

            try
            {
                string[] Fech5 = Date4.Split(' ');
                string[] Fech6 = Fech5[0].Split('/');
                string[] Fech1 = Date1.Split('-');
                string[] Fech2 = Date2.Split(' ');
                string[] Fech3 = Fech2[0].Split('-');
                string[] Fech4 = Date3.Split('-');

                datos[2] = Fech6[0] + "-" + Fech6[1] + "-" + Fech6[2] + " " + Fech5[1];
                datos[11] = Fech1[2] + "-" + Fech1[1] + "-" + Fech1[0];                
                datos[18] = Fech1[2] + "-" + Fech3[1] + "-" + Fech3[0] + " " + Fech2[1];
                datos[21] = Fech4[2] + "-" + Fech4[1] + "-" + Fech4[0];
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.ValidateDates", e.Message.ToString());
            }
            //////('20160602', '20160602') '23-05-2016 09:50:02' NO VALIDO
            //////('21/06/2016', '21/06/2016')
            //////('21-06-2016', '21-06-2016')            
        }

        private Boolean ValidateFileXml(String[]att)
        {
            var xml = ".txt";

            try
            {
                att = att.Where(c => c != null).ToArray(); /// limpia los valores nulos

                int ed = att.Length;
                for (int i = 0; i < ed; i++)
                {
                    bool has;
                    if (has = att[i].Contains(xml))
                    {
                        return false;
                    }
                }
                return true;
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.ValidateFileXml", e.Message.ToString());
                return false;
            }
        }
        private Boolean ValidateFileTxt(String[] att)
        {            
            var xml = ".xml";

            try
            {
                att = att.Where(c => c != null).ToArray(); /// limpia los valores nulos

                int ed = att.Length;
                for (int i = 0; i < ed; i++)
                {
                    bool has;
                    if (has = att[i].Contains(xml))
                    {
                        return false;
                    }
                }
                return true;
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.ValidateFileTxt", e.Message.ToString());
                return false;
            }            
        }
        
        private void CopyMailToPendingFolder(String sourcePath)
        {
            string targetPath = @"\\10.177.1.230\mailacf$\RPETC_ACF\PENDIENTES";//+  "\\";
            string fileName = System.IO.Path.GetFileName(sourcePath);


            string destFile = System.IO.Path.Combine(targetPath, fileName);
            try
            {
                if (System.IO.Directory.Exists(targetPath))
                {
                    System.IO.File.Copy(sourcePath, destFile, true);
                }
                else
                {
                    Logs.Log("FilterMail.CopyMailToPendingFolder", "No se logra obtener la ruta");
                }
            }
            catch (Exception e)
            {   
                Logs.Log("FilterMail.ValidateFileTxt", e.Message.ToString());                
            }
            

        }
        private void BackUpFilter(Mail oMail,int i)
        {
            //String attname = "";
            //String FoldRouteAdj = @"\\10.177.1.230\mailacf$\ADJUNTOS"; //// ruta de adjuntos rptec          
            String SavePath = @"\\10.177.1.230\mailacf$\RESPALDOS" + "\\"; ////debe tener (\) al final
            String Ofrom = (oMail.From).ToString();            
            ShopEmail(ref Ofrom);
            Ofrom = Ofrom + "\\";

            SavePath = SavePath + Ofrom;

            try
            {
                ReceiveMail RV = new ReceiveMail();                

                RV.SaveMail(oMail, i, ref SavePath);
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.BackUpFilter", e.Message.ToString());
            }           
            //SaveAttach(oMail, ref attname, FoldRouteAdj);            
        }   
        private void ReadBodyEmail(ref String BodyEmail,ref String xRut,String oFrom)
        {
            try
            {
                if (oFrom == "facturaselectronicas")
                {
                    string Ndoc = "Nº Documento:";
                    string Nrut = "Rut: ";
                    string Rzon = "Razón Social:";

                    int Nd, Nr, Rz;

                    Nd = BodyEmail.IndexOf(Ndoc, 0) + Ndoc.Length;
                    Nr = BodyEmail.IndexOf(Nrut, 0);
                    Rz = BodyEmail.IndexOf(Rzon, 0);
                    Ndoc = BodyEmail.Substring(Nd, Nr - Nd);
                    Nr = BodyEmail.IndexOf(Nrut, 0) + Nrut.Length; ;
                    Nrut = BodyEmail.Substring(Nr, Rz - Nr);
                    xRut = Nrut;
                }
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.ReadBodyEmail", e.Message.ToString());
            }            
        }
        private void ReadAttach (ref String[] attname,ref String[] StrBodyAtt)
        {
            string tx = ".txt";
            string xml = ".xml";

            try
            {
                for (int i = 0; i < attname.Length; i++)
                {
                    if (attname[i] != null)
                    {
                        if (attname[i].IndexOf(tx, 0) != -1 || attname[i].IndexOf(xml, 0) != -1)
                        {
                            var fileStream = new FileStream(attname[i], FileMode.Open, FileAccess.Read);
                            using (var streamReader = new StreamReader(fileStream))
                            {
                                //if (streamReader.ReadToEnd() != null)
                                //{

                                String Value = streamReader.ReadToEnd();
                                StrBodyAtt[i] = Value;

                                StrBodyAtt = StrBodyAtt.Where(c => c != null).ToArray();

                            }

                        }
                    }
                    else
                    {
                        //Array.Sort(StrBodyAtt);
                        StrBodyAtt = StrBodyAtt.Where(c => c != null).ToArray();
                        break;
                    }

                }
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.ReadAttach", e.Message.ToString());
            }                      
           
        }
        private void SaveAttach(Mail oMail,ref String[] attname, String FoldRoute)
        {

            try {
                StringBuilder hdr = new StringBuilder();
                int n = 0;
                System.Globalization.CultureInfo cur = new
                            System.Globalization.CultureInfo("en-US");
                string sdate = Today.ToString("yyyyMMdd", cur);

                string html = oMail.HtmlBody;
                FoldRoute = FoldRoute + "\\" + sdate;

                // Parse attachments and save to local folder
                Attachment[] atts = oMail.Attachments;
                int count = atts.Length;
                if (count > 0)
                {
                    if (!Directory.Exists(FoldRoute))
                        Directory.CreateDirectory(FoldRoute);

                    hdr.Append("<b>Attachments:</b>");
                    for (int i = 0; i < count; i++)
                    {
                        Attachment att = atts[i];

                        attname[n] = String.Format("{0}\\{1}", FoldRoute, att.Name);
                        att.SaveAs(attname[n], true);
                        hdr.Append(String.Format("<a href=\"{0}\" target=\"_blank\">{1}</a> ",
                                attname, att.Name));
                        #region attcommented

                        //if (att.ContentID.Length > 0)
                        //{

                        //    //// Show embedded images.
                        //    html = html.Replace("cid:" + att.ContentID, attname[n]);
                        //}
                        //else if (String.Compare(att.ContentType, 0, "image/", 0,
                        //            "image/".Length, true) == 0)
                        //{

                        //    ////// show attached images.
                        //    html = html + String.Format("<hr><img src=\"{0}\">", attname);
                        //}
                        #endregion
                    }
                }
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.SaveAttach", e.Message.ToString());
            }
        }
        private void ShopODestino(ref String oFrom)
        {

            try {

                //Corta el correo solo deja nombre son @ ni dominio(Criz@dominio==Criz)      
                int Pos2 = 1 + (oFrom.IndexOf("<"));
                oFrom = oFrom.Substring(Pos2);
                int Pos = oFrom.IndexOf("@");
                oFrom = oFrom.Substring(0, Pos);
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.ShopODestino", e.Message.ToString());
            }
        }
        private void ShopEmail(ref String oFrom)
        {

            try
            {
                //Corta el correo solo deja nombre del emai standar(Criz@dominio)      
                int Pos2 = 1 + (oFrom.IndexOf("<"));
                oFrom = oFrom.Substring(Pos2);
                int Pos = oFrom.IndexOf(">");
                oFrom = oFrom.Substring(0, Pos);
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.ShopODestino", e.Message.ToString());
            }           

        }

        private void GetXmlInfo(String BodyAtt,ref String[] Datos)
        {

            // String[] Datos = new string[30];
            try
            {
                Datos[0] = GetTipoDteXML(BodyAtt);
                Datos[1] = GetRutEmisorXML(BodyAtt); 
                Datos[2] = GetRutRecepXML(BodyAtt); 
                Datos[3] = GetFolioXML(BodyAtt); 
                Datos[4] = GetFechEmisionXML(BodyAtt);
                Datos[5] = GetMontoXML(BodyAtt); 
                Datos[6] = GetRutCedentXML(BodyAtt); 
                Datos[7] = GetRznSocialXML(BodyAtt);
                Datos[8] = GetDireccCedentXML(BodyAtt);
                Datos[9] = GetEmailCedentXML(BodyAtt);                 
                Datos[10] = GetRutAutoXML(BodyAtt);
                Datos[11] = GetNomAutoXML(BodyAtt);
                Datos[12] = GetDeclJuradaXML(BodyAtt);
                Datos[13] = GetRutCesionarioXML(BodyAtt);               
                Datos[14] = GetRznSocCesionarioXML(BodyAtt);
                Datos[15] = GetDireccCesionarioXML(BodyAtt);
                Datos[16] = GetEmailCesionarioXML(BodyAtt);
                Datos[17] = GetMontoCesionarioXML(BodyAtt);
                Datos[18] = GetUltimoVencXML(BodyAtt);
                Datos[19] = GetMailDeudorXML(BodyAtt);                
                Datos[20] = GetTmstCesionXML(BodyAtt);
                
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetXmlInfo", e.Message.ToString());
            }
        }
        private void GetRpetcInfo(String BodyAtt,String[]Datos)
        {                       
           try
            {
                Datos[0] = GetRutEnvio(BodyAtt);
                Datos[1] = GetIdentEnvio(BodyAtt);
                Datos[2] = GetFechaRecep(BodyAtt);
                Datos[3] = GetEstEnvio(BodyAtt);
                Datos[4] = GetTipDoc(BodyAtt);
                Datos[5] = GetNroDoc(BodyAtt);
                Datos[6] = GetMonTotal(BodyAtt);
                Datos[7] = GetRutEmisor(BodyAtt);
                Datos[8] = GetNomEmisor(BodyAtt);
                Datos[9] = GetRutReceptor(BodyAtt);
                Datos[10] = GetNomReceptor(BodyAtt);
                Datos[11] = GetDocCedFechaEmision(BodyAtt);
                Datos[12] = GetRutCedente(BodyAtt);
                Datos[13] = GetNomCedente(BodyAtt);
                Datos[14] = GetDirecCedente(BodyAtt);
                Datos[15] = GetEmailCedente(BodyAtt);
                Datos[16] = GetRutCesionario(BodyAtt);
                Datos[17] = GetNombreCesionario(BodyAtt);
                Datos[18] = GetFechaCesion(BodyAtt);
                Datos[19] = GetDeclJurada(BodyAtt);
                Datos[20] = GetMontoCedido(BodyAtt);
                Datos[21] = GetUltimoVenc(BodyAtt);
                Datos[22] = GetIdentCesion(BodyAtt);
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetRpetcInfo", e.Message.ToString().ToString());
            }
            #region Commented 01
            //String RutEnvio = String.Empty;
            //String IdentEnvio = String.Empty;
            //String FechaRecep = String.Empty;
            //String EstEnvio = String.Empty;
            //String TipDoc = String.Empty;
            //String NroDoc = String.Empty;
            //String NroFact = String.Empty;
            //String MonTotal = String.Empty;
            //String RutEmisor = String.Empty;
            //String NomEmisor = String.Empty;
            //String RutReceptor = String.Empty;
            //String NomReceptor = String.Empty;
            //String DocCedFechaEmision = String.Empty;
            //String RutCedente = String.Empty;
            //String NomCedente = String.Empty;
            //String DirecCedente = String.Empty;
            //String EmailCedente = String.Empty;
            //String RutCesionario = String.Empty;
            //String NombreCesionario = String.Empty;
            //String FechaCesion = String.Empty;
            //String DeclJurada = String.Empty;
            //String MontoCedido = String.Empty;
            //String UltimoVenc = String.Empty;
            //String IdentCesion = String.Empty;

            //RutEnvio = GetRutEnvio(BodyAtt);
            //IdentEnvio = GetIdentEnvio(BodyAtt);
            //FechaRecep = GetFechaRecep(BodyAtt);
            //EstEnvio = GetEstEnvio(BodyAtt);
            //TipDoc = GetTipDoc(BodyAtt);
            //NroDoc = GetNroDoc(BodyAtt);
            ////NroFact = GetNroFact(BodyAtt);
            //MonTotal = GetMonTotal(BodyAtt);
            //RutEmisor = GetRutEmisor(BodyAtt);
            //NomEmisor = GetNomEmisor(BodyAtt);
            //RutReceptor = GetRutReceptor(BodyAtt);
            //NomReceptor = GetNomReceptor(BodyAtt);
            //DocCedFechaEmision = GetDocCedFechaEmision(BodyAtt);            
            //RutCedente = GetRutCedente(BodyAtt);
            //NomCedente = GetNomCedente(BodyAtt);
            //DirecCedente = GetDirecCedente(BodyAtt);
            //EmailCedente = GetEmailCedente(BodyAtt);
            //RutCesionario = GetRutCesionario(BodyAtt);
            //NombreCesionario = GetNombreCesionario(BodyAtt);
            //FechaCesion = GetFechaCesion(BodyAtt);
            //DeclJurada = GetDeclJurada(BodyAtt);
            //MontoCedido = GetMontoCedido(BodyAtt);
            //UltimoVenc = GetUltimoVenc(BodyAtt);
            //IdentCesion = GetIdentCesion(BodyAtt);
            #endregion
        }
        private void SaveRptcInfo(String[] Datos)
        {
            try
            {

                SqlCommand cmd = new SqlCommand();
                string connString = ConfigurationManager.ConnectionStrings["myConnectionString"].ConnectionString;
                SqlConnection Connex = new SqlConnection(connString);

                Connex.Open();
                if (Connex != null && Connex.State == ConnectionState.Open)
                {



                    StringBuilder sb = new StringBuilder();

                    














                    cmd = new SqlCommand("Sp_insertar_Rptcinfo");
                    cmd.CommandType = CommandType.StoredProcedure;


                    DateTime df_FechaRecep = DateTime.Parse(Datos[2]);
                    int dn_NroDoc = int.Parse(Datos[5]);
                    int dn_MonTotal = int.Parse(Datos[6]);
                    DateTime df_DocCedFechaEmision = DateTime.Parse(Datos[11]);
                    DateTime df_FechaCesion = DateTime.Parse(Datos[18]);
                    int dn_MontoCedido = int.Parse(Datos[20]);
                    DateTime df_UltimoVenc = DateTime.Parse(Datos[21]);

                    ////----	[df_FechaRecep]////        [datetime]
                    ////----	[dn_NroDoc]    ////        [int] NULL,
                    ////----	[dn_MonTotal]////        [int] NULL,
                    ////----	[df_DocCedFechaEmision]////        [datetime]
                    ////----	[df_FechaCesion]////        [datetime]
                    ////----	[dn_MontoCedido]////        [int] NULL,
                    ////----	[df_UltimoVenc]////        [datetime]
                    ////----	[df_FechaReg]////        [datetime]-----pasa sin nada es nulo, getdate();

                    cmd.Parameters.AddWithValue("@dc_RutEnvio", Datos[0]);
                    cmd.Parameters.AddWithValue("@dc_IdentEnvio", Datos[1]);
                    cmd.Parameters.AddWithValue("@df_FechaRecep", df_FechaRecep);
                    cmd.Parameters.AddWithValue("@dg_EstEnvio", Datos[3]);
                    cmd.Parameters.AddWithValue("@dn_TipDoc", Datos[4]);
                    cmd.Parameters.AddWithValue("@dn_NroDoc", dn_NroDoc);
                    cmd.Parameters.AddWithValue("@dn_MonTotal", dn_MonTotal);
                    cmd.Parameters.AddWithValue("@dc_RutEmisor", Datos[7]);
                    cmd.Parameters.AddWithValue("@dg_NomEmisor", Datos[8]);
                    cmd.Parameters.AddWithValue("@dc_RutReceptor", Datos[9]);
                    cmd.Parameters.AddWithValue("@dg_NomReceptor", Datos[10]);
                    cmd.Parameters.AddWithValue("@df_DocCedFechaEmision", df_DocCedFechaEmision);
                    cmd.Parameters.AddWithValue("@dc_RutCedente", Datos[12]);
                    cmd.Parameters.AddWithValue("@dg_NomCedente", Datos[13]);
                    cmd.Parameters.AddWithValue("@dg_DirecCedente", Datos[14]);
                    cmd.Parameters.AddWithValue("@dg_EmailCedente", Datos[15]);
                    cmd.Parameters.AddWithValue("@dg_RutCesionario", Datos[16]);
                    cmd.Parameters.AddWithValue("@dg_NomCesionario", Datos[17]);
                    cmd.Parameters.AddWithValue("@df_FechaCesion", df_FechaCesion);
                    cmd.Parameters.AddWithValue("@dg_DeclJurada", Datos[19]);
                    cmd.Parameters.AddWithValue("@dn_MontoCedido", dn_MontoCedido);
                    cmd.Parameters.AddWithValue("@df_UltimoVenc", df_UltimoVenc);
                    cmd.Parameters.AddWithValue("@dc_IdentCesion", Datos[22]);
                    cmd.Parameters.AddWithValue("@df_FechaReg", ""); //getdate()
                    cmd.Connection = Connex;
                    cmd.ExecuteNonQuery();

                    Connex.Close();

                }
                else
                {
                    Logs.Log("FilterMail.SaveRptcInfo", "Error en la Conexion, o algo fallo en el StringConnection");
                }

            }
            catch (Exception e)
            {

                //if(e.Message="")
                Logs.Log("FilterMail.GetRpetcInfo", e.Message.ToString().ToString());
            }
        }

        private void SaveXmlInfo(ref String[]Datos)
        {
            try
            {
                SqlCommand cmd = new SqlCommand();
                string connString = ConfigurationManager.ConnectionStrings["myConnectionString"].ConnectionString;
                SqlConnection Connex = new SqlConnection(connString);

                Connex.Open();
                if (Connex != null && Connex.State == ConnectionState.Open)
                {
                    cmd = new SqlCommand("Sp_insertar_XmlInfo");
                    cmd.CommandType = CommandType.StoredProcedure;

                    DateTime Fech_Emision = DateTime.Parse(Datos[4]);
                    DateTime UltimoVenc = DateTime.Parse(Datos[18]);
                    DateTime TmsCesion = DateTime.Parse(Datos[20]);

                    cmd.Parameters.AddWithValue("@dn_TipoDte", Datos[0]);
                    cmd.Parameters.AddWithValue("@dc_RutEmisor", Datos[1]);
                    cmd.Parameters.AddWithValue("@dc_RutReceptor", Datos[2]);
                    cmd.Parameters.AddWithValue("@dn_Folio", Datos[3]);
                    cmd.Parameters.AddWithValue("@df_Fech_Emision", Fech_Emision);
                    cmd.Parameters.AddWithValue("@dn_MontoTotal", Datos[5]);
                    cmd.Parameters.AddWithValue("@dc_RutCedente", Datos[6]);
                    cmd.Parameters.AddWithValue("@dg_RznSocial", Datos[7]);
                    cmd.Parameters.AddWithValue("@dg_DireccCedente", Datos[8]);
                    cmd.Parameters.AddWithValue("@dg_MailCedente", Datos[9]);
                    cmd.Parameters.AddWithValue("@dc_RutAutorizado", Datos[10]);
                    cmd.Parameters.AddWithValue("@dg_NomAutorizado", Datos[11]);
                    cmd.Parameters.AddWithValue("@dg_DeclJurada", Datos[12]);
                    cmd.Parameters.AddWithValue("@dc_RutCesionario", Datos[13]);
                    cmd.Parameters.AddWithValue("@dg_RznSocCesionario", Datos[14]);                   
                    cmd.Parameters.AddWithValue("@dg_DireccCesionario", Datos[15]);
                    cmd.Parameters.AddWithValue("@dg_MailCesionario", Datos[16]);
                    cmd.Parameters.AddWithValue("@dn_MontoCesion", Datos[17]);
                    cmd.Parameters.AddWithValue("@df_UltimoVenc", UltimoVenc);
                    cmd.Parameters.AddWithValue("@dg_MailDeudor", Datos[19]);
                    cmd.Parameters.AddWithValue("@df_TmsCesion", TmsCesion);
                    cmd.Parameters.AddWithValue("@df_FechaReg", "");

                    cmd.Connection = Connex;
                    cmd.ExecuteNonQuery();

                    Connex.Close();
                }
                else
                {
                    Logs.Log("FilterMail.SaveXmlInfo", "Error en la Conexion, o algo fallo en el StringConnection");
                }

            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.SaveXmlInfo", e.Message.ToString().ToString());
            }


        }

        private void SaveDicomInfo(String[]Datos)
        {
            try
            {
                SqlCommand cmd = new SqlCommand();
                string connString = ConfigurationManager.ConnectionStrings["myConnectionString"].ConnectionString;
                SqlConnection Connex = new SqlConnection(connString);

                String FechaPDF =Datos[1]+" " + Datos[2];

                Connex.Open();
                if (Connex != null && Connex.State == ConnectionState.Open)
                {
                    cmd = new SqlCommand("sp_i_cert_dicom");
                    cmd.CommandType = CommandType.StoredProcedure;

                    DateTime Fech_Pdf = DateTime.Parse(FechaPDF);                   

                    cmd.Parameters.AddWithValue("@dc_rut_empresa", Datos[0]);
                    cmd.Parameters.AddWithValue("@df_emision", Fech_Pdf);
                    cmd.Parameters.AddWithValue("@dg_nombre_archivo", Datos[4]);
                                      

                    cmd.Connection = Connex;
                    cmd.ExecuteNonQuery();

                    Connex.Close();
                }
                else
                {
                    Logs.Log("FilterMail.SaveXmlInfo", "Error en la Conexion, o algo fallo en el StringConnection");
                }

            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.SaveXmlInfo", e.Message.ToString().ToString());
            }


        }

        private string GetRutEnvio(String BodyAtt)
    {
            
            string strStart = "Rut Envio              :";
            string strEnd = "Identificador de Envio :";

            try
            {
                int Start, End;
                if (BodyAtt.Contains(strStart) && BodyAtt.Contains(strEnd))
                {
                    Start = BodyAtt.IndexOf(strStart, 0) + strStart.Length;
                    End = BodyAtt.IndexOf(strEnd, Start);                    
                    BodyAtt = BodyAtt.Substring(Start, End - Start);
                    BodyAtt = (BodyAtt).Trim();
                    return BodyAtt;
                }
                else
                {
                    return "";
                }
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetRutEnvio", e.Message.ToString().ToString());
                return "";
            }

            
        }
        private string GetIdentEnvio(String BodyAtt)
        {
            
            string strStart = "Identificador de Envio :";
            string strEnd = "Fecha de Recepcion     :";

            try
            {
                int Start, End;
                if (BodyAtt.Contains(strStart) && BodyAtt.Contains(strEnd))
                {
                    Start = BodyAtt.IndexOf(strStart, 0) + strStart.Length;
                    End = BodyAtt.IndexOf(strEnd, Start);
                    BodyAtt = BodyAtt.Substring(Start, End - Start);
                    BodyAtt = (BodyAtt).Trim();
                    return BodyAtt;
                }
                else
                {
                    return "";
                }
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetRutEnvio", e.Message.ToString().ToString());
                return "";
            }
        }
        private string GetFechaRecep( String BodyAtt)
        {
            {

                string strStart = "Fecha de Recepcion     :";
                string strEnd = "Estado del Envio       :";

                try
                {
                    int Start, End;
                    if (BodyAtt.Contains(strStart) && BodyAtt.Contains(strEnd))
                    {
                        Start = BodyAtt.IndexOf(strStart, 0) + strStart.Length;
                        End = BodyAtt.IndexOf(strEnd, Start);
                        BodyAtt = BodyAtt.Substring(Start, End - Start);
                        BodyAtt = (BodyAtt).Trim();
                        return BodyAtt;
                    }
                    else
                    {
                        return "";
                    }
                }
                catch (Exception e)
                {
                    Logs.Log("FilterMail.GetRutEnvio", e.Message.ToString().ToString());
                    return "";
                }               

            }
        }
        private string GetEstEnvio(String BodyAtt)
        {
            string strStart = "Estado del Envio       :";
            string strEnd = "Detalle de Cesion Electronica de Credito";

            try
            {
                int Start, End;
                if (BodyAtt.Contains(strStart) && BodyAtt.Contains(strEnd))
                {
                    Start = BodyAtt.IndexOf(strStart, 0) + strStart.Length;
                    End = BodyAtt.IndexOf(strEnd, Start);
                    BodyAtt = BodyAtt.Substring(Start, End - Start);
                    BodyAtt = (BodyAtt).Trim();
                    return BodyAtt;
                }
                else
                {
                    return "";
                }
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetEstEnvio", e.Message.ToString().ToString());
                return "";
            }
        }
        private string GetTipDoc(String BodyAtt)
        {
            string strAfecta = "FACTURA ELECTRONICA";
            string strExenta = "EXENTA";

           try
            {
                if (BodyAtt.Contains(strAfecta))
                {
                    return "Afecta";
                }
                if (BodyAtt.Contains(strExenta))
                {
                    return "Exenta";
                }
                else
                {
                    return "";
                }
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetTipDoc", e.Message.ToString().ToString());
                return "";
            }
        }
        private string GetNroDoc(String BodyAtt)
        {
            string strStart = "ELECTRONICA N";
            string strEnd = "- Monto Total:";

            try
            {
                int Start, End;
                if (BodyAtt.Contains(strStart) && BodyAtt.Contains(strEnd))
                {
                    Start = BodyAtt.IndexOf(strStart, 0) + strStart.Length + 1;
                    End = BodyAtt.IndexOf(strEnd, Start);
                    BodyAtt = BodyAtt.Substring(Start, End - Start);
                    BodyAtt = (BodyAtt).Trim();
                    return BodyAtt;
                }
                else
                {
                    return "";
                }
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetNroDoc", e.Message.ToString().ToString());
                return "";
            }

        }
        //private string GetNroFact(String BodyAtt)
        //{
                        
        //    string strStart = "-Monto Total: $";
        //    string strEnd = "Emisor  :";

        //    int Start, End;
        //    if (BodyAtt.Contains(strStart) && BodyAtt.Contains(strEnd))
        //    {
        //        Start = BodyAtt.IndexOf(strStart, 0) + strStart.Length;
        //        End = BodyAtt.IndexOf(strEnd, Start);
        //        BodyAtt = BodyAtt.Substring(Start, End - Start);
        //        BodyAtt = (BodyAtt).Trim();
        //        return BodyAtt;
        //    }
        //    else
        //    {
        //        return "";
        //    }

        //}
        private string GetMonTotal(String BodyAtt)
        {
            string strStart = "- Monto Total: $";
            string strEnd = "Emisor  :";
            try
            {
                int Start, End;
                if (BodyAtt.Contains(strStart) && BodyAtt.Contains(strEnd))
                {
                    Start = BodyAtt.IndexOf(strStart, 0) + strStart.Length;
                    End = BodyAtt.IndexOf(strEnd, Start);
                    BodyAtt = BodyAtt.Substring(Start, End - Start);
                    BodyAtt = (BodyAtt).Trim();
                    return BodyAtt;
                }
                else
                {
                    return "";
                }

            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetMonTotal", e.Message.ToString().ToString());
                return "";
            }   
        }
        private string GetRutEmisor(String BodyAtt)
        {
            string strStart = "Emisor  :";
            string strEnd = strStart+13;
            try
            {
                int Start, End;
                if (BodyAtt.Contains(strStart))
                {
                    Start = BodyAtt.IndexOf(strStart, 0) + strStart.Length;
                    End = Start + 13;
                    BodyAtt = BodyAtt.Substring(Start, End - Start);
                    BodyAtt = (BodyAtt).Trim();
                    return BodyAtt;
                }
                else
                {
                    return "";
                }
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetRutEmisor", e.Message.ToString());
                return "";
            }        
        }
        private string GetNomEmisor(String BodyAtt)
        {
            string strStart = "Emisor  :";            
            string strEnd = "Receptor:";

            try
            {
                int Start, End;
                if (BodyAtt.Contains(strStart))
                {

                    Start = BodyAtt.IndexOf(strStart, 0) + strStart.Length + 12;
                    End = BodyAtt.IndexOf(strEnd, Start);
                    BodyAtt = BodyAtt.Substring(Start, End - Start);
                    BodyAtt = (BodyAtt).Trim();
                    return BodyAtt;
                }
                else
                {
                    return "";
                }
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetNomEmisor", e.Message.ToString());
                return "";
            }            
        }        
        private string GetRutReceptor(String BodyAtt)
        {
            string strStart = "Receptor:";
            string strEnd = strStart + 13;
            try
            {
                int Start, End;
                if (BodyAtt.Contains(strStart))
                {
                    Start = BodyAtt.IndexOf(strStart, 0) + strStart.Length;
                    End = Start + 13;
                    BodyAtt = BodyAtt.Substring(Start, End - Start);
                    BodyAtt = (BodyAtt).Trim();
                    return BodyAtt;
                }
                else
                {
                    return "";
                }
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetRutReceptor", e.Message.ToString());
                return "";
            }
        }
        private string GetNomReceptor(String BodyAtt)
        {
            string strStart = "Receptor:";
            string strEnd = "Fecha Emision:";
            try
            {
                int Start, End;
                if (BodyAtt.Contains(strStart))
                {
                    Start = BodyAtt.IndexOf(strStart, 0) + strStart.Length + 12;
                    End = BodyAtt.IndexOf(strEnd, Start);
                    BodyAtt = BodyAtt.Substring(Start, End - Start);
                    BodyAtt = (BodyAtt).Trim();
                    return BodyAtt;
                }
                else
                {
                    return "";
                }
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetNomReceptor", e.Message.ToString());
                return "";
            }
            
        }
        private string GetDocCedFechaEmision(String BodyAtt)
        {
            
            string strStart = "Fecha Emision:";
            string strEnd = "Cedente";
            try
            {
                int Start, End;
                if (BodyAtt.Contains(strStart) && BodyAtt.Contains(strEnd))
                {
                    Start = BodyAtt.IndexOf(strStart, 0) + strStart.Length;
                    End = BodyAtt.IndexOf(strEnd, Start);
                    BodyAtt = BodyAtt.Substring(Start, End - Start);
                    BodyAtt = (BodyAtt).Trim();
                    return BodyAtt;
                }
                else
                {
                    return "";
                }
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetDocCedFechaEmision", e.Message.ToString());
                return "";
            }
        }
        private string GetRutCedente(String BodyAtt)
        {
            string strStart = "Cedido por:";           
            string strEnd = strStart + 13;
            try
            {
                int Start, End;
                if (BodyAtt.Contains(strStart))
                {
                    Start = BodyAtt.IndexOf(strStart, 0) + strStart.Length;
                    End = Start + 13;
                    BodyAtt = BodyAtt.Substring(Start, End - Start);
                    BodyAtt = (BodyAtt).Trim();
                    return BodyAtt;
                }
                else
                {
                    return "";
                }
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetRutCedente", e.Message.ToString());
                return "";
            }
        }
        private string GetNomCedente(String BodyAtt)
        {
            string strStart = "Cedido por: ";
            string strEnd = "Direccion:";

            try
            {
                int Start, End;
                if (BodyAtt.Contains(strStart))
                {

                    Start = BodyAtt.IndexOf(strStart, 0) + strStart.Length + 13;
                    End = BodyAtt.IndexOf(strEnd, Start);
                    BodyAtt = BodyAtt.Substring(Start, End - Start);
                    BodyAtt = (BodyAtt).Trim();
                    return BodyAtt;
                }
                else
                {
                    return "";
                }
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetNomCedente", e.Message.ToString());
                return "";
            }           

        }
        private string GetDirecCedente(String BodyAtt)
        {
            string strStart = "Direccion:";
            string strEnd = "eMail:";

            try
            {
                int Start, End;
                if (BodyAtt.Contains(strStart) && BodyAtt.Contains(strEnd))
                {
                    Start = BodyAtt.IndexOf(strStart, 0) + strStart.Length;
                    End = BodyAtt.IndexOf(strEnd, Start);
                    BodyAtt = BodyAtt.Substring(Start, End - Start);
                    BodyAtt = (BodyAtt).Trim();
                    return BodyAtt;
                }
                else
                {
                    return "";
                }
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetDirecCedente", e.Message.ToString());
                return "";
            }           

        }
        private string GetEmailCedente(String BodyAtt)
        {
        
            string strStart = "eMail:";
            string strEnd = "Cesionario";


            try
            {
                int Start, End;
                if (BodyAtt.Contains(strStart) && BodyAtt.Contains(strEnd))
                {
                    Start = BodyAtt.IndexOf(strStart, 0) + strStart.Length;
                    End = BodyAtt.IndexOf(strEnd, Start);
                    BodyAtt = BodyAtt.Substring(Start, End - Start);
                    BodyAtt = (BodyAtt).Trim();
                    return BodyAtt;
                }
                else
                {
                    return "";
                }
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetEmailCedente", e.Message.ToString());
                return "";
            }            
        }
        private string GetRutCesionario(String BodyAtt)
        {
            
            string strStart = "Cedido a:";            
            string strEnd = strStart + 13;
            try {
                int Start, End;
                if (BodyAtt.Contains(strStart))
                {
                    Start = BodyAtt.IndexOf(strStart, 0) + strStart.Length;
                    End = Start + 13;
                    BodyAtt = BodyAtt.Substring(Start, End - Start);
                    BodyAtt = (BodyAtt).Trim();
                    return BodyAtt;
                }
                else
                {
                    return "";
                }
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetRutCesionario", e.Message.ToString());
                return "";
            }
        }  //
        
        private string GetDeclJuradaXML(string BodyAtt)
        {
            string strStart = "DeclaracionJurada:";
            string strEnd = "Cesionario:";
            try
            {
                int Start, End;
                if (BodyAtt.Contains(strStart))
                {
                    Start = BodyAtt.IndexOf(strStart, 0) + strStart.Length;
                    End = BodyAtt.IndexOf(strEnd, Start);
                    BodyAtt = BodyAtt.Substring(Start, End - Start);
                    BodyAtt = (BodyAtt).Trim();
                    return BodyAtt;
                }
                else
                {
                    return "";
                }
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetNombreCesionario", e.Message.ToString());
                return "";
            }

        }
        private string GetNombreCesionario(String BodyAtt)
        {
            string strStart = "Cedido a:";
            string strEnd = "Direccion:";
            try
            {
                int Start, End;
                if (BodyAtt.Contains(strStart))
                {
                    Start = BodyAtt.IndexOf(strStart, 0) + strStart.Length + 13;
                    End = BodyAtt.IndexOf(strEnd, Start);
                    BodyAtt = BodyAtt.Substring(Start, End - Start);
                    BodyAtt = (BodyAtt).Trim();
                    return BodyAtt;
                }
                else
                {
                    return "";
                }
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetNombreCesionario", e.Message.ToString());
                return "";
            }
            
        }
        private string GetFechaCesion(String BodyAtt)
        {
            string strStart = "Fecha de la Cesion:";
            string strEnd = "Declaracion Jurada:";

            try
            {
                int Start, End;
                if (BodyAtt.Contains(strStart))
                {
                    Start = BodyAtt.IndexOf(strStart, 0) + strStart.Length;
                    End = BodyAtt.IndexOf(strEnd, Start);
                    BodyAtt = BodyAtt.Substring(Start, End - Start);
                    BodyAtt = (BodyAtt).Trim();
                    return BodyAtt;
                }
                else
                {
                    return "";
                }

            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetFechaCesion", e.Message.ToString());
                return "";
            }
           
        }
        private string GetDeclJurada(String BodyAtt)
        {
            string strStart = "Declaracion Jurada:";
            string strEnd = "Monto Cedido:";
            try
            {
                int Start, End;
                if (BodyAtt.Contains(strStart))
                {
                    Start = BodyAtt.IndexOf(strStart, 0) + strStart.Length;
                    End = BodyAtt.IndexOf(strEnd, Start);
                    BodyAtt = BodyAtt.Substring(Start, End - Start);
                    BodyAtt = (BodyAtt).Trim();
                    return BodyAtt;
                }
                else
                {
                    return "";
                }

            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetDeclJurada", e.Message.ToString());
                return "";
            }
                        
        }
        private string GetMontoCedido(String BodyAtt)
        {
            string strStart = "Monto Cedido: $";
            string strEnd = "Ultimo Vencimiento:";
            try
            {
                int Start, End;
                if (BodyAtt.Contains(strStart))
                {
                    Start = BodyAtt.IndexOf(strStart, 0) + strStart.Length;
                    End = BodyAtt.IndexOf(strEnd, Start);
                    BodyAtt = BodyAtt.Substring(Start, End - Start);
                    BodyAtt = (BodyAtt).Trim();
                    return BodyAtt;
                }
                else
                {
                    return "";
                }

            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetMontoCedido", e.Message.ToString());
                return "";
            }            
        }
        private string GetUltimoVenc(String BodyAtt)
        {
            string strStart = "Ultimo Vencimiento:";
            string strEnd = "Otras Condiciones:";
            try
            {
                int Start, End;
                if (BodyAtt.Contains(strStart))
                {
                    Start = BodyAtt.IndexOf(strStart, 0) + strStart.Length;
                    End = BodyAtt.IndexOf(strEnd, Start);
                    BodyAtt = BodyAtt.Substring(Start, End - Start);
                    BodyAtt = (BodyAtt).Trim();
                    return BodyAtt;
                }
                else
                {
                    return "";
                }
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetUltimoVenc", e.Message.ToString());
                return "";
            }
        }
        private string GetIdentCesion(String BodyAtt)
        {
                        
            string strStart = "Identificador de Cesion:";
            string strEnd = strStart + 25;
            try
            {
                int Start, End;
                if (BodyAtt.Contains(strStart))
                {
                    Start = BodyAtt.IndexOf(strStart, 0) + strStart.Length;
                    End = Start + 25;
                    BodyAtt = BodyAtt.Substring(Start, End - Start);
                    BodyAtt = (BodyAtt).Trim();
                    return BodyAtt;
                }
                else
                {
                    return "";
                }
            }
            catch (Exception e)
            {
                Logs.Log("FilterMail.GetIdentCesion", e.Message.ToString());
                return "";
            }          

        }


        private Boolean AskFileExistOnTableRptc(String[]DatRptc)
        {            
            string NroDoc = string.Empty;
            string MontTotal = string.Empty;
            string TipoDoc = string.Empty;
            string IdCesion = DatRptc[22];
            StringBuilder sb = new StringBuilder();
            SqlCommand cmd = new SqlCommand();
            SqlDataReader reader;

            string connString = ConfigurationManager.ConnectionStrings["myConnectionString"].ConnectionString;
            SqlConnection Connex = new SqlConnection(connString);

            //ShopEmail(ref xMail);


            sb.Append("if exists(select dc_IdentCesion from tb_rptc_info where dc_IdentCesion = ");
            sb.Append("'" + IdCesion + "')");
            sb.Append("begin select * from tb_rptc_info where dc_IdentCesion =");
            sb.Append("'" + IdCesion + "'end");

            Connex.Open();
            if (Connex != null && Connex.State == ConnectionState.Open)
            {
                string query = sb.ToString();
                cmd.CommandText = query;
                cmd.Connection = Connex;
                reader = cmd.ExecuteReader();
                reader.Read();
                
                if (reader.HasRows)
                {
                    TipoDoc = reader["dn_TipDoc"].ToString();
                    NroDoc = reader["dn_NroDoc"].ToString();
                    MontTotal = reader["dn_MonTotal"].ToString();
                    IdCesion = reader["dc_IdentCesion"].ToString();

                    Connex.Close();
                    return true;
                }
                else
                {
                    Connex.Close();
                    return false;
                }
                //Connex.Close();

            }
            else
            {
                //error al log
                Connex.Close();
                return false;
            }
            //Connex.Close();
            //return false;
        }

        private Boolean AskFileExistOnTableXML(String[] DatRptc)
        {
            string dn_TipoDte = string.Empty;
            string dc_RutEmisor = string.Empty;
            string dn_Folio = string.Empty;            
            StringBuilder sb = new StringBuilder();
            SqlCommand cmd = new SqlCommand();
            SqlDataReader reader;

            string connString = ConfigurationManager.ConnectionStrings["myConnectionString"].ConnectionString;
            SqlConnection Connex = new SqlConnection(connString);

            //ShopEmail(ref xMail);

           //// dn_TipoDte, dc_RutEmisor, dn_Folio


            sb.Append("if exists(select * from tb_xml_info where dn_Folio = ");
            sb.Append("'" + DatRptc[3] + "')");
            sb.Append("'and dc_RutEmisor=" + dc_RutEmisor + "')");
            sb.Append("' and dn_TipoDte=" + dn_TipoDte + "')");
            //sb.Append("begin select * from tb_xml_info where dn_Folio =");
            //sb.Append("'" + dn_Folio + "'end");

            Connex.Open();
            if (Connex != null && Connex.State == ConnectionState.Open)
            {
                string query = sb.ToString();
                cmd.CommandText = query;
                cmd.Connection = Connex;
                reader = cmd.ExecuteReader();
                reader.Read();

                if (reader.HasRows)
                {
                    dn_TipoDte = reader["dn_TipoDte"].ToString();
                    dc_RutEmisor = reader["dn_NroDoc"].ToString();
                    dn_Folio = reader["dn_MonTotal"].ToString();
                                                             

                    Connex.Close();
                    return true;
                }
                else
                {
                    Connex.Close();
                    return false;
                }
                //Connex.Close();

            }
            else
            {
                //error al log
                Connex.Close();
                return false;
            }
            //Connex.Close();
            //return false;
        }
    }
}






//  String oFrom = (oMail.From).ToString();
//  ShopODestino(ref oFrom);
//  bool Saved = false;

//  if (oFrom == "crosas")  /////  || oFrom== "crosas"
//  {
//      XmlFilter(oMail,i);
//      Saved = true;
//  }
//  if (oFrom == "tcuello")  /////  || oFrom== "crosas"//tcuello@arayahermanos.cl
//  {
//      XmlFilter(oMail, i);
//      Saved = true;
//  }


//  if (oFrom == "administracion")  /////  || oFrom== "crosas"//administracion@labotec.cl
//  {
//      XmlFilter(oMail, i);
//      Saved = true;
//  }

//  if (oFrom == "zuniga")  /////  || oFrom== "crosas"  //zuniga@decomuebles.cl
//  {
//      XmlFilter(oMail, i);
//      Saved = true;
//  }

//  if (oFrom == "zuniga")  /////  || oFrom== "crosas" //dte@bsale.cl
//  {
//      XmlFilter(oMail, i);
//      Saved = true;
//  }              

//  //dte_dilorsa@empresasorion.cl        
//  if (oFrom == "dte_dilorsa")  /////  || oFrom== "crosas"//administracion@labotec.cl
//  {
//      XmlFilter(oMail, i);
//      Saved = true;
//  }

//  ////sii@crosan.cl
//  if (oFrom == "sii")  /////  || oFrom== "crosas"   ////sii@crosan.cl
//  {
//      XmlFilter(oMail, i);
//      Saved = true;
//  }

//  //simtexx@simtexx.cl
//  if (oFrom == "simtexx")  /////  || oFrom== "crosas"//simtexx@simtexx.cl
//  {
//      XmlFilter(oMail, i);
//      Saved = true;
//  }


//  ////si es de servifactura  --- enviodte  -- enviodte@servifactura.cl
//  if (oFrom == "enviodte")  /////  || oFrom== "crosas"
//  {                    
//    ServiFacturaFilter(oMail, i);
//      Saved = true;
//  }

//  // Si es RPTC
//  if (oFrom == "rpetc" )  /////  || oFrom== "crosas"
//  {
//  RptcFilter(oMail, i);
//      Saved = true;            
//  }
//  // SI ES DE dte@move-up.cl
//  if (oFrom == "dte")  /////  || oFrom== "crosas"  //dte@move-up.cl
//  {
//      RptcFilter(oMail, i);
//      Saved = true;
//  }
//  //si es  DeptoControlGestion
//  if (oFrom == "DeptoControlGestion") /////|| oFrom == "crosas"
//  {
//    DeptoControlGestionFilter(oMail, i);
//      Saved = true;
//  }
//  //si es dicom
//  if (oFrom == "informes.equifax")  //////informes.equifax@equifax.cl
//  {
// DicomFilter(oMail, i);
//      Saved = true;
//  }

//  //// si no entra a ninguno de los if lo graba igual en los descartados
//  if (Saved == false) 
//  { 
// DiscardFilter(oMail, i);
//  }

//  /////-------- delarar salida de servidor: Marcado de correos descargados y cerraar conexion
////  PoClient

