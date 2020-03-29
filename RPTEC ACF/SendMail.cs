using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using System.Net.NetworkInformation;
using System.Management;



namespace RPTEC_ACF
{
    class SendMail
    {
        public Boolean SenderMail(String WhoReceive, String Subject, String Body)
        {
            try
            {
                String WhoSendFalse = "Repositorio@acfcapital.cl";
                /// Command line argument must the the SMTP host.
                SmtpClient client = new SmtpClient();
                client.Port = 25;
                client.Host = "smtp1.acf";
                client.EnableSsl = false;
                client.Timeout = 10000;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.UseDefaultCredentials = false;
                client.Credentials = new System.Net.NetworkCredential("depuradormail@acfcapital.cl", "depu1606");
                //WhoReceive = "crosas@acfcapital.cl";

                MailMessage mm = new MailMessage(WhoSendFalse, WhoReceive, Subject, Body);
                mm.Bcc.Add("ogarrido@acfcapital.cl");
                mm.Bcc.Add("crosas@acfcapital.cl");
                mm.BodyEncoding = UTF8Encoding.UTF8;
                mm.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure;
                System.DateTime Nowww = DateTime.Now;

                //client.Send(mm);
                Logs.Log("Se Envio Correo a: " + WhoReceive, " Con el Asunto : " + Subject + " a las : " + Nowww.ToString());
                return true;
            }
            catch (Exception ep)
            {
                Logs.Log("SendMail.SenderMail", ep.Message.ToString());
                return false;
            }
        }



        public Boolean SenderMailAlert(String WhoReceive, string Subject,string Body)
        {
            try
            {
                
                String WhoSendFalse = "Repositorio@acfcapital.cl";
                //String WhoReceive = "Crosas@acfcapital.cl";
                //String Subject = "ALERTA DEPURADOR MAIL ACF- Error De Conexion Con NAS 10.177.1.220";
                /// Command line argument must the the SMTP host.
                SmtpClient client = new SmtpClient();
                client.Port = 25;
                client.Host = "smtp1.acf";
                client.EnableSsl = false;
                client.Timeout = 10000;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.UseDefaultCredentials = false;
                client.Credentials = new System.Net.NetworkCredential("depuradormail@acfcapital.cl", "depu1606");
                //WhoReceive = "crosas@acfcapital.cl";

                MailMessage mm = new MailMessage(WhoSendFalse, WhoReceive, Subject, Body);
                //mm.Bcc.Add("ogarrido@acfcapital.cl");
                mm.Bcc.Add("jbriceno@acfcapital.cl");
                mm.BodyEncoding = UTF8Encoding.UTF8;
                mm.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure;
                System.DateTime Nowww = DateTime.Now;

                client.Send(mm);
                //Logs.Log("Se Envio Correo a: " + WhoReceive, " Con el Asunto : " + Subject + " a las : " + Nowww.ToString());
                return true;
            }
            catch (Exception ep)
            {
                Logs.Log("SendMail.SenderMail", ep.Message.ToString());
                return false;
            }
        }


        //SEND MULTIPLES MAILS
        //SmtpClient client = new SmtpClient("smtphost", 25);
        //MailMessage msg = new MailMessage("x@y.com", "a@b.com,c@d.com");
        //msg.Subject = "sdfdsf";
        //msg.Body = "sdfsdfdsfd";
        //client.UseDefaultCredentials = true;
        //client.Send(msg);




        //OTHER CODE
        //1.The ACCOUNT
        //        MailAddress fromAddress = new MailAddress("myaccount@myaccount.com", "my display name");
        //        String fromPassword = "password";

        //        //2.The Destination email Addresses
        //        MailAddressCollection TO_addressList = new MailAddressCollection();

        ////        3.Prepare the Destination email Addresses list
        //foreach (var curr_address in mailto.Split(new [] {";"}, StringSplitOptions.RemoveEmptyEntries))
        //{
        //    MailAddress mytoAddress = new MailAddress(curr_address, "Custom display name");
        //        TO_addressList.Add(mytoAddress);
        //}

        //    //4.The Email Body Message
        //    String body = bodymsg;

        //    //5.Prepare GMAIL SMTP: with SSL on port 587
        //    var smtp = new SmtpClient
        //    {
        //        Host = "smtp.gmail.com",
        //        Port = 587,
        //        EnableSsl = true,
        //        DeliveryMethod = SmtpDeliveryMethod.Network,
        //        Credentials = new NetworkCredential(fromAddress.Address, fromPassword),
        //        Timeout = 30000
        //    };


        ////6.Complete the message and SEND the email:
        //using (var message = new MailMessage()
        //       {
        //           From = fromAddress,
        //           Subject = subject,
        //           Body = body,
        //       })
        //{
        //    message.To.Add(TO_addressList.ToString());
        //    smtp.Send(message);
        //}

        /////


        ////ANOTHER CODE
        //This list can be a parameter of metothd
        //    List<MailAddress> lst = new List<MailAddress>();

        //    lst.Add(new MailAddress("mouse@xxxx.com"));
        //lst.Add(new MailAddress("duck@xxxx.com"));
        //lst.Add(new MailAddress("goose@xxxx.com"));
        //lst.Add(new MailAddress("wolf@xxxx.com"));


        //try
        //{


        //    MailMessage objeto_mail = new MailMessage();
        //    SmtpClient client = new SmtpClient();
        //    client.Port = 25;
        //    client.Host = "10.15.130.28"; //or SMTP name
        //    client.Timeout = 10000;
        //    client.DeliveryMethod = SmtpDeliveryMethod.Network;
        //    client.UseDefaultCredentials = false;
        //    client.Credentials = new System.Net.NetworkCredential("from@email.com", "password");
        //    objeto_mail.From = new MailAddress("from@email.com");

        //    //add each email adress
        //    foreach (MailAddress m in lst)
        //    {
        //        objeto_mail.To.Add(m);
        //    }


        //objeto_mail.Subject = "Sending mail test";
        //    objeto_mail.Body = "Functional test for automatic mail :-)";
        //    client.Send(objeto_mail);


        //}
        //catch (Exception ex)
        //{
        //    MessageBox.Show(ex.Message);
        //}



    }
}
