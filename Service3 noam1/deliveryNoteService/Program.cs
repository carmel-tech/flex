using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Xml;
using System.Net;
using System.IO;

namespace deliveryNoteService
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
//        const string Prefix = "http://+:80/dashboard/";
        static string responseString = "";
        static HttpListener Listener = null;
        static deliveryNoteService svc = null;
        /// </summary>
        static void Main(string[] args)
        {
            ServiceBase[] ServicesToRun;
            svc = new deliveryNoteService();
            ServicesToRun = new ServiceBase[] 
			{ 
				 svc
			};
            ReadHtmlTemplate();
            StartHtpListener();
            ServiceBase.Run(ServicesToRun);
            StopHtpListener();
        }
        // This example requires the System and System.Net namespaces. 
        // netsh http add urlacl http://+:8008/ user=Everyone listen=yes
        public static void ReadHtmlTemplate()
        {
            string folder = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string path = Path.Combine(folder, "Index.html");
            // Open the stream and read it back. 
            try
            {
                responseString = System.IO.File.ReadAllText(path);
                //              Console.WriteLine(responseString);
            }
            catch (System.IO.IOException)
            {
            }
        }
        public static void StartHtpListener()
        {
            if (!HttpListener.IsSupported)
            {
                //Console.WriteLine("Windows XP SP2 or Server 2003 is required to use the HttpListener class.");
                return;
            }
            // URI prefixes are required, 
            // for example "http://contoso.com:8080/index/".

            Listener = new HttpListener();
            Listener.Prefixes.Add(Properties.Settings.Default.HTTPPrefix);
            try
            {
                Listener.Start();
                // Begin waiting for requests.
                IAsyncResult ar = Listener.BeginGetContext(GetContextCallback, null);
            }
            catch (System.Net.HttpListenerException)
            {   
            }
        }
        public static void StopHtpListener()
        {
            if (Listener != null)
            {
                Listener.Close();
                Listener = null;
            }
        }
        static void GetContextCallback(IAsyncResult ar)
        {
            if (Listener == null)
            {
                return;
            }
            // Get the context
            var context = Listener.EndGetContext(ar);
            // listen for the next request
            Listener.BeginGetContext(GetContextCallback, null);
            // get the request
            //            var NowTime = DateTime.UtcNow;
            //            Console.WriteLine("{0}: {1}", NowTime.ToString("R"), context.Request.RawUrl);
            // format response
            // string.Format("<html><body>Your request, \"{0}\", was received at {1}.<br/>It is request #{2:N0} since {3}.", context.Request.RawUrl, NowTime.ToString("R"), req, StartupDate.ToString("R"));
            string s = responseString.Replace("%date%", DateTime.Today.ToString("dd/MM/yyyy"));
            s = s.Replace("%inv_bad%", svc.InvBad.ToString());
            s = s.Replace("%inv_bad_pcnt%", svc.InvBadPercantage.ToString("F2", System.Globalization.CultureInfo.InvariantCulture));
            s = s.Replace("%inv_good%", svc.InvGood.ToString());
            s = s.Replace("%inv_total%", svc.InvTotal.ToString());
            s = s.Replace("%ps_bad%", svc.PsBad.ToString());
            s = s.Replace("%ps_bad_pcnt%", svc.PsBadPercantage.ToString("F2", System.Globalization.CultureInfo.InvariantCulture));
            s = s.Replace("%ps_good_signed%", svc.PsGood_signed.ToString());
            s = s.Replace("%ps_good_archive%", svc.PsGood_archive.ToString());
            s = s.Replace("%ps_total%", svc.PsTotal.ToString());
            s = s.Replace("%ps_bad_1%", svc.PsBadOfaqim.ToString());
            s = s.Replace("%ps_bad_pcnt_1%", svc.PsBadPercantageOfaqim.ToString("F2", System.Globalization.CultureInfo.InvariantCulture));
            s = s.Replace("%ps_good_1_signed%", svc.PsGoodOfaqim_signed.ToString());
            s = s.Replace("%ps_good_1_archive%", svc.PsGoodOfaqim_archive.ToString());
            s = s.Replace("%ps_total_1%", svc.PsTotalOfaqim.ToString());
            //Melanox
            s = s.Replace("%ps_bad_1_mel%", svc.PsBadMellanox.ToString());
            s = s.Replace("%ps_bad_pcnt_1_mel%", svc.PsBadPercantageMellanox.ToString("F2", System.Globalization.CultureInfo.InvariantCulture));
            s = s.Replace("%ps_good_1_signed_mel%", svc.PsGoodMellanox_signed.ToString());
            s = s.Replace("%ps_good_1_archive_mel%", svc.PsGoodMellanox_archive.ToString());
            s = s.Replace("%ps_total_1_mel%", svc.PsTotalMellanox.ToString());
            //Shipments
            //COCDeliveryNote_getbarcode
            s = s.Replace("%COC_getbarcode%", svc.COCGetBarcode.ToString());
            s = s.Replace("%COC_getdata%", svc.COCGetData.ToString());
            s = s.Replace("%COC_good%", svc.COCGood.ToString());
            s = s.Replace("%COC_bad%", svc.COCBad.ToString());
            s = s.Replace("%COC_total%", svc.COCTotal.ToString());
            s = s.Replace("%COC_bad_pcnt%", svc.COCBadPcnt.ToString("F2", System.Globalization.CultureInfo.InvariantCulture));
            //Invoice
            s = s.Replace("%Invoice_getbarcode%", svc.InvoiceGetBarcode.ToString());
            s = s.Replace("%Invoice_getdata%", svc.InvoiceGetData.ToString());
            s = s.Replace("%Invoice_good%", svc.InvoiceGood.ToString());
            s = s.Replace("%Invoice_bad%", svc.InvoiceBad.ToString());
            s = s.Replace("%Invoice_total%", svc.InvoiceTotal.ToString());
            s = s.Replace("%Invoice_bad_pcnt%", svc.InvoiceBadPcnt.ToString("F2", System.Globalization.CultureInfo.InvariantCulture));
            //Delivery Note
            s = s.Replace("%DeliveryNote_getbarcode%", svc.DeliveryNoteGetBarcode.ToString());
            s = s.Replace("%DeliveryNote_getdata%", svc.DeliveryNoteGetData.ToString());
            s = s.Replace("%DeliveryNote_good%", svc.DeliveryNoteGood.ToString());
            s = s.Replace("%DeliveryNote_bad%", svc.DeliveryNoteBad.ToString());
            s = s.Replace("%DeliveryNote_total%", svc.DeliveryNoteTotal.ToString());
            s = s.Replace("%DeliveryNote_bad_pcnt%", svc.DeliveryNoteBadPcnt.ToString("F2", System.Globalization.CultureInfo.InvariantCulture));
         
            if (svc.InvBadPercantage > 5.0m)
            {
                s = s.Replace("%inv_fnt_color%", "#FF0000;");
            }
            else
            {
                s = s.Replace("%inv_fnt_color%", "#000000;");
            }
            if (svc.PsBadPercantage > 5.0m)
            {
                s = s.Replace("%ps_fnt_color%", "#FF0000;");
            }
            else
            {
                s = s.Replace("%ps_fnt_color%", "#000000;");
            }

            if (svc.PsBadPercantageOfaqim > 5.0m)
            {
                s = s.Replace("%ps_fnt_color_1%", "#FF0000;");
            }
            else
            {
                s = s.Replace("%ps_fnt_color_1%", "#000000;");
            }
            if (svc.PsBadPercantageMellanox > 5.0m)
            {
                s = s.Replace("%ps_fnt_color_1_mel%", "#FF0000;");
            }
            else
            {
                s = s.Replace("%ps_fnt_color_1_mel%", "#000000;");
            }
            //COC
            if (svc.COCBadPcnt > 5.0m)
            {
                s = s.Replace("%COC_fnt_color%", "#FF0000;");
            }
            else
            {
                s = s.Replace("%COC_fnt_color%", "#000000;");
            }
            //Invoice
            if (svc.InvoiceBadPcnt > 5.0m)
            {
                s = s.Replace("%Invoice_fnt_color%", "#FF0000;");
            }
            else
            {
                s = s.Replace("%Invoice_fnt_color%", "#000000;");
            }
            //DeliveryNote
            if (svc.DeliveryNoteBadPcnt > 5.0m)
            {
                s = s.Replace("%DeliveryNote_fnt_color%", "#FF0000;");
            }
            else
            {
                s = s.Replace("%DeliveryNote_fnt_color%", "#000000;");
            }

            byte[] buffer = Encoding.UTF8.GetBytes(s);
            // and send it
            var response = context.Response;
            response.ContentType = "text/html";
            response.ContentLength64 = buffer.Length;
            response.StatusCode = 200;
            response.OutputStream.Write(buffer, 0, buffer.Length);
            response.OutputStream.Close();
        }
    }
}
