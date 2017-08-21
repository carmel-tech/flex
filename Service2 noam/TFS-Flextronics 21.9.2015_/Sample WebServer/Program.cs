using System;
using System.Net;
using System.Text;
using System.IO;
using System.Xml;
namespace TestServer
{
    class ServerMain
    {
        // To enable this so that it can be run in a non-administrator account:
        // Open an Administrator command prompt.
        //  netsh http add urlacl http://+:8008/ user=Everyone listen=yes
        const string Prefix = "http://+:80/dashboard/";
        static string text = "";
        static HttpListener Listener = null;
        static void Main(string[] args)
        {
            string test = "Num {0} = {0}";
            test = String.Format(test, 100);
            //---------------------------------------------------------------------------------------            
            SaveStatistics();
            //---------------------------------------------------------------------------------------            
            string folder = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string path = Path.Combine(folder, "statistics.xml");
            using (XmlReader reader = XmlReader.Create(path))
            {

                // Parse the XML document.  ReadString is used to 
                // read the text content of the elements.
                reader.Read();
                reader.ReadStartElement("Statitics");
                reader.ReadStartElement("Date");
                reader.ReadString();
                reader.ReadEndElement();
                
                reader.ReadStartElement("PsGood");
                int PsGood=reader.ReadContentAsInt();
                reader.ReadEndElement();

                reader.ReadStartElement("PsBad");
                int PsBad = reader.ReadContentAsInt();
                reader.ReadEndElement();

                reader.ReadStartElement("InvGood");
                int InvGood = reader.ReadContentAsInt();
                reader.ReadEndElement();

                reader.ReadStartElement("InvBad");
                int InvBad = reader.ReadContentAsInt();
                reader.ReadEndElement();
                
                reader.ReadEndElement();

            }

            //---------------------------------------------------------------------------------------            
            folder = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            path = Path.Combine(folder, "Index.htm"); 
            // Open the stream and read it back. 
            try
            {
                string s = System.IO.File.ReadAllText(path);
                s = s.Replace("%date%", DateTime.Today.ToShortDateString());
                text += s;
                Console.WriteLine(s);
            }
            catch (System.IO.IOException)
//            catch (System.IO.FileNotFoundException)
            {
            }
            if (!HttpListener.IsSupported)
            {
                Console.WriteLine("HttpListener is not supported on this platform.");
                return;
            }
            using (Listener = new HttpListener())
            {
                Listener.Prefixes.Add(Prefix);
                try
                {
                    Listener.Start();
                    // Begin waiting for requests.
                    IAsyncResult ar = Listener.BeginGetContext(GetContextCallback, null);
                    Console.WriteLine("Listening.  Press Enter to stop.");
                    Console.ReadLine();
                    Listener.Close();
                    Listener = null;
                }
                catch(System.Net.HttpListenerException)
                {
                }
            }
        }
        static int RequestNumber = 0;
        static readonly DateTime StartupDate = DateTime.UtcNow;
        static void GetContextCallback(IAsyncResult ar)
        {
            if (Listener == null)
            {
                return;
            }
            int req = ++RequestNumber;
            // Get the context
            var context = Listener.EndGetContext(ar);
            // listen for the next request
            Listener.BeginGetContext(GetContextCallback, null);
            // get the request
            var NowTime = DateTime.UtcNow;
            Console.WriteLine("{0}: {1}", NowTime.ToString("R"), context.Request.RawUrl);
            // format response
            string responseString = text;
            // string.Format("<html><body>Your request, \"{0}\", was received at {1}.<br/>It is request #{2:N0} since {3}.", context.Request.RawUrl, NowTime.ToString("R"), req, StartupDate.ToString("R"));
            byte[] buffer = Encoding.UTF8.GetBytes(responseString);
            // and send it
            var response = context.Response;
            response.ContentType = "text/html";
            response.ContentLength64 = buffer.Length;
            response.StatusCode = 200;
            response.OutputStream.Write(buffer, 0, buffer.Length);
            response.OutputStream.Close();
        }
        protected static void SaveStatistics()
        {
            string folder = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string path = Path.Combine(folder, "statistics.xml");
            int PsGood = 12, PsBad = 123, InvGood = 1023, InvBad = 1;
            XmlTextWriter writer = new XmlTextWriter(path, null);
            //Write the root element
            writer.WriteStartElement("Statitics");
            //Write sub-elements
            writer.WriteElementString("Date", "16/10/12");
            writer.WriteElementString("PsGood", PsGood.ToString());
            writer.WriteElementString("PsBad", PsBad.ToString());
            writer.WriteElementString("InvGood", InvGood.ToString());
            writer.WriteElementString("InvBad", InvBad.ToString());
            // end the root element
            writer.WriteEndElement();

            //Write the XML to file and close the writer
            writer.Close();
        }
    }
}