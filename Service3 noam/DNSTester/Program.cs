using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Data.Odbc;
using BussinessLogicLayer;
using System.Globalization;

namespace DNSTester
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                string invoice = "20528774";
                GetFlexInvoice(invoice);
            }
            else
            {
                foreach (string invoice in args)
                {
                    GetFlexInvoice(invoice);
                }
            }
            //System.Timers.Timer timer1 = new System.Timers.Timer();
                //timer1.Interval = 1;
                //timer1.Elapsed += new System.Timers.ElapsedEventHandler(timer_1_Elapsed);
                ////timer1.AutoReset = false;
                //timer1.Start();
                //System.Timers.Timer timer2 = new System.Timers.Timer();
                //timer2.Interval = 5;
                //timer2.Elapsed += new System.Timers.ElapsedEventHandler(timer_2_Elapsed);
                ////timer2.AutoReset = false;
                //timer2.Start();
                //Console.WriteLine("Press any key to exit");
                //Console.ReadKey();
                //return;
        }
        private static void FindInvoice(string company, string inv_no)
        {
            string token = null;
            DOXAPI.DocType flexInvoiceType = null;
            DOXAPI.ServiceSoapClient dox = new DOXAPI.ServiceSoapClient();                   // Initialize web-service
            token = dox.Login("baanint", "fl3x8aan1n7", "Flex");
            if (token == null )
            {
                Console.WriteLine("Cannot login to DOXPRO");
                return;
            }
            Console.WriteLine("Login to DOXPRO succeeded, token is \"{0}\"", token);
            DOXAPI.DocType[] allTypes = dox.GetAllDocTypes(token);  // Get a list of doc-types from DOX-Pro
            // Find and keep the types used in this program
            foreach (DOXAPI.DocType dt in allTypes)
            {
                if (dt.Name == "Customer Invoice")
                {
                    flexInvoiceType = dox.GetDocType(token, dt.ID);
                    break;
                }
            }
            if (flexInvoiceType == null)
            {
                Console.WriteLine("Cannot obtain invoice DocType");
                return;
            }
            DOXAPI.SearchField[] fields = new DOXAPI.SearchField[1];

            DOXAPI.SearchField field = new DOXAPI.SearchField();
            field.FieldName = "Invoice No";
            field.SearchType = DOXAPI.SearchTypes.StartWith;
            field.FieldValue = company + inv_no;
            fields[0] = field;

            Console.WriteLine("Searching for Document in DOXPRO...");
            DOXAPI.TreeItemWithDocType[] invoices = dox.FindTreeItemWithDocType(token, fields, flexInvoiceType.DocTypeId);
            foreach (DOXAPI.TreeItemWithDocType ti in invoices)
            {
                // Fetch entity from result set
                DOXAPI.TreeItemWithDocType inv = dox.GetTreeItemWithDocType(token, ti);
                Console.WriteLine("Get Document ID={0}", inv.ID);
                DOXAPI.TreeItemWithDocType invFound = dox.GetTreeItemWithDocType(token, inv);
                DOXAPI.Document doc = invFound as DOXAPI.Document;
                if (doc != null)
                {
                    DOXAPI.Document foundDoc = dox.GetDocument(token, doc);
                    Console.WriteLine("FileName: {0}, FileName\"{1}\"", doc.FileName, foundDoc.FileName);
                }
                string url = dox.GetDocumentURL(token, inv.ID);
                Console.WriteLine("ID: {0}, url \"{1}\"", inv.ID, url);
            }
        }
        private static void GetFlexInvoice(string invoice)
        {
            FlexInvoiceExt inv = new FlexInvoiceExt(400, invoice);//"20500860"
          //  inv.FetchFromDB(new OdbcConnection("DSN=BAAN"), true);
            //inv.SerializeToPRM(@"c:\test\invoice.prm");
      //      inv.SerializeToXML(@"c:\test\SIS" + invoice + ".xml", invoice, false);
            //FindInvoice("400", invoice);
        }
        static void timer_1_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            Console.WriteLine("In Thread #1 - enter");
            System.Threading.Thread.Sleep(10000);
            System.Timers.Timer t = (System.Timers.Timer)sender;
            t.Enabled = true;
            Console.WriteLine("In Thread #1 - exit");
        }
        static void timer_2_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            Console.WriteLine("In Thread #2 - enter");
            System.Threading.Thread.Sleep(10000);
            Console.WriteLine("In Thread #2 - exit");
        }
        //public static bool FetchFromDB(ref InvoiceXML inv_xml, string CompanyCode, string InvoiceNumber)
        //{
        //    using (OdbcConnection DbConnection = new OdbcConnection("DSN=BAAN"))
        //    {
        //        try
        //        {
        //            if (DbConnection.State != System.Data.ConnectionState.Open)
        //            {
        //                DbConnection.Close();
        //                DbConnection.Open();
        //            }
        //        }
        //        catch (Exception ex)
        //        {
        //            Console.WriteLine("Could not connect to  BAANDB: " + ex.Message);
        //        }
        //        //Query #1
        //        using (OdbcCommand DbCommand = DbConnection.CreateCommand())
        //        {
        //            string cmd = "select t_ttyp, t_docd, t_rate, t_ccur, t_amth, t_amnt, t_vath from baandb.ttfacr200{0} where t_ninv={1}  and t_tdoc=\"\"";

        //            DbCommand.CommandText = String.Format(cmd, CompanyCode, InvoiceNumber);
        //            OdbcDataReader reader;
        //            try
        //            {
        //                reader = DbCommand.ExecuteReader();
        //            }
        //            catch (Exception ex)
        //            {
        //                Console.WriteLine("Query #1: " + DbCommand.CommandText + "\n" + ex.Message, System.Diagnostics.EventLogEntryType.Error);
        //                return false;
        //            }
        //            reader.Read();
        //            {
        //                if (inv_xml.Envelope == null)
        //                {
        //                    inv_xml.Envelope = new InvoiceXMLEnvelope();
        //                }
        //                if (inv_xml.Envelope.Header == null)
        //                {
        //                    inv_xml.Envelope.Header = new InvoiceXMLEnvelopeHeader();
        //                }
        //                inv_xml.Envelope.DocType = "MMINVC";
        //                inv_xml.Envelope.MessageDate = DateTime.Now.ToShortDateString();
        //                inv_xml.Envelope.MessageTime = DateTime.Now.ToUniversalTime().ToShortTimeString();
        //                //inv_xml.Envelope.Header.InvoiceType = ;
        //                inv_xml.Envelope.Header.Date = reader.GetDateTime(1);
        //                inv_xml.Envelope.Header.ExchangeRate = reader.GetFloat(2).ToString();
        //                inv_xml.Envelope.Header.Currency = reader.GetString(3).TrimEnd();
        //                inv_xml.Envelope.Header.CurrencyDocSum = reader.GetFloat(4).ToString();
        //                inv_xml.Envelope.Header.DocSum = reader.GetFloat(5).ToString();
        //                inv_xml.Envelope.Header.TaxSum = reader.GetFloat(6).ToString();
        //            }
        //            reader.Close();
        //        }
        //        //Query #1A
        //        using (OdbcCommand DbCommand = DbConnection.CreateCommand())
        //        {
        //            string cmd = "select first 1 t_pvat from baandb.ttcmcs032{0} where t_edat<=\"{1}\" and t_cvat=\"      001\" order by t_edat";
        //            DbCommand.CommandText = String.Format(cmd, CompanyCode, inv_xml.Envelope.Header.Date.ToString("d", CultureInfo.CreateSpecificCulture("en-US")));
        //            OdbcDataReader reader;
        //            try
        //            {
        //                reader = DbCommand.ExecuteReader();
        //            }
        //            catch (Exception ex)
        //            {
        //                Console.WriteLine("Query #1: " + DbCommand.CommandText + "\n" + ex.Message, System.Diagnostics.EventLogEntryType.Error);
        //                return false;
        //            }
        //            reader.Read();
        //            {
        //                inv_xml.Envelope.Header.TaxRate = reader.GetString(0).TrimEnd();
        //            }
        //        }
        //        //Query #2
        //        using (OdbcCommand DbCommand = DbConnection.CreateCommand())
        //        {
        //            string cmd = "select t_nama, t_vatn from baandb.ttccom000{0} where t_ncmp={0}";

        //            DbCommand.CommandText = String.Format(cmd, CompanyCode);
        //            OdbcDataReader reader;
        //            try
        //            {
        //                reader = DbCommand.ExecuteReader();
        //            }
        //            catch (Exception ex)
        //            {
        //                Console.WriteLine("Query #2: " + DbCommand.CommandText + "\n" + ex.Message, System.Diagnostics.EventLogEntryType.Error);
        //                return false;
        //            }
        //            reader.Read();
        //            {
        //                inv_xml.Envelope.Header.SupplierName = reader.GetString(0).TrimEnd();
        //                inv_xml.Envelope.Header.SupplierNo = reader.GetString(1).TrimEnd();
        //            }
        //            reader.Close();
        //        }
        //        //Query #3
        //        using (OdbcCommand DbCommand = DbConnection.CreateCommand())
        //        {
        //            string cmd = "select t_fovn , t_nama, t_namc, t_name, t_ccty, t_pstc from baandb.ttccom010{0} where t_cuno in (select t_cuno from baandb.ttdsls040{0} where t_orno in (select distinct t_orno from baandb.ttdsls045{0} as DSLS045 where DSLS045.t_invn={1}))";

        //            DbCommand.CommandText = String.Format(cmd, CompanyCode, InvoiceNumber);
        //            OdbcDataReader reader;
        //            try
        //            {
        //                reader = DbCommand.ExecuteReader();
        //            }
        //            catch (Exception ex)
        //            {
        //                Console.WriteLine("Query #3: " + DbCommand.CommandText + "\n" + ex.Message, System.Diagnostics.EventLogEntryType.Error);
        //                return false;
        //            }
        //            reader.Read();
        //            {
        //                if (inv_xml.Envelope.Header.BillTo == null)
        //                {
        //                    inv_xml.Envelope.Header.BillTo = new InvoiceXMLEnvelopeHeaderBillTo();
        //                }
        //                inv_xml.Envelope.Header.PrivateCompanyCode = reader.GetString(0).TrimEnd();
        //                inv_xml.Envelope.Header.BillTo.CompanyName = reader.GetString(1).TrimEnd();
        //                inv_xml.Envelope.Header.BillTo.Address = reader.GetString(2).TrimEnd();
        //                inv_xml.Envelope.Header.BillTo.City = reader.GetString(3).TrimEnd();
        //                inv_xml.Envelope.Header.BillTo.Country = reader.GetString(4).TrimEnd();
        //                inv_xml.Envelope.Header.BillTo.ZipCode = reader.GetString(5).TrimEnd();
        //            }
        //            reader.Close();
        //        }
        //        //Query #4
        //        using (OdbcCommand DbCommand = DbConnection.CreateCommand())
        //        {

        //            //select t_dsca from ttcmcs013400 where pay = ???
        //            string cmd = "select DSLS041.t_epos, DSLS045.t_dqua, DSLS045.t_pric, DSLS045.t_amnt, DSLS045.t_item, IITM001.t_dsca, TTCMCS041.t_dsca, DSLS040.t_refa, DSLS040.t_orno, DSLS040.t_odat, TCMCS013.t_dsca, DSLS040.t_orno from baandb.ttdsls045{0} as DSLS045, baandb.ttiitm001{0} as IITM001, baandb.ttdsls040{0} as DSLS040, baandb.ttdsls041{0} as DSLS041, baandb.ttcmcs041{0} as TTCMCS041, baandb.ttcmcs013{0} as TCMCS013 where  DSLS045.t_invn={1} and IITM001.t_item=DSLS045.t_item and DSLS041.t_item=DSLS045.t_item and DSLS040.t_orno=DSLS045.t_orno and (DSLS041.t_pono=DSLS045.t_pono and DSLS041.t_orno=DSLS045.t_orno) and TTCMCS041.t_cdec=DSLS040.t_cdec and TCMCS013.t_cpay=DSLS040.t_cpay";
        //            DbCommand.CommandText = String.Format(cmd, CompanyCode, InvoiceNumber);
        //            OdbcDataReader reader;
        //            try
        //            {
        //                reader = DbCommand.ExecuteReader();
        //            }
        //            catch (Exception ex)
        //            {
        //                Console.WriteLine("Query #4: " + DbCommand.CommandText + "\n" + ex.Message, System.Diagnostics.EventLogEntryType.Error);
        //                return false;
        //            }
        //            int num_lines = 0;
                    
        //            while (reader.Read())
        //            {
        //                num_lines++;
        //                InvoiceXMLEnvelopeLine line = new InvoiceXMLEnvelopeLine();

        //                line.LineNo = reader.GetString(0).TrimEnd();
        //                line.UnitsQty = reader.GetInt32(1).ToString();
        //                line.ItemPriceBruto = reader.GetFloat(2).ToString();
        //                line.LineSum = reader.GetFloat(3).ToString();
        //                line.GoodsBarcode = reader.GetString(4).TrimEnd();
        //                line.ItemDesc = reader.GetString(5).TrimEnd();
        //                //line.Delivery = reader.GetString(6);
        //                //line.ReferenceType = reader.GetString(7);
        //                //line.ReferenceNumber = reader.GetInt32(8);
        //                //line.ReferenceDate = reader.GetDate(9);
        //                //line.PaymentTerms = reader.GetString(10);
        //                //line.SalesOrder = reader.GetString(11);
        //                //line.PartNumber = "";
        //                //Add(line);
        //                if (inv_xml.Envelope.Details == null)
        //                {
        //                    inv_xml.Envelope.Details = new InvoiceXMLEnvelopeLine[1];
        //                    inv_xml.Envelope.Details[0] = line;
        //                }
        //                else
        //                {
        //                    InvoiceXMLEnvelopeLine[] lines = inv_xml.Envelope.Details;
        //                    Array.Resize(ref lines, inv_xml.Envelope.Details.Length + 1);
        //                    lines[num_lines - 1] = line;
        //                    inv_xml.Envelope.Details = lines;
        //                }
        //            }
        //            inv_xml.Envelope.Header.NumOfLines = num_lines.ToString();
        //            reader.Close();
        //            System.IO.FileStream fs = System.IO.File.Create(@"c:\test\invoice.xml");
        //            System.Xml.Serialization.XmlSerializer x = new System.Xml.Serialization.XmlSerializer(inv_xml.GetType());
        //            //x.Serialize(Console.Out, inv_xml);
        //            //Console.WriteLine();
        //            x.Serialize(fs, inv_xml);
        //        }
        //    }
        //    return true;
        //}
    }
}
