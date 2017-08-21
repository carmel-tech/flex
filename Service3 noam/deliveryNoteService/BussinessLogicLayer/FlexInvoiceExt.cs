using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.Odbc;
using System.Globalization;
using System.Xml;
using System.IO;

namespace BussinessLogicLayer
{
    public class FlexInvoiceExt
    {
        private System.IO.FileStream st;
        /// <summary>
        ///  Methods
        /// </summary>
        public FlexInvoiceExt(int company_code, string invoice_num)
        {
            CompanyCode = company_code;
            InvoiceNumber = invoice_num;
            MsgDateTime = DateTime.Now                ;
        }
        private FlexInvoiceExt()//to prohibit void construction
        {
        }
//        public bool FetchFromDB(string dsn = "DSN=BAAN", bool debug = false)
        public bool FetchFromDB(OdbcConnection DbConnection, bool debug = true)
        {
            if (debug) Console.WriteLine("0");
  //          using (OdbcConnection DbConnection = new OdbcConnection(dsn))
            {
                if (debug) Console.WriteLine("1");
                
                //try
                //{
                //    if (DbConnection.State != System.Data.ConnectionState.Open)
                //    {
                //        DbConnection.Close();
                //        DbConnection.Open();
                //    }
                //}
                //catch (Exception ex)
                //{
                //    string msg = "Could not connect to  BAANDB: " + ex.Message;
                //    if (debug) 
                //    {
                //        Console.WriteLine(msg);
                //    }
                //    else
                //    {
                //        throw new System.Exception(msg);
                //    }
                //}
                //if (debug) Console.WriteLine("2");
                //Query #1
                using (OdbcCommand DbCommand = DbConnection.CreateCommand())
                {
                    string cmd = "select t_ttyp, t_docd, t_rate, t_ccur, t_amth, t_amnt, t_vath from baandb.ttfacr200{0} where t_ninv={1}  and t_tdoc=\"\"";

                    DbCommand.CommandText = String.Format(cmd, CompanyCode, InvoiceNumber);
                    OdbcDataReader reader;
                    if (debug) Console.WriteLine("3");
                    try
                    {
                        reader = DbCommand.ExecuteReader();
                    }
                    catch (Exception ex)
                    {
                        string msg = "Query #1: " + DbCommand.CommandText + "\n" + ex.Message;
                        if (debug)
                        {
                            Console.WriteLine(msg);
                        }
                        else
                        {
                            throw new System.Exception(msg);
                        }
                        return false;
                    }
                    if (debug) Console.WriteLine("4");
                    if (reader.Read())
                    {
                        InvoiceType = reader.GetString(0);
                        InvoiceDate = reader.GetDateTime(1);
                        ExchangeRate = reader.GetDecimal(2);
                        Currentcy = reader.GetString(3);
                        CurrencyDocumentSum = reader.GetDecimal(4);
                        DocumentSum = reader.GetDecimal(5);
                        TaxSum = reader.GetDecimal(6);
                        LineSum = CurrencyDocumentSum - TaxSum;
                        if (debug)
                        {
                            Console.WriteLine("Invoice #: " + InvoiceNumber );
                            Console.WriteLine("InvoiceType: " + InvoiceType);
                            Console.WriteLine("InvoiceDate: " + InvoiceDate.ToShortDateString());
                            Console.WriteLine("ExchangeRate: " + ExchangeRate);
                            Console.WriteLine("Currentcy: " + Currentcy);
                            Console.WriteLine("CurrencyDocumentSum: " + CurrencyDocumentSum);
                            Console.WriteLine("DocumentSum: " + DocumentSum);
                            Console.WriteLine("LineSum: " + LineSum);
                            Console.WriteLine("TaxSum: " + TaxSum);
                        }
                    }
                    else
                    {
                        if (debug) Console.WriteLine("5");
                        return false;
                    }
                    reader.Close();
                    
                    cmd = "select first 1 t_pvat from baandb.ttcmcs032{0} where t_edat<=\"{1}\" and t_cvat=\"      001\" order by t_edat desc";
                    if (debug) Console.WriteLine("6");
                    string dt = InvoiceDate.ToString("d", CultureInfo.CreateSpecificCulture("en-US"));
                    DbCommand.CommandText = String.Format(cmd, CompanyCode, dt);
                    try
                    {
                        reader = DbCommand.ExecuteReader();
                    }
                    catch (Exception ex)
                    {
                        string msg = "Query #1: " + DbCommand.CommandText + "\n" + ex.Message;
                        if (debug) Console.WriteLine("7");
                        if (debug)
                        {
                            Console.WriteLine(msg);
                        }
                        else
                        {
                            throw new System.Exception(msg);
                        }
                        return false;
                    }
                    if ( reader.Read() )
                    {
                        if (debug) Console.WriteLine("8");
                        TaxRate = reader.GetString(0);
                        if (debug)
                        {
                            Console.WriteLine("TaxRate: " + TaxRate);
                        }
                    }
                    else
                    {
                        if (debug) Console.WriteLine("9");
                        return false;
                    }
                    reader.Close();
                    //Query #2
                    cmd = "select t_nama, t_vatn from baandb.ttccom000{0} where t_ncmp={0}";
                    if (debug) Console.WriteLine("10");

                    DbCommand.CommandText = String.Format(cmd, CompanyCode);
                    try
                    {
                        reader = DbCommand.ExecuteReader();
                    }
                    catch (Exception ex)
                    {
                        if (debug) Console.WriteLine("11");

                        string msg = "Query #2: " + DbCommand.CommandText + "\n" + ex.Message;
                        if (debug)
                        {
                            Console.WriteLine(msg);
                        }
                        else
                        {
                            if (debug) Console.WriteLine("12");
                            throw new System.Exception(msg);
                        }
                        return false;
                    }
                    if ( reader.Read() )
                    {
                        if (debug) Console.WriteLine("13");

                        SupplierName = reader.GetString(0);
                        SupplierPrivateCompanyCode = reader.GetString(1);
                        if (debug)
                        {
                            Console.WriteLine("SupplierName: " + SupplierName);
                            Console.WriteLine("SupplierPrivateCompanyCode: " + SupplierPrivateCompanyCode);
                        }
                    }
                    else
                    {
                        if (debug) Console.WriteLine("14");
                        return false;
                    }
                    reader.Close();
                    //Query #3                    
                    cmd = "select t_fnum from baandb.ttccom810{0} where t_cuno in (select t_cuno from baandb.ttdsls040{0} where t_orno in (select distinct t_orno from baandb.ttdsls045{0} as DSLS045 where DSLS045.t_invn={1}))";
                    if (debug) Console.WriteLine("15");

                    DbCommand.CommandText = String.Format(cmd, CompanyCode, InvoiceNumber);
                    try
                    {
                        reader = DbCommand.ExecuteReader();                       
                    }
                        
                    catch (Exception ex)
                    {
                        if (debug) Console.WriteLine("16");
                        string msg = "Query #3: " + DbCommand.CommandText + "\n" + ex.Message;
                        if (debug)
                        {
                            Console.WriteLine(msg);
                        }
                        else
                        {
                            throw new System.Exception(msg);
                        }
                        return false;
                    }
                    if (reader.Read())
                    {                       
                        if (debug) Console.WriteLine("17");
                        Receiver = reader.GetString(0);                        
                        if (debug)
                        {
                            Console.WriteLine("Receiver: " + Receiver);

                        }
                    }
                    else
                    {
                        if (debug) Console.WriteLine("18");

                        return false;
                    }
                    reader.Close();
                //Query #4
                    cmd = "select t_fovn , t_nama, t_namc, t_name, t_ccty, t_pstc from baandb.ttccom010{0} where t_cuno in (select t_cuno from baandb.ttdsls040{0} where t_orno in (select distinct t_orno from baandb.ttdsls045{0} as DSLS045 where DSLS045.t_invn={1}))";
                    if (debug) Console.WriteLine("19");

                    DbCommand.CommandText = String.Format(cmd, CompanyCode, InvoiceNumber);
                    try
                    {
                        reader = DbCommand.ExecuteReader();
                    }
                    catch (Exception ex)
                    {
                        if (debug) Console.WriteLine("20");
                        string msg = "Query #4: " + DbCommand.CommandText + "\n" + ex.Message;
                        if (debug)
                        {
                            Console.WriteLine(msg);
                        }
                        else
                        {
                            throw new System.Exception(msg);
                        }
                        return false;
                    }
                    if ( reader.Read() )
                    {
                        if (debug) Console.WriteLine("21");

                        RetailerPrivateCompanyCode = reader.GetString(0);
                        CompanyName = reader.GetString(1);
                        Address = reader.GetString(2);
                        City = reader.GetString(3);
                        Country = reader.GetString(4);
                        Zipcode = reader.GetString(5);
                      //  Receiver =  reader.GetString(6);
                        if (debug)
                        {
                            Console.WriteLine("Sender: " + Sender);
                            Console.WriteLine("Receiver: " + Receiver);
                            Console.WriteLine("RetailerPrivateCompanyCode: " + RetailerPrivateCompanyCode);
                            Console.WriteLine("CompanyName: " + CompanyName);
                            Console.WriteLine("Address: " + Address);
                            Console.WriteLine("City: " + City);
                            Console.WriteLine("Country: " + Country);
                            Console.WriteLine("Zipcode: " + Zipcode);
                        }
                    }
                    else
                    {
                        if (debug) Console.WriteLine("22");

                        return false;
                    }
                    reader.Close();
                    //Query #5
                    //select t_dsca from ttcmcs013400 where pay = ???
                    //cmd = "select DSLS041.t_epos, DSLS045.t_dqua, DSLS045.t_pric, DSLS045.t_amnt, DSLS045.t_item, IITM001.t_dsca, TTCMCS041.t_dsca, DSLS040.t_refa, DSLS040.t_orno, DSLS040.t_odat, TCMCS013.t_dsca, DSLS040.t_orno, DSLS045.t_damt_1 from baandb.ttdsls045{0} as DSLS045, baandb.ttiitm001{0} as IITM001, baandb.ttdsls040{0} as DSLS040, baandb.ttdsls041{0} as DSLS041, baandb.ttcmcs041{0} as TTCMCS041, baandb.ttcmcs013{0} as TCMCS013 where  DSLS045.t_invn={1} and IITM001.t_item=DSLS045.t_item and DSLS041.t_item=DSLS045.t_item and DSLS040.t_orno=DSLS045.t_orno and (DSLS041.t_pono=DSLS045.t_pono and DSLS041.t_orno=DSLS045.t_orno) and TTCMCS041.t_cdec=DSLS040.t_cdec and TCMCS013.t_cpay=DSLS040.t_cpay";
                    cmd = "select DSLS041.t_epos, DSLS045.t_dqua, DSLS045.t_pric, DSLS045.t_amnt, DSLS045.t_item, IITM001.t_dsca, TTCMCS041.t_dsca, DSLS040.t_cotp, DSLS040.t_refa, DSLS040.t_odat, TCMCS013.t_dsca, DSLS040.t_orno, DSLS045.t_damt_1 from baandb.ttdsls045{0} as DSLS045, baandb.ttiitm001{0} as IITM001, baandb.ttdsls040{0} as DSLS040, baandb.ttdsls041{0} as DSLS041, baandb.ttcmcs041{0} as TTCMCS041, baandb.ttcmcs013{0} as TCMCS013 where  DSLS045.t_invn={1} and IITM001.t_item=DSLS045.t_item and DSLS041.t_item=DSLS045.t_item and DSLS040.t_orno=DSLS045.t_orno and (DSLS041.t_pono=DSLS045.t_pono and DSLS041.t_orno=DSLS045.t_orno) and TTCMCS041.t_cdec=DSLS040.t_cdec and TCMCS013.t_cpay=DSLS040.t_cpay";
                    DbCommand.CommandText = String.Format(cmd, CompanyCode, InvoiceNumber);
                    if (debug) Console.WriteLine("23");
                    
                    try
                    {
                        reader = DbCommand.ExecuteReader();
                    }
                    catch (Exception ex)
                    {
                        if (debug) Console.WriteLine("24");
                        string msg = "Query #5: " + DbCommand.CommandText + "\n" + ex.Message;
                        if (debug)
                        {
                            Console.WriteLine(msg);
                        }
                        else
                        {
                            throw new System.Exception(msg);
                        }
                        return false;
                    }
                    int num_lines = 0;
                    if (debug) Console.WriteLine("25");
                    OdbcDataReader readerCustBarcode;
                    OdbcCommand DbCommandCustBarcode;
                    while (reader.Read())
                    {
                        num_lines++;
                        InvoiceLine line = new InvoiceLine();

                        line.LineNo = reader.GetString(0);
                        if (String.IsNullOrEmpty(line.LineNo))
                        {
                            line.LineNo = "1";
                        }
                        //line.UnitQuantity = reader.GetInt32(1);
                        //line.UnitQuantity =Math.Round(reader.GetDouble(1),2);
                        string qty = reader.GetDouble(1).ToString();
                      if (qty.IndexOf('.') != -1)
                          line.UnitQuantity=qty.Substring(0, qty.IndexOf('.') + 3);
                        else
                          line.UnitQuantity= qty;

                        line.ItemPriceBruto = reader.GetDecimal(2);
                        line.LineSum = reader.GetDecimal(3);
                        line.ItemBarcode = reader.GetString(4);
                        line.ItemDescription = reader.GetString(5);
                        line.Delivery = reader.GetString(6);
                        line.ReferenceType = reader.GetString(7);
                        line.ReferenceNumber = reader.GetString(8);
                        line.ReferenceDate = reader.GetDate(9);
                        line.PaymentTerms = reader.GetString(10);
                        line.SalesOrder = reader.GetString(11);
                        line.Discount = reader.GetString(12);
                        line.PartNumber = "";
                        #region getCustomerBarcode
                            //Query #5.5 added by Liat for CustomerBarcode
                        cmd = "select TTII.t_mitm from baandb.ttiitm950{0} as TTII where TTII.t_mnum=\"999\" and TTII.t_item=\"{1}\" and  TTII.t_cuno in (select t_cuno from baandb.ttccom810{0} where t_fnum={2})";
                            if (debug) Console.WriteLine("26");
                            
                            DbCommandCustBarcode  = DbConnection.CreateCommand();
                            string itemBarcode=DbCommandCustBarcode.CommandText = String.Format(cmd, CompanyCode, line.ItemBarcode.PadLeft(16,' ') , Receiver);
                            
                            try
                            {
                                readerCustBarcode = DbCommandCustBarcode.ExecuteReader();
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.EventLog.WriteEntry("error in query 5.5", ex.ToString());
                                if (debug) Console.WriteLine("27");
                                string msg = "Query #5.5: " + DbCommand.CommandText + "\n" + ex.Message;
                                if (debug)
                                {
                                    Console.WriteLine(msg);
                                }
                                else
                                {
                                    throw new System.Exception(msg);
                                }
                                return false;
                            }
                            if (readerCustBarcode.Read())
                            {
                                if (debug) Console.WriteLine("28");
                                line.CustomerBarcode = readerCustBarcode.GetString(0);
                                System.Diagnostics.EventLog.WriteEntry("CustomerBarcode", "barcode exists: " + line.CustomerBarcode + "this is the query: " + DbCommand.CommandText);
                                if (debug)
                                {
                                    Console.WriteLine("CustomerBarcode: " + line.CustomerBarcode);
                                }
                            }
                            else
                            {                                
                                System.Diagnostics.EventLog.WriteEntry("CustomerBarcode", "Customer Barcode wasnt found this is the query"  + cmd);
                                if (debug) Console.WriteLine("29");

                             //   return false;
                            }
                            readerCustBarcode.Close();

                            
                            //Query #6 added by Liat for TaxSumNIS
                            cmd = "select t_rate_c from baandb.ttfgld018{0} where  t_docn={1} and t_ttyp=\"{2}\" ";
                            if (debug) Console.WriteLine("30");
                            OdbcDataReader readerSumNIS;
                            OdbcCommand DbCommandSumNIS;
                            DbCommandSumNIS = DbConnection.CreateCommand();
                            DbCommandSumNIS.CommandText = String.Format(cmd, CompanyCode, InvoiceNumber, InvoiceType);
                            try
                            {
                                readerSumNIS = DbCommandSumNIS.ExecuteReader();
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.EventLog.WriteEntry("sum", ex.ToString());
                                System.Diagnostics.EventLog.WriteEntry("query", DbCommandSumNIS.CommandText);
                                string msg = "Query #6: " + DbCommandSumNIS.CommandText + "\n" + ex.Message;
                                if (debug) Console.WriteLine("31");
                                if (debug)
                                {
                                    Console.WriteLine(msg);
                                }
                                else
                                {
                                    throw new System.Exception(msg);
                                }
                                return false;
                            }
                            if (readerSumNIS.Read())
                            {
                                System.Diagnostics.EventLog.WriteEntry("TaxSumNIS-query", "TaxSumNIS exists: this is the query: " + DbCommandSumNIS.CommandText);
                                
                                try
                                {
                                    CurrencyRate = readerSumNIS.GetDecimal(0);
                                    TaxSumNIS = CurrencyRate * TaxSum;
                                    System.Diagnostics.EventLog.WriteEntry("TaxSumNIS", "CurrencyRate: " + CurrencyRate + "TaxSum: " + TaxSum + "TaxSumNIS: " + TaxSumNIS + "this is the query: " + DbCommandSumNIS.CommandText);
                                }
                                catch (Exception ex)
                                {
                                    System.Diagnostics.EventLog.WriteEntry("error in kefel", ex.ToString());
                                    System.Diagnostics.EventLog.WriteEntry("error in kefel values", "CurrencyRate: " + CurrencyRate + "TaxSum: " + TaxSum + "TaxSumNIS: " + TaxSumNIS + "this is the query: " + DbCommandSumNIS.CommandText);
                                }
 
                            }
                            else
                            {
                                System.Diagnostics.EventLog.WriteEntry("TaxSumNIS", "TaxSumNIS doesnt exist this is the query: " + DbCommandSumNIS.CommandText);
                                if (debug) Console.WriteLine("33");
                                return false;
                            }
                            readerSumNIS.Close();
                     

                        
                        #endregion
                        Add(line);
                        if (debug)
                        {
                            Console.WriteLine("..................................................................................");
                            Console.WriteLine("Line #: " + num_lines);
                            Console.WriteLine("Customer lineNo: " + line.LineNo);
                            Console.WriteLine("UnitQuantity: " + line.UnitQuantity);
                            Console.WriteLine("ItemPriceBruto: " + line.ItemPriceBruto);
                            Console.WriteLine("LineSum: " + line.LineSum);
                            Console.WriteLine("PartNumber: " + line.PartNumber);
                            Console.WriteLine("ItemDescription: " + line.ItemDescription);
                            Console.WriteLine("Delivery: " + line.Delivery);
                            Console.WriteLine("ReferenceType: " + line.ReferenceType);
                            Console.WriteLine("ReferenceNumber: " + line.ReferenceNumber);
                            Console.WriteLine("ReferenceDate: " + line.ReferenceDate.ToShortDateString());
                            Console.WriteLine("PaymentTerms: " + line.PaymentTerms);
                            Console.WriteLine("SalesOrder: " + line.SalesOrder);
                            Console.WriteLine("Discount: " + line.Discount);
                        }
                    }
                    if ( debug ) Console.WriteLine("..................................................................................");
                    TotalLines = num_lines;
                    reader.Close();
                }
                if (debug) Console.WriteLine("exit");
            }
            return true;
        }
        public bool CheckIsCompnayExistInDB(OdbcConnection DbConnection, string CompanyCode, bool debug = true)
        {
            if (debug) Console.WriteLine("0");
            
            if (debug) Console.WriteLine("1");

            using (OdbcCommand DbCommand = DbConnection.CreateCommand())
            {
                OdbcDataReader reader;  
                try
                {
                    string cmd = "select first 1 t_cuno from baandb.ttccom810400 where t_cuno={0}";                 
                    DbCommand.CommandText = String.Format(cmd, CompanyCode);                             
                    reader = DbCommand.ExecuteReader();
               
                }
                catch (Exception ex)
                {
                    System.Diagnostics.EventLog.WriteEntry("dms error in db", ex.ToString());
                    string msg = "Query #1: " + DbCommand.CommandText + "\n" + ex.Message;
                    if (debug)
                    {
                        Console.WriteLine(msg);
                    }
                    else
                    {
                        throw new System.Exception(msg);
                    }
                    return false;
                }
                bool isCompanyExist;
                if (debug) Console.WriteLine("4");
                if (reader.Read())
                {                    
                    isCompanyExist= true;
                }
                else
                {
                    if (debug) Console.WriteLine("5");
                    isCompanyExist= false;
                }
                reader.Close();
                return isCompanyExist;
            }           
        }
        public bool SerializeToPRM(string prm_path)
        {
            using (st = System.IO.File.OpenWrite(prm_path))
            {
                WriteToPRM("DocType", DocType);
                WriteToPRM("Sender", Sender);
                WriteToPRM("Receiver", Receiver);
                //WriteToPRM("MsgDateTime", MsgDateTime.ToShortDateString());
                WriteToPRM("MsgDate", MsgDate);
                WriteToPRM("MsgTime", MsgTime);
                WriteToPRM("InvoiceType", InvoiceType);
                WriteToPRM("InvoiceNumber", InvoiceNumber);
                WriteToPRM("SupplierName", SupplierName);
                WriteToPRM("DiscountAmount", DiscountAmount);
                WriteToPRM("CompanyCode", CompanyCode.ToString());
                WriteToPRM("InvoiceDate", InvoiceDate.ToShortDateString());
                WriteToPRM("ExchangeRate", ExchangeRate.ToString());
                WriteToPRM("Currentcy", Currentcy);
                WriteToPRM("TaxInvoice", TaxInvoice);
                WriteToPRM("DocumentInvoiceType", DocumentInvoiceType);
                WriteToPRM("SupplierPrivateCompanyCode", SupplierPrivateCompanyCode);
                WriteToPRM("RetailerPrivateCompanyCode", RetailerPrivateCompanyCode);
                WriteToPRM("CurrencyDocumentSum", CurrencyDocumentSum.ToString());
                WriteToPRM("DocumentSum", DocumentSum.ToString());
                WriteToPRM("TaxSum", TaxSum.ToString());
                WriteToPRM("TaxSumNIS", TaxSumNIS.ToString());
                WriteToPRM("CurrencyRate", CurrencyRate.ToString());  
                WriteToPRM("TaxRate", TaxRate);
                WriteToPRM("CompanyName", CompanyName);
                WriteToPRM("Address", Address);
                WriteToPRM("City", City);
                WriteToPRM("State", State);
                WriteToPRM("Country", Country);
                WriteToPRM("POB", POB);
                WriteToPRM("Zipcode", Zipcode);
                WriteToPRM("Bank", Bank);
                WriteToPRM("Account", Account);
                WriteToPRM("Details", Details);
                WriteToPRM("TotalLines", TotalLines.ToString());
                foreach (InvoiceLine line in Lines)
                {
                    WriteToPRM(line);
                }
            }
            return true;
        }
        public bool SerializeToXML(string path,string packingSlip, bool debug=true)
        {
            InvoiceXML inv_xml = new InvoiceXML();
            if (inv_xml.Envelope == null)
            {
                inv_xml.Envelope = new InvoiceXMLEnvelope();
            }
            if (inv_xml.Envelope.Header == null)
            {
                inv_xml.Envelope.Header = new InvoiceXMLEnvelopeHeader();
            }
            #region Note for auto-generated Invoice.cs
            // Note 1: For some reasons standalone attribute minOccurs brings problems - some fields are not serialized, so we have to remove them manually from XSD schema before class generation
            // Note 2: if Invoice.xsd was changed and Invoice.cs was re-geenrated using xsd.exe => the field below should be changed accordingly
            /// <remarks>
            /// This property was changed manually in order to control its output format 
            /// </remarks>
            //  //[System.Xml.Serialization.XmlElementAttribute(DataType="time")]
            //public string MessageTime
            //{
            //    get
            //    {
            //        return System.String.Format("{0:HH:mm:ss}", this.messageTimeField);
            //    }
            //    set
            //    {
            //        this.messageTimeField = System.DateTime.Parse(value);
            //    }
            //}
            #endregion
            inv_xml.Envelope.MessageDate = System.String.Format("{0:yyyy-MM-dd}", DateTime.Now);
            inv_xml.Envelope.MessageTime = DateTime.Now.ToShortTimeString();
            inv_xml.Envelope.Sender = Sender;
            inv_xml.Envelope.Receiver = Receiver;
            inv_xml.Envelope.Header.Account = Account;
            inv_xml.Envelope.Header.Address = Address;
            inv_xml.Envelope.Header.Bank = Bank;
            inv_xml.Envelope.Header.City = City;
            inv_xml.Envelope.Header.CompanyCode = CompanyCode.ToString().TrimEnd();
            inv_xml.Envelope.Header.CompanyName = CompanyName;
            inv_xml.Envelope.Header.Country = Country;
            inv_xml.Envelope.Header.Currency = Currentcy;
            inv_xml.Envelope.Header.CurrencyDocSum = CurrencyDocumentSum.ToString();
//            inv_xml.Envelope.Header.Delivery = "";
            inv_xml.Envelope.Header.DiscountAmount = DiscountAmount;
            inv_xml.Envelope.Header.DocInvoiceType = InvoiceXMLEnvelopeHeaderDocInvoiceType.Original;//DocumentInvoiceType.ToString();
            inv_xml.Envelope.Header.DocSum = DocumentSum;
            inv_xml.Envelope.Header.ExchangeRate = ExchangeRate.ToString();
            inv_xml.Envelope.Header.InvoiceDate = System.String.Format("{0:yyyy-MM-dd}", InvoiceDate);
            inv_xml.Envelope.Header.InvoiceNo = InvoiceNumber;
            inv_xml.Envelope.Header.InvoiceType = InvoiceType;
            inv_xml.Envelope.Header.LineSum = LineSum;
            inv_xml.Envelope.Header.NumOfLines = TotalLines.ToString();
//            inv_xml.Envelope.Header.PaymentTerms = "";
            inv_xml.Envelope.Header.POB = POB;
            inv_xml.Envelope.Header.Reference = new InvoiceXMLEnvelopeHeaderReference[2/*TotalLines*/];
            inv_xml.Envelope.Header.RetailerPrivateCompanyCode = RetailerPrivateCompanyCode;
            inv_xml.Envelope.Header.State = State;
            inv_xml.Envelope.Header.SupplierName = SupplierName;
            inv_xml.Envelope.Header.SupplierPrivateCompanyCode = SupplierPrivateCompanyCode;
            inv_xml.Envelope.Header.TaxInvoice = InvoiceXMLEnvelopeHeaderTaxInvoice.TaxInvoice;

            string CurrencyRateStr = CurrencyRate.ToString();
            if (CurrencyRateStr.IndexOf(".") > 0)
            {
                if (CurrencyRateStr.IndexOf(".") + 5 <= CurrencyRateStr.Length)
                {
                    CurrencyRateStr = CurrencyRateStr.Substring(0, CurrencyRateStr.IndexOf(".") + 5);
                }
            }
            string TaxSumNISStr = TaxSumNIS.ToString();
            if (TaxSumNISStr.IndexOf(".") > 0)
            {
                if (TaxSumNISStr.IndexOf(".") + 3 <= TaxSumNISStr.Length)
                {
                    TaxSumNISStr = TaxSumNISStr.Substring(0, TaxSumNISStr.IndexOf(".") + 3);
                }
            }
            
            inv_xml.Envelope.Header.TaxRate = TaxRate == null ? 0: System.Decimal.Parse(TaxRate);
            inv_xml.Envelope.Header.TaxSum = TaxSum;
            inv_xml.Envelope.Header.TaxSumNIS = TaxSumNISStr;
            inv_xml.Envelope.Header.CurrencyRate = CurrencyRateStr;
            inv_xml.Envelope.Header.Zipcode = Zipcode;

            //Nasty hack
            inv_xml.Envelope.Header.SNAttachName = System.IO.Path.GetFileName(path).Replace(".xml", ".pdf"); ;
            inv_xml.Envelope.Details = new InvoiceXMLEnvelopeLine[TotalLines];

            int counter=0;
            foreach (InvoiceLine line in Lines)
            {
                //Console.WriteLine(String.Format("Preparing line #{0} from {1}:", counter, TotalLines));
                //Console.Out.Flush();

                if (counter == 0)
                {
                    InvoiceXMLEnvelopeHeaderReference Reference = new InvoiceXMLEnvelopeHeaderReference();
                    Reference.RefDate = System.String.Format("{0:yyyy-MM-dd}",line.ReferenceDate);
                    Reference.RefNo = line.ReferenceNumber;
                    Reference.RefType = InvoiceXMLEnvelopeHeaderReferenceRefType.purchaseOrder;
                    inv_xml.Envelope.Header.Reference[counter] = Reference;
                 //22.4 added by liat form shipment reference   
                    InvoiceXMLEnvelopeHeaderReference Reference2 = new InvoiceXMLEnvelopeHeaderReference();
                    Reference2.RefDate = System.String.Format("{0:yyyy-MM-dd}", line.ReferenceDate);
                    Reference2.RefNo = packingSlip;
                    Reference2.RefType = InvoiceXMLEnvelopeHeaderReferenceRefType.shipment ;
                    inv_xml.Envelope.Header.Reference[counter+1] = Reference2;

                }
                //fix itemPriceBruto
                ///change by Pesya 23/07/2014 from 2 digit to 4 digit
                System.Diagnostics.EventLog.WriteEntry("DMS", "" + "before itemPriceBruto: " + line.ItemPriceBruto.ToString(), System.Diagnostics.EventLogEntryType.Information);
                string itemPriceBrutoStr = line.ItemPriceBruto.ToString();
                if (itemPriceBrutoStr.IndexOf(".")>0)
                {
                    if (itemPriceBrutoStr.IndexOf(".") + 5 <= itemPriceBrutoStr.Length)
                    {
                        itemPriceBrutoStr = itemPriceBrutoStr.Substring(0, itemPriceBrutoStr.IndexOf(".") +5);

                    }
                }
                System.Diagnostics.EventLog.WriteEntry("DMS", "" + "after itemPriceBrutoStr:  " + itemPriceBrutoStr, System.Diagnostics.EventLogEntryType.Information);
                ////////////////////////////
                //fix LineSum
                string LineSumStr = line.LineSum.ToString();
                if (LineSumStr.IndexOf(".") > 0)
                {
                    if (LineSumStr.IndexOf(".") + 3 <= LineSumStr.Length)
                    {
                        LineSumStr = LineSumStr.Substring(0, LineSumStr.IndexOf(".") + 3);

                    }
                }
                    ////////////////////////////
                InvoiceXMLEnvelopeLine Detailes = new InvoiceXMLEnvelopeLine();
                Detailes.ItemBarcode = line.ItemBarcode;
                Detailes.CustomerBarcode = line.CustomerBarcode;
                Detailes.ItemDescription = line.ItemDescription;
                Detailes.ItemPriceBruto = itemPriceBrutoStr;
                Detailes.LineNo = line.LineNo;
                Detailes.LineSum = LineSumStr;
                Detailes.PartNumber = line.PartNumber;
                Detailes.UnitsQty = line.UnitQuantity.ToString();

                InvoiceXMLEnvelopeLineReference lRef = new InvoiceXMLEnvelopeLineReference();
                lRef.RefDate = System.String.Format("{0:yyyy-MM-dd}", line.ReferenceDate);
                lRef.RefNo = line.ReferenceNumber;
                lRef.RefType = InvoiceXMLEnvelopeLineReferenceRefType.purchaseOrder;
                
                Detailes.Reference = new InvoiceXMLEnvelopeLineReference[2];
                Detailes.Reference[0] = lRef; 
                //added by Liat 22.4 reference for shipment
                InvoiceXMLEnvelopeLineReference lRef1 = new InvoiceXMLEnvelopeLineReference();
                lRef1.RefDate = System.String.Format("{0:yyyy-MM-dd}", line.ReferenceDate);
                lRef1.RefType = InvoiceXMLEnvelopeLineReferenceRefType.shipment ;
                lRef1.RefNo = packingSlip;
                Detailes.Reference[1]=lRef1;

                inv_xml.Envelope.Details[counter] = Detailes;

                //I don't understand how per item enteties are going up to the parent level, but that's what I was told...
                inv_xml.Envelope.Header.PaymentTerms = line.PaymentTerms;
                inv_xml.Envelope.Header.Delivery = line.Delivery;
                counter++;
            }
            try
            {
                //using (StreamWriter output = new StreamWriter(new FileStream(path, FileMode.OpenOrCreate), Encoding.GetEncoding("windows-1255")))
                {
                    System.IO.FileStream fs = System.IO.File.Create(path);
                    XmlWriterSettings xmlWriterSettings = new XmlWriterSettings
                    {
                        Indent = true,
                        OmitXmlDeclaration = false,
                        Encoding = Encoding.GetEncoding(1255)
                    }; 
                    System.Xml.XmlWriter writer = System.Xml.XmlWriter.Create(fs, xmlWriterSettings);
                    System.Xml.Serialization.XmlSerializerNamespaces nss = new System.Xml.Serialization.XmlSerializerNamespaces();
                    nss.Add("", "");
                    System.Xml.Serialization.XmlSerializer x = new System.Xml.Serialization.XmlSerializer(inv_xml.GetType());
                    x.Serialize(writer, inv_xml, nss);
                    fs.Close();
                    if (debug)
                    {
                        x.Serialize(Console.Out, inv_xml);
                    }
                }
            }
            catch(Exception e)
            {
                System.Diagnostics.EventLog.WriteEntry("DMS","Exception in serialization method: " + e.ToString(), System.Diagnostics.EventLogEntryType.Error);
                if (debug)
                {
                   
                    Console.WriteLine("Exception in serialization method: " + e.Message);
                }
                return false;
            }
            return true;
        }
        
        
        void WriteToPRM(string name, string value, string sep = "\r\n")
        {
            value.Replace("{", "%7B");
            value.Replace("}", "%7D");
            byte[] bytes = System.Text.Encoding.Unicode.GetBytes(name + ":" + "{" + value.TrimEnd() + "}" + sep);
            st.Write(bytes,0, bytes.Length);
        }
        void WriteToPRM(string value)
        {
            byte[] bytes = System.Text.Encoding.Unicode.GetBytes(value);
            st.Write(bytes, 0, bytes.Length);
        }
        void WriteToPRM(InvoiceLine line)
        {
            WriteToPRM("Line:{");
            WriteToPRM("LineNo", line.LineNo, ",");
            WriteToPRM("Delivery", line.Delivery, ",");
            WriteToPRM("ReferenceType", line.ReferenceType, ",");
            WriteToPRM("ReferenceNumber", line.ReferenceNumber.ToString(), ",");
            WriteToPRM("ReferenceDate", line.ReferenceDate.ToShortDateString(), ",");
            WriteToPRM("PaymentTerms", line.PaymentTerms, ",");
            WriteToPRM("UnitQuantity", line.UnitQuantity.ToString(), ",");
            WriteToPRM("ItemPriceBruto", line.ItemPriceBruto.ToString(), ",");
            WriteToPRM("LineSum", line.LineSum.ToString(), ",");
            WriteToPRM("ItemDescription", line.ItemDescription, ",");
            WriteToPRM("PartNumber", line.PartNumber, ",");
            WriteToPRM("ItemBarcode", line.ItemBarcode, ",");
            WriteToPRM("CustomerBarcode", line.CustomerBarcode, ",");
            WriteToPRM("SalesOrder", line.SalesOrder, "}\r\n");
        }
        /// <summary>
        /// Fieldes and Properties 
        /// </summary>
        public string DocType
        {
            get { return "MMINVC"; }
        }
        public string Sender
        {
            get { return "7290058131372"; }
        }
        string flexReceiver;
        public string Receiver
        {
            get { return flexReceiver; }//"7290058174799";
            set { flexReceiver=value.Trim(); }
        }
        private DateTime flexMsgDate;

        public DateTime MsgDateTime
        {
            get { return flexMsgDate; }
            set { flexMsgDate = value; }
        }

        public string MsgDate
        {
            get { return MsgDateTime.Date.ToShortDateString(); }
        }
        public string MsgTime
        {
            get { return MsgDateTime.ToUniversalTime().ToShortTimeString(); }
        }
    

        private string flexInvoiceType;

        public string InvoiceType
        {
            get { return flexInvoiceType; }
            set { flexInvoiceType = value; }
        }

        private string flexInvoiceNumber;

        public string InvoiceNumber
        {
            get { return flexInvoiceNumber; }
            set { flexInvoiceNumber = value.TrimEnd(); }
        }

        private string flexSupplierName;

        public string SupplierName
        {
            get { return flexSupplierName; }
            set { flexSupplierName = value.TrimEnd(); }
        }

        private string flexDiscountAmount;

        public string DiscountAmount
        {
            get { return flexDiscountAmount; }
            set { flexDiscountAmount = value; }
        }

        private int flexCompanyCode;

        public int CompanyCode
        {
            get { return flexCompanyCode; }
            set { flexCompanyCode = value; }
        }

        private DateTime flexInvoiceDate;

        public DateTime InvoiceDate
        {
            get { return flexInvoiceDate; }
            set { flexInvoiceDate = value; }
        }

        private decimal flexExchangeRate;

        public decimal ExchangeRate
        {
            get { return flexExchangeRate; }
            set { flexExchangeRate = value; }
        }

        private string flexCurrentcy;

        public string Currentcy
        {
            get { return flexCurrentcy; }
            set { flexCurrentcy = value.TrimEnd(); }
        }

        /// <value>Tax Invoice</value>
        public string TaxInvoice
        {
            get
            {
                return "Tax Invoice";
            }
        }

        public string DocumentInvoiceType
        {
            get
            {
                return "Original";
            }
        }

        private string flexSupplierPrivateCompanyCode;

        public string SupplierPrivateCompanyCode
        {
            get { return flexSupplierPrivateCompanyCode; }
            set { flexSupplierPrivateCompanyCode = value.TrimEnd(); }
        }

        private string flexRetailerPrivateCompanyCode;

        public string RetailerPrivateCompanyCode
        {
            get { return flexRetailerPrivateCompanyCode; }
            set { flexRetailerPrivateCompanyCode = value.TrimEnd(); }
        }

        private decimal flexCurrencyDocumentSum;

        public decimal CurrencyDocumentSum
        {
            get { return flexCurrencyDocumentSum; }
            set { flexCurrencyDocumentSum = value; }
        }

        private decimal flexDocumentSum;

        public decimal DocumentSum
        {
            get { return flexDocumentSum; }
            set { flexDocumentSum = value; }
        }

        private decimal flexTaxSum;

        public decimal TaxSum
        {
            get { return flexTaxSum; }
            set { flexTaxSum = value; }
        }
        private decimal flexTaxSumNIS;

        public decimal TaxSumNIS
        {
            get { return flexTaxSumNIS; }
            set { flexTaxSumNIS = value; }
        }
        private decimal flexCurrencyRate;

        public decimal CurrencyRate
        {
            get { return flexCurrencyRate; }
            set { flexCurrencyRate = value; }
        }
        
        string flexTaxRate;
        public string TaxRate
        {
            get { return flexTaxRate; }//(CurrencyDocumentSum == (decimal)0.0 ? (decimal)0.0 : (flexTaxSum * (decimal)100.0) / CurrencyDocumentSum).ToString("F2");
            set { flexTaxRate = value.TrimEnd(); }
        }

        private string flexCompanyName;

        public string CompanyName
        {
            get { return flexCompanyName; }
            set { flexCompanyName = value.TrimEnd(); }
        }
        private string flexAddress;

        public string Address
        {
            get { return flexAddress; }
            set { flexAddress = value.TrimEnd(); }
        }
        private string flexCity;

        public string City
        {
            get { return flexCity; }
            set { flexCity = value.TrimEnd(); }
        }

        public string State
        {
            get { return ""; }
        }
        private string flexCountry;

        public string Country
        {
            get { return flexCountry; }
            set { flexCountry = value.Trim(); }
        }
        public string POB
        {
            get { return ""; }
        }
        private string flexZipcode;

        public string Zipcode
        {
            get { return flexZipcode; }
            set { flexZipcode = value.TrimEnd(); }
        }
        public string Bank 
        {
            get { return ""; }
        }
        public string Account 
        {
            get { return ""; }
        }
        public string Details 
        {
            get { return ""; }
        }

        private int flexTotalLines;

        public int TotalLines
        {
            get { return flexTotalLines; }
            set 
            {
                if (value < 0)
                {
                    throw new System.ArgumentOutOfRangeException();
                }
                flexTotalLines = value;
            }
        }

        private List<InvoiceLine> flexLines = new List<InvoiceLine>();
        private decimal flexLineSum;

        public decimal LineSum
        {
            get { return flexLineSum; }
            set { flexLineSum = value; }
        }

        public List<InvoiceLine> Lines
        {
            get { return flexLines; }
            set { flexLines = value; }
        }
        public void Add(InvoiceLine line)
        {
            flexLines.Add(line);
        }
        private void OpenConnection(OdbcConnection DbConnection)
        {
            try
            {
                if (DbConnection.State != System.Data.ConnectionState.Open)
                {
                    DbConnection.Close();
                    DbConnection.Open();
                }
            }
            catch (Exception )
            {
                //throw (new Exception("Could not connect to DB with " + doxParams["BAANDB"] + "\n" + ex.Message));
            }
        }
    }
}
