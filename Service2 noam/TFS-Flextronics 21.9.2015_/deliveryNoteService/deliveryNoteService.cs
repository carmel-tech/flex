using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.ServiceProcess;
using System.IO;
using System.Threading;
using BussinessLogicLayer;
using System.Xml;
using System.Data.Odbc;
using deliveryNoteService.Properties;
using System.Collections.Specialized;
using DataAccessLayer;
using System.Data;

namespace deliveryNoteService
{
    public partial class deliveryNoteService : ServiceBase
    {

        const string logName = "DMS Pdf Office";
        const string logSource = "DMS Pdf Office 2";
        private DoxHandler dox;
        Dictionary<string, string> doxParams;
        private string today = DateTime.Today.ToShortDateString();
        string connectionString = "Data Source=localhost;Initial Catalog=DoxPro_Env_Flex;Integrated Security=True";
        System.Timers.Timer invoice_timer = new System.Timers.Timer();
      
        System.Timers.Timer ShipmentGetBarcode_timer = new System.Timers.Timer();
        
        public int InvTotal
        {
            get
            {
                lock (this)
                {
                    return inv_good + inv_bad;
                }
            }
        }
        public decimal InvBadPercantage
        {
            get
            {
                lock (this)
                {
                    decimal total = (decimal)(inv_good + inv_bad);
                    return total != 0.0m ? ((decimal)(100.0m * inv_bad)) / total : 0.0m;
                }
            }
        }
        int inv_good = 0;

        public int InvGood
        {
            get
            {
                lock (this)
                {
                    return inv_good;
                }
            }
            set
            {
                lock (this)
                {
                    inv_good = value;
                }
            }
        }
        int inv_bad = 0;

        public int InvBad
        {
            get
            {
                lock (this)
                {
                    return inv_bad;
                }
            }
            set
            {
                lock (this)
                {
                    inv_bad = value;
                }
            }
        }

        public int PsTotal
        {
            get
            {
                lock (this)
                {
                    return ps_good_signed + ps_bad;
                }
            }
        }
        public decimal PsBadPercantage
        {
            get
            {
                lock (this)
                {
                    decimal total = (decimal)(ps_good_signed + ps_bad + ps_good_archive);
                    return total != 0.0m ? ((decimal)(100.0m * ps_bad)) / total : 0.0m;
                }
            }
        }
        int ps_good_archive = 0;

        public int PsGood_archive
        {
            get
            {
                lock (this)
                {
                    return ps_good_archive;
                }
            }
            set
            {
                lock (this)
                {
                    ps_good_archive = value;
                }
            }
        }


        int ps_good_signed = 0;

        public int PsGood_signed
        {
            get
            {
                lock (this)
                {
                    return ps_good_signed;
                }
            }
            set
            {
                lock (this)
                {
                    ps_good_signed = value;
                }
            }
        }


        int ps_bad = 0;

        public int PsBad
        {
            get
            {
                lock (this)
                {
                    return ps_bad;
                }
            }
            set
            {
                lock (this)
                {
                    ps_bad = value;
                }
            }
        }
        //-----------------------------
        public int PsTotalOfaqim
        {
            get
            {
                lock (this)
                {
                    return ps_good_1_signed + ps_bad_1;
                }
            }
        }
        public decimal PsBadPercantageOfaqim
        {
            get
            {
                lock (this)
                {
                    decimal total = (decimal)(ps_good_1_signed + ps_bad_1 + ps_good_1_archive);
                    return total != 0.0m ? ((decimal)(100.0m * ps_bad_1)) / total : 0.0m;
                }
            }
        }
        int ps_good_1_signed = 0;

        public int PsGoodOfaqim_signed
        {
            get
            {
                lock (this)
                {
                    return ps_good_1_signed;
                }
            }
            set
            {
                lock (this)
                {
                    ps_good_1_signed = value;
                }
            }
        }

        int ps_good_1_archive = 0;

        public int PsGoodOfaqim_archive
        {
            get
            {
                lock (this)
                {
                    return ps_good_1_archive;
                }
            }
            set
            {
                lock (this)
                {
                    ps_good_1_archive = value;
                }
            }
        }

        int ps_bad_1 = 0;

        public int PsBadOfaqim
        {
            get
            {
                lock (this)
                {
                    return ps_bad_1;
                }
            }
            set
            {
                lock (this)
                {
                    ps_bad_1 = value;
                }
            }
        }
        //-------------Mellanox ------------------
        public int PsTotalMellanox
        {
            get
            {
                lock (this)
                {
                    return ps_good_1_signed_mel + ps_bad_1_mel;
                }
            }
        }
        public decimal PsBadPercantageMellanox
        {
            get
            {
                lock (this)
                {
                    decimal total = (decimal)(ps_good_1_signed_mel + ps_bad_1_mel + ps_good_1_archive_mel);
                    return total != 0.0m ? ((decimal)(100.0m * ps_bad_1_mel)) / total : 0.0m;
                }
            }
        }
        int ps_good_1_signed_mel = 0;

        public int PsGoodMellanox_signed
        {
            get
            {
                lock (this)
                {
                    return ps_good_1_signed_mel;
                }
            }
            set
            {
                lock (this)
                {
                    ps_good_1_signed_mel = value;
                }
            }
        }

        int ps_good_1_archive_mel = 0;

        public int PsGoodMellanox_archive
        {
            get
            {
                lock (this)
                {
                    return ps_good_1_archive_mel;
                }
            }
            set
            {
                lock (this)
                {
                    ps_good_1_archive_mel = value;
                }
            }
        }

        int ps_bad_1_mel = 0;

        public int PsBadMellanox
        {
            get
            {
                lock (this)
                {
                    return ps_bad_1_mel;
                }
            }
            set
            {
                lock (this)
                {
                    ps_bad_1_mel = value;
                }
            }
        }
        //Ayala - Shipment
        //COC
        private int COC_Good = 0;
        public int COCGood
        {
            get
            {
                lock (this)
                {
                    return COC_Good;
                }
            }
            set
            {
                lock (this)
                {
                    COC_Good = value;
                }
            }
        }
        private int COC_Bad = 0;
        public int COCBad
        {
            get
            {
                lock (this)
                {
                    return COC_Bad;
                }
            }
            set
            {
                lock (this)
                {
                    COC_Bad = value;
                }
            }
        }

        public int COCTotal
        {
            get
            {
                lock (this)
                {
                    return COC_Bad + COC_Good;
                }
            }

        }
        public decimal COCBadPcnt
        {
            get
            {
                lock (this)
                {
                    decimal total = (decimal)(COC_Bad + COC_Good);
                    return total != 0.0m ? ((decimal)(100.0m * COC_Bad)) / total : 0.0m;
                }
            }
        }
        //Invoice
        private int Invoice_Good = 0;
        public int InvoiceGood
        {
            get
            {
                lock (this)
                {
                    return Invoice_Good;
                }
            }
            set
            {
                lock (this)
                {
                    Invoice_Good = value;
                }
            }
        }
        private int Invoice_Bad = 0;
        public int InvoiceBad
        {
            get
            {
                lock (this)
                {
                    return Invoice_Bad;
                }
            }
            set
            {
                lock (this)
                {
                    Invoice_Bad = value;
                }
            }
        }

        public int InvoiceTotal
        {
            get
            {
                lock (this)
                {
                    return Invoice_Bad + Invoice_Good;
                }
            }

        }
        public decimal InvoiceBadPcnt
        {
            get
            {
                lock (this)
                {
                    decimal total = (decimal)(Invoice_Bad + Invoice_Good);
                    return total != 0.0m ? ((decimal)(100.0m * Invoice_Bad)) / total : 0.0m;
                }
            }
        }
        //Delivery Note
        private int deliveryNoteGetBarcode = 0;
        public int DeliveryNoteGetBarcode
        {
            get
            {
                lock (this)
                {
                    return deliveryNoteGetBarcode;
                }
            }
            set
            {
                lock (this)
                {
                    deliveryNoteGetBarcode = value;
                }
            }
        }
        private int cOCGetBarcode = 0;
        public int COCGetBarcode
        {
            get
            {
                lock (this)
                {
                    return cOCGetBarcode;
                }
            }
            set
            {
                lock (this)
                {
                    cOCGetBarcode = value;
                }
            }
        }
        private int invoiceGetBarcode = 0;
        public int InvoiceGetBarcode
        {
            get
            {
                lock (this)
                {
                    return invoiceGetBarcode;
                }
            }
            set
            {
                lock (this)
                {
                    invoiceGetBarcode = value;
                }
            }
        }
        private int cOCGetData = 0;
        public int COCGetData
        {
            get
            {
                lock (this)
                {
                    return cOCGetData;
                }
            }
            set
            {
                lock (this)
                {
                    cOCGetData = value;
                }
            }
        }

        private int invoiceGetData = 0;
        public int InvoiceGetData
        {
            get
            {
                lock (this)
                {
                    return invoiceGetData;
                }
            }
            set
            {
                lock (this)
                {
                    invoiceGetData = value;
                }
            }
        }
        private int deliveryNoteGetData = 0;
        public int DeliveryNoteGetData
        {
            get
            {
                lock (this)
                {
                    return deliveryNoteGetData;
                }
            }
            set
            {
                lock (this)
                {
                    deliveryNoteGetData = value;
                }
            }
        }
        private int DeliveryNote_Good = 0;
        public int DeliveryNoteGood
        {
            get
            {
                lock (this)
                {
                    return DeliveryNote_Good;
                }
            }
            set
            {
                lock (this)
                {
                    DeliveryNote_Good = value;
                }
            }
        }
        private int DeliveryNote_Bad = 0;
        public int DeliveryNoteBad
        {
            get
            {
                lock (this)
                {
                    return DeliveryNote_Bad;
                }
            }
            set
            {
                lock (this)
                {
                    DeliveryNote_Bad = value;
                }
            }
        }

        public int DeliveryNoteTotal
        {
            get
            {
                lock (this)
                {
                    return DeliveryNote_Bad + DeliveryNote_Good;
                }
            }

        }
        public decimal DeliveryNoteBadPcnt
        {
            get
            {
                lock (this)
                {
                    decimal total = (decimal)(DeliveryNote_Bad + DeliveryNote_Good);
                    return total != 0.0m ? ((decimal)(100.0m * DeliveryNote_Bad)) / total : 0.0m;
                }
            }
        }
        //---------------------------------
        public deliveryNoteService()
        {
            System.Diagnostics.Debugger.Launch();
            InitializeComponent();
            if (!System.Diagnostics.EventLog.SourceExists(logSource))
            {
                System.Diagnostics.EventLog.CreateEventSource(logSource, logName);
            }
            eventLogger.Source = logSource;
            eventLogger.Log = logName;
        }
        public void OnStart()
        {
            eventLogger.WriteEntry("Delivery Notes Service started");
            doxParams = new Dictionary<string, string>();
            try
            {

                doxParams.Add("BAANLog", Properties.Settings.Default.BAANLog);
                doxParams.Add("BAANOrd99Log", Properties.Settings.Default.BAANOrd99Log);
                doxParams.Add("BAANDB", Properties.Settings.Default.BAANDB);
                doxParams.Add("ItemsQ", Properties.Settings.Default.ItemsQuery);
                doxParams.Add("CustQ", Properties.Settings.Default.CustomerOrderQuery);
                doxParams.Add("SuppQ", Properties.Settings.Default.SuppliersQuery);
                doxParams.Add("SuppEmailQ", Properties.Settings.Default.SuppliersEmailsQuery);
                doxParams.Add("BarcodeFile", Properties.Settings.Default.BarcodeDebugFile);
                doxParams.Add("DefectedQueuePath", Properties.Settings.Default.DefectedQueuePath);
                doxParams.Add("CustPeopleQ", Properties.Settings.Default.CustomerPeopleQuery);
                doxParams.Add("InvoiceLogPath", Properties.Settings.Default.InvoiceLogPath);
                doxParams.Add("NewOrder99Path", Properties.Settings.Default.NewOrder99Path);
                doxParams.Add("NewOrder99ReturnedDocumentPath", Properties.Settings.Default.NewOrder99ReturnedDocumentPath);
                doxParams.Add("Ord99EmptyPDFDocument", Properties.Settings.Default.Ord99EmptyPDFDocument);
                doxParams.Add("NewInvoicePath", Properties.Settings.Default.NewInvoicePath);
                doxParams.Add("StoreNextPath", Properties.Settings.Default.StoreNextPath);
                doxParams.Add("DoxEnv", Properties.Settings.Default.DoxEnv);
                doxParams.Add("StorenextXMLLandingFolder", Properties.Settings.Default.StorenextXMLLandingFolder);
                doxParams.Add("StorenextPDFLandingFolder", Properties.Settings.Default.StorenextPDFLandingFolder);
                doxParams.Add("EnableStorenextIntegration", Properties.Settings.Default.EnableStorenextIntegration.ToString());
                doxParams.Add("StorenextRanFile", Properties.Settings.Default.StorenextRanFile.ToString());
                doxParams.Add("IntegrationQueue", Properties.Settings.Default.IntegrationQueue);
                doxParams.Add("", Properties.Settings.Default.IntegrationWakeUpTime.ToShortTimeString());
                doxParams.Add("FlexStorenextSuppQuery", Properties.Settings.Default.GetFlexStorenextSuppliersQuery);
                doxParams.Add("SuppInvErrQ", Properties.Settings.Default.SupplierInvoiceErrorQueue);
                doxParams.Add("StorenextSuppInvDocPath", Properties.Settings.Default.StorenextSuppInvDocPath);
                doxParams.Add("MachsanErrors", Properties.Settings.Default.MachsanErrors);



                RestoreStatistics();
            }
            catch (Exception ex)
            {
                eventLogger.WriteEntry("Exception on startup of Dox handler: \"" + ex.Message + "\"");
            }

            try
            {
                dox = new DoxHandler(doxParams, eventLogger);
               
            }
            catch (Exception ex)
            {
                eventLogger.WriteEntry(ex.Message, EventLogEntryType.Error);
            }
            if (dox != null)
            {
                eventLogger.WriteEntry("DOX handler started");
            }
            else
            {
                eventLogger.WriteEntry("DOX handler failed to start", EventLogEntryType.Error);
            }

           
            ShipmentGetBarcode_timer.Interval = 15000;
            ShipmentGetBarcode_timer.AutoReset = false;
            ShipmentGetBarcode_timer.Elapsed += new System.Timers.ElapsedEventHandler(ShipmentGetBarcode_timer_Elapsed);
            ShipmentGetBarcode_timer.Start();

            

        }

        protected override void OnStart(string[] args)
        {
            eventLogger.WriteEntry("Delivery Notes Service started");
            doxParams = new Dictionary<string, string>();
            try
            {

                doxParams.Add("BAANLog", Properties.Settings.Default.BAANLog);
                doxParams.Add("BAANOrd99Log", Properties.Settings.Default.BAANOrd99Log);
                doxParams.Add("BAANDB", Properties.Settings.Default.BAANDB);
                doxParams.Add("ItemsQ", Properties.Settings.Default.ItemsQuery);
                doxParams.Add("CustQ", Properties.Settings.Default.CustomerOrderQuery);
                doxParams.Add("SuppQ", Properties.Settings.Default.SuppliersQuery);
                doxParams.Add("SuppEmailQ", Properties.Settings.Default.SuppliersEmailsQuery);
                doxParams.Add("BarcodeFile", Properties.Settings.Default.BarcodeDebugFile);
                doxParams.Add("DefectedQueuePath", Properties.Settings.Default.DefectedQueuePath);
                doxParams.Add("CustPeopleQ", Properties.Settings.Default.CustomerPeopleQuery);
                doxParams.Add("InvoiceLogPath", Properties.Settings.Default.InvoiceLogPath);
                doxParams.Add("NewOrder99Path", Properties.Settings.Default.NewOrder99Path);
                doxParams.Add("NewOrder99ReturnedDocumentPath", Properties.Settings.Default.NewOrder99ReturnedDocumentPath);
                doxParams.Add("Ord99EmptyPDFDocument", Properties.Settings.Default.Ord99EmptyPDFDocument);
                doxParams.Add("NewInvoicePath", Properties.Settings.Default.NewInvoicePath);
                doxParams.Add("StoreNextPath", Properties.Settings.Default.StoreNextPath);
                doxParams.Add("DoxEnv", Properties.Settings.Default.DoxEnv);
                doxParams.Add("StorenextXMLLandingFolder", Properties.Settings.Default.StorenextXMLLandingFolder);
                doxParams.Add("StorenextPDFLandingFolder", Properties.Settings.Default.StorenextPDFLandingFolder);
                doxParams.Add("EnableStorenextIntegration", Properties.Settings.Default.EnableStorenextIntegration.ToString());
                doxParams.Add("StorenextRanFile", Properties.Settings.Default.StorenextRanFile.ToString());
                doxParams.Add("IntegrationQueue", Properties.Settings.Default.IntegrationQueue);
                doxParams.Add("", Properties.Settings.Default.IntegrationWakeUpTime.ToShortTimeString());
                doxParams.Add("FlexStorenextSuppQuery", Properties.Settings.Default.GetFlexStorenextSuppliersQuery);
                doxParams.Add("SuppInvErrQ", Properties.Settings.Default.SupplierInvoiceErrorQueue);
                doxParams.Add("StorenextSuppInvDocPath", Properties.Settings.Default.StorenextSuppInvDocPath);
                doxParams.Add("MachsanErrors", Properties.Settings.Default.MachsanErrors);



                RestoreStatistics();
            }
            catch (Exception ex)
            {
                eventLogger.WriteEntry("Exception on startup of Dox handler: \"" + ex.Message + "\"");
            }

            try
            {
                dox = new DoxHandler(doxParams, eventLogger);
            }
            catch (Exception ex)
            {
                eventLogger.WriteEntry(ex.Message, EventLogEntryType.Error);
            }
            if (dox != null)
            {
                eventLogger.WriteEntry("DOX handler started");
            }
            else
            {
                eventLogger.WriteEntry("DOX handler failed to start", EventLogEntryType.Error);
            }

          
            //Ayala 13.04.2015
            ShipmentGetBarcode_timer.Interval = 15000;
            ShipmentGetBarcode_timer.AutoReset = false;
          
            ShipmentGetBarcode_timer.Elapsed += new System.Timers.ElapsedEventHandler(ShipmentGetBarcode_timer_Elapsed);
          
            ShipmentGetBarcode_timer.Start();
          
           
        }
        
        // This Method will be called when the service is to be stopped
        // It stpos the timer, waits until current record processing is completed, and release the FineReader engine
        protected override void OnStop()
        {
            eventLogger.WriteEntry("DeliveryService has been stopped");
            if (invoice_timer != null)
            {
                invoice_timer.Enabled = false;
                // Stop the timer completly
                invoice_timer.Stop();
                invoice_timer = null;
            }
            //if (ps_timer != null)
            //{
            //    ps_timer.Enabled = false;
            //    // Stop the timer completly
            //    ps_timer.Stop();
            //    ps_timer = null;
            //}
        }

        //stay
        //Ayala 13.04.2015
        //The method search for Shipment  - pdf file and split it by the Stickers in order to archive it
        private void ShipmentGetBarcode_timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            
            ShipmentGetBarcode_timer.Enabled = false;
            eventLogger.WriteEntry(" Shipment_timer_Elapsed");
          
            System.IO.DirectoryInfo spTray = new System.IO.DirectoryInfo("\\\\mignt002\\storenext\\DOXTemp\\DOX");//(Properties.Settings.Default.ShipmentsDir);//Properties.Settings.Default.ShipmentsDir
            //  System.IO.DirectoryInfo[] spDirectories = spTray.GetDirectories();
            eventLogger.WriteEntry(" get directory " + spTray.FullName);
            //foreach (DirectoryInfo dir in spDirectories)
            //{
            System.IO.DirectoryInfo[] Directories_3 = spTray.GetDirectories();
           
            foreach (DirectoryInfo item in Directories_3)
            {
                
                if (item.Name.Equals("Delivery note") || item.Name.Equals("COC") || item.Name.Equals("Invoice"))
                // if (item.Name.Equals("1") || item.Name.Equals("2") || item.Name.Equals("3"))
                {
                    System.IO.FileInfo[] spFilesInfo = item.GetFiles("*.pdf");

                  
                    
                    foreach (FileInfo mrbFile in spFilesInfo)
                    {
                        //try
                        //          {
                        //       using (IDataSupplier dataSupplier = DataManager.GetDataSupplier(DataManager.defaultType, connectionString))
                        //    {
                        //        dataSupplier.OpenQuery();
                        //        dataSupplier.AddParameter("FileName", FileName);
                        //        dataSupplier.AddParameter("KindOfDoc", KindOfDoc);
                        //        dataSupplier.AddParameter("IsSuccess", IsSuccess);
                        //        dataSupplier.AddParameter("ReasonId", ReasonId);
                        //        dataSupplier.Execute("INSERT INTO [LogArchiveShipment]([FileName],[KindOfDoc] ,[IsSuccess],[ReasonId],[date]) \r\n\t\t\t\t     VALUES ( @FileName,@KindOfDoc,@IsSuccess ,@ReasonId, getdate())");
                        //    }
                        //}
                        //catch (Exception ex)
                        //{
                        //    eventLogger.WriteEntry("Error in write to LogArchiveShipment - " + ex.Message);

                        //}
                        //   string typeItem=item.Name == "1" ? "COC" : item.Name == "2"?"Delivery Note":"Invoice";
                        //if (IsFileLocked(mrbFile))
                        //{
                        //    File.AppendAllText("c:\\batya\\damaged.txt", "damaged = " + mrbFile.FullName +Environment.NewLine);
                        //    if (IsFileLocked(mrbFile))
                                
                                
                        //}
                      
                        eventLogger.WriteEntry(" get file " + mrbFile.FullName);
                       
                        string response = dox.handleShipmentDoc(mrbFile.FullName, Path.GetFileNameWithoutExtension(mrbFile.FullName), item.Name);
                       eventLogger.WriteEntry(response + " - responseformachsan");
                        int good = int.Parse(response.Substring(0, response.IndexOf("_")));


                     
                       if (item.Name.Equals("COC"))

                            COCGetBarcode += good;

                        else if (item.Name.Equals("Invoice"))
                            InvoiceGetBarcode += good;
                        else
                            DeliveryNoteGetBarcode += good;


                    }


                }






                else if (item.Name.Equals("CDI-FromQueue"))
                {
                    try
                    {
                        System.IO.FileInfo[] spFilesInfo = item.GetFiles("*.pdf");
                        foreach (FileInfo shipFile in spFilesInfo)
                        {

                            using (IDataSupplier dataSupplier = DataManager.GetDataSupplier(DataManager.defaultType, connectionString))
                            {
                                dataSupplier.OpenQuery();
                                dataSupplier.AddParameter("FileName", shipFile.FullName);
                                dataSupplier.AddParameter("lotsAndMakats", shipFile.Name.Replace("DeliveryNote_", "").Replace("COC_", "").Replace("Invoice_", "")
                                    .Replace("DeliveryNote", "").Replace("COC", "").Replace("Invoice", "").Replace("_", ",").Replace(".pdf", ""));
                                dataSupplier.AddParameter("type", shipFile.Name.ToLower().StartsWith("COC".ToLower()) ? "COC" : shipFile.Name.ToLower().StartsWith("Invoice".ToLower()) ? "Invoice" : "Delivery Note");
                                dataSupplier.AddParameter("status", "0");
                                dataSupplier.AddParameter("dategetfile", DateTime.Now.ToString("MM-dd-yyyy HH:mm:ss"));
                                dataSupplier.Execute("if @FileName not in (select FileName from shipmentArchive where status='1' or status='0')" +
                                    " insert into shipmentArchive (FileName,type,lotsAndMakats,status,dategetfile) values (@FileName,@type,@lotsAndMakats,@status,@dategetfile)");
                            }


                        }
                    }
                    catch (Exception ex)
                    {
                        eventLogger.WriteEntry("Error in  update to shipmentArchive- " + ex.Message);


                    }



                }
            } ShipmentGetBarcode_timer.Enabled = true;
        }
      
        private void openConnection(OdbcConnection DbConnection)
        {
            try
            {
                if (DbConnection.State != System.Data.ConnectionState.Open)
                {
                    //                    DbConnection.Close();
                    DbConnection.Open();
                }
            }
            catch (Exception ex)
            {
                throw (new Exception("Could not connect to DB with " + doxParams["BAANDB"] + "\n" + ex.Message));
            }
        }
        protected virtual bool IsFileLocked(FileInfo file)
        {
            FileStream stream = null;

            try
            {
                stream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            }
            catch (IOException ex)
            {

                eventLogger.WriteEntry("Error IOException " + ex.Message, EventLogEntryType.Information);
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }

            //file is not locked
            return false;
        }
       
        void handleStorenextSuppInvDoc(string fullName)
        {
            try
            {
                bool retry = true;
                int retries = 0;
                while (retry)
                {
                    try
                    {
                        string[] entityDesc = System.IO.File.ReadAllLines(fullName);
                        retry = false;
                    }
                    catch (Exception ex)
                    {
                        ++retries;
                        if (ex.Message.IndexOf("used by another process") > 0 && retries < 3)
                        {
                            eventLogger.WriteEntry("Cannot access " + fullName + " retry.\n" + ex.Message);
                            Thread.Sleep(2000);
                        }
                        else
                        {
                            eventLogger.WriteEntry("Cannot access " + fullName + "\n" + ex.Message);
                            retry = false;
                        }

                    }
                }
                string fileName = Path.GetFileName(fullName);
                string[] FileProp = fileName.Split('_');
                string PdfFileName = Path.GetDirectoryName(fullName) + @"\" + FileProp[2] + "_" + FileProp[3] + "_" + Path.GetFileNameWithoutExtension(FileProp[4]) + ".pdf";
                string response = dox.handleStorenextInvoiceDoc(fullName, PdfFileName, FileProp[0].Substring(3));
                if (response != "")
                {
                    eventLogger.WriteEntry(response, EventLogEntryType.Error);

                }

            }
            catch (Exception ex)
            {
                eventLogger.WriteEntry(ex.Message, EventLogEntryType.Error);

            }

        }
        private void ResetStatistics()
        {
            string now = DateTime.Today.ToShortDateString();
            if (now != today)
            {
                today = now;
                InvGood = 0;
                InvBad = 0;
                PsBad = 0;
                PsGood_archive = 0;
                PsGood_signed = 0;
                PsBadOfaqim = 0;
                PsGoodOfaqim_archive = 0;
                PsGoodOfaqim_signed = 0;
                COCBad = 0;
                COCGood = 0;
                InvoiceBad = 0;
                InvoiceGood = 0;
                DeliveryNoteBad = 0;
                DeliveryNoteGood = 0;
                PsBadMellanox = 0;
                PsGoodMellanox_archive = 0;
                PsGoodMellanox_signed = 0;
                deliveryNoteGetData = 0;
                deliveryNoteGetData = 0;
                COCGetBarcode = 0;
                COCGetData = 0;
                InvoiceGetBarcode = 0;
                InvoiceGetData = 0;

            }
        }
        protected void SaveStatistics()
        {
            lock (this)
            {
                string folder = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                string path = Path.Combine(folder, "statistics.xml");
                XmlTextWriter writer = new XmlTextWriter(path, null);
                //Write the root element
                writer.WriteStartElement("Statitics");
                //Write sub-elements
                writer.WriteElementString("Date", today);
                writer.WriteElementString("PsGood_signed", PsGood_signed.ToString());
                writer.WriteElementString("PsGood_archive", PsGood_archive.ToString());
                writer.WriteElementString("PsBad", PsBad.ToString());
                writer.WriteElementString("InvGood", InvGood.ToString());
                writer.WriteElementString("InvBad", InvBad.ToString());
                writer.WriteElementString("PsGoodOfaqim_signed", PsGoodOfaqim_signed.ToString());
                writer.WriteElementString("PsGoodOfaqim_archive", PsGoodOfaqim_archive.ToString());
                writer.WriteElementString("PsBadOfaqim", PsBadOfaqim.ToString());
                writer.WriteElementString("COCBad", COCBad.ToString());
                writer.WriteElementString("COCGood", COCGood.ToString());
                writer.WriteElementString("DeliveryNoteGetData", DeliveryNoteGetData.ToString());
                writer.WriteElementString("DeliveryNoteGetBarcode", DeliveryNoteGetBarcode.ToString());
                writer.WriteElementString("COCGetBarcode", COCGetBarcode.ToString());
                writer.WriteElementString("COCGetData", COCGetData.ToString());
                writer.WriteElementString("InvoiceGetBarcode", InvoiceGetBarcode.ToString());
                writer.WriteElementString("InvoiceGetData", InvoiceGetData.ToString());
                writer.WriteElementString("InvoiceBad", InvoiceBad.ToString());
                writer.WriteElementString("InvoiceGood", InvoiceGood.ToString());
                writer.WriteElementString("DeliveryNoteBad", DeliveryNoteBad.ToString());
                writer.WriteElementString("DeliveryNoteGood", DeliveryNoteGood.ToString());
                writer.WriteElementString("PsGoodMellanox_signed", PsGoodMellanox_signed.ToString());
                writer.WriteElementString("PsGoodMellanox_archive", PsGoodMellanox_archive.ToString());
                writer.WriteElementString("PsBadMellanox", PsBadMellanox.ToString());

                // end the root element
                writer.WriteEndElement();

                //Write the XML to file and close the writer
                writer.Close();
            }
        }
        protected void RestoreStatistics()
        {
            string folder = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string path = Path.Combine(folder, "statistics.xml");
            using (XmlReader reader = XmlReader.Create(path))
            {

                // Parse the XML document.  ReadString is used to 
                // read the text content of the elements.
                reader.Read();
                reader.ReadStartElement("Statitics");
                reader.ReadStartElement("Date");
                today = reader.ReadString();
                reader.ReadEndElement();
                reader.ReadStartElement("PsGood_signed");
                PsGood_signed = reader.ReadContentAsInt();
                reader.ReadEndElement();
                reader.ReadStartElement("PsGood_archive");
                PsGood_archive = reader.ReadContentAsInt();
                reader.ReadEndElement();

                reader.ReadStartElement("PsBad");
                PsBad = reader.ReadContentAsInt();
                reader.ReadEndElement();

                reader.ReadStartElement("InvGood");
                InvGood = reader.ReadContentAsInt();
                reader.ReadEndElement();

                reader.ReadStartElement("InvBad");
                InvBad = reader.ReadContentAsInt();
                reader.ReadEndElement();




                reader.ReadStartElement("PsGoodOfaqim_signed");
                PsGoodOfaqim_signed = reader.ReadContentAsInt();
                reader.ReadEndElement();

                reader.ReadStartElement("PsGoodOfaqim_archive");
                PsGoodOfaqim_archive = reader.ReadContentAsInt();
                reader.ReadEndElement();
                reader.ReadStartElement("PsBadOfaqim");
                PsBadOfaqim = reader.ReadContentAsInt();
                reader.ReadEndElement();

                reader.ReadStartElement("COCBad");
                COCBad = reader.ReadContentAsInt();
                reader.ReadEndElement();
                reader.ReadStartElement("COCGood");
                COCGood = reader.ReadContentAsInt();
                reader.ReadEndElement();




                reader.ReadStartElement("DeliveryNoteGetData");
                DeliveryNoteGetData = reader.ReadContentAsInt();
                reader.ReadEndElement();
                reader.ReadStartElement("DeliveryNoteGetBarcode");
                DeliveryNoteGetBarcode = reader.ReadContentAsInt();
                reader.ReadEndElement();
                reader.ReadStartElement("COCGetBarcode");
                COCGetBarcode = reader.ReadContentAsInt();
                reader.ReadEndElement();
                reader.ReadStartElement("COCGetData");
                COCGetData = reader.ReadContentAsInt();
                reader.ReadEndElement();
                reader.ReadStartElement("InvoiceGetBarcode");
                InvoiceGetBarcode = reader.ReadContentAsInt();
                reader.ReadEndElement();
                reader.ReadStartElement("InvoiceGetData");
                InvoiceGetData = reader.ReadContentAsInt();
                reader.ReadEndElement();

                reader.ReadStartElement("InvoiceBad");
                InvoiceBad = reader.ReadContentAsInt();
                reader.ReadEndElement();
                reader.ReadStartElement("InvoiceGood");
                InvoiceGood = reader.ReadContentAsInt();
                reader.ReadEndElement();

                reader.ReadStartElement("DeliveryNoteBad");
                DeliveryNoteBad = reader.ReadContentAsInt();
                reader.ReadEndElement();
                reader.ReadStartElement("DeliveryNoteGood");
                DeliveryNoteGood = reader.ReadContentAsInt();
                reader.ReadEndElement();















                reader.ReadStartElement("PsGoodMellanox_signed");
                PsGoodMellanox_signed = reader.ReadContentAsInt();
                reader.ReadEndElement();

                reader.ReadStartElement("PsGoodMellanox_archive");
                PsGoodMellanox_archive = reader.ReadContentAsInt();
                reader.ReadEndElement();
                reader.ReadStartElement("PsBadMellanox");
                PsBadMellanox = reader.ReadContentAsInt();
                reader.ReadEndElement();
                reader.ReadEndElement();

            }
        }
        ////---------------------------------------------------
        private void handleAndArchivePackingSlip(string fullName)
        {
            bool migdal = false;
            bool Mellanox = false;
            try
            {
                if (fullName.IndexOf("_err") > 0)
                {
                    FileInfo psFile = new FileInfo(fullName);
                    psFile.MoveTo(fullName.Replace("_err", ""));
                    fullName = psFile.FullName;
                }
                FileInfo fi = new FileInfo(fullName);
                string name = fi.Name;
                int n = 0;

                if (fullName.Contains("mellanox_"))
                    Mellanox = true;
                else if (!fullName.Contains("ofk_"))
                    migdal = true;

                //if ((name.Length >= 4 && int.TryParse(name.Substring(0, 4), out n))
                //    ||
                //    (name.Length >= 5 && name.StartsWith("BC"))
                //  )
                //{
                //    migdal = true;
                //}
                eventLogger.WriteEntry("handlePackingSlip: try to Archive " + fullName, EventLogEntryType.Information);
                string response = dox.handleArchivePackingSlip(fullName);
                // eventLogger.WriteEntry("response: " + response == "" ? " Success to Archive " :( " fail to Archive the response : " +response), EventLogEntryType.Information);

                if (response == "MRB")
                {
                    eventLogger.WriteEntry("handlePackingSlip: " + response, EventLogEntryType.Warning);
                    if (migdal)
                    {
                        PsGood_archive++;
                        PsGood_signed++;
                    }
                    else if (Mellanox)
                    {
                        PsGoodMellanox_archive++;
                        PsGoodMellanox_signed++;

                    }
                    else
                    {
                        PsGoodOfaqim_archive++;
                        PsGoodOfaqim_signed++;
                    }
                }
                else if (response.Contains("MRB"))
                {
                    if (migdal)
                    {
                        PsBad++;
                    }
                    else if (Mellanox)
                    {
                        PsBadMellanox++;
                    }
                    else
                    {
                        PsBadOfaqim++;
                    }
                }
                else if (response != "")
                {
                    eventLogger.WriteEntry("handlePackingSlip: " + response, EventLogEntryType.Warning);
                    if (migdal)
                    {
                        PsBad++;
                    }
                    else if (Mellanox)
                    {
                        PsBadMellanox++;
                    }
                    else
                    {
                        PsBadOfaqim++;
                    }
                }
                else
                {
                    if (migdal)
                    {
                        PsGood_archive++;
                    }
                    else if (Mellanox)
                    {
                        PsGoodMellanox_archive++;
                    }
                    else
                        PsGoodOfaqim_archive++;
                }
            }
            catch (Exception ex)
            {
                eventLogger.WriteEntry("handlePackingSlip: Exception while handling packing slip " + ex.Message + "  " + ex.ToString(), EventLogEntryType.Error);
            }
        }
        //-----------------------------------------------
        void handlePackingSlip(string fullName, string fileName, string pathDirectory, string type)
        {



            string fullname2 = "";
            string dateti = "_" + DateTime.Now.ToString().Replace("/", "_").Replace(":", "_").Replace(" AM", "").Replace(" PM", "").Replace(" ", "_");
            int n = 0;
            bool migdal = false, Mellanox = false;
            if (type != "")
                Mellanox = true;
            else if (fullName.Contains("mellanox_"))
            {
                Mellanox = true;
            }
            else
                if ((fileName.Length >= 4 && int.TryParse(fileName.Substring(0, 4), out n))
                            ||
                            (fileName.Length >= 5 && fileName.StartsWith("BC"))
                  )
                {
                    migdal = true;
                }

            eventLogger.WriteEntry("Mellanox - " + Mellanox + " , migdal - " + migdal);
            string PackingSlipBarcode = "Error";
            Thread.Sleep(800);
            //FIRST LEVEL GET BARCODE
            if (Path.GetFileNameWithoutExtension(fullName).StartsWith("BC"))
            {
                if (Path.GetFileNameWithoutExtension(fullName.ToLower()).Contains("M".ToLower()))
                    PackingSlipBarcode = "400M  " + Path.GetFileNameWithoutExtension(fullName).Substring(2).Replace("M", "").Replace("m", "");
                else
                    PackingSlipBarcode = "400  " + Path.GetFileNameWithoutExtension(fullName).Substring(2);
                if (PackingSlipBarcode.Contains("_"))
                    PackingSlipBarcode = PackingSlipBarcode.Substring(0, PackingSlipBarcode.IndexOf("_"));
                eventLogger.WriteEntry("barcode " + PackingSlipBarcode, System.Diagnostics.EventLogEntryType.Information);


            }
            else if (Path.GetFileNameWithoutExtension(fullName).StartsWith("mellanox_"))
            {

                if (!Path.GetFileNameWithoutExtension(fullName).Contains("mellanox_400"))
                    PackingSlipBarcode = "400  " + Path.GetFileNameWithoutExtension(fullName).Replace("mellanox_", "");
                else
                    PackingSlipBarcode = Path.GetFileNameWithoutExtension(fullName).Replace("mellanox_", "");
                if (PackingSlipBarcode.Contains("_"))
                    PackingSlipBarcode = PackingSlipBarcode.Substring(0, PackingSlipBarcode.IndexOf("_"));
                eventLogger.WriteEntry("barcode " + PackingSlipBarcode, System.Diagnostics.EventLogEntryType.Information);

            }
            else
                PackingSlipBarcode = dox.GetBarcodesForPS(fullName);

            eventLogger.WriteEntry("PackingSlipBarcode - " + PackingSlipBarcode);
            if (Mellanox && !(fullName.Contains("mellanox_")) && !PackingSlipBarcode.ToLower().Contains("POINT".ToLower()))
            {
                try
                {
                    fullname2 = Path.GetDirectoryName(fullName) + "\\mellanox_" + fileName;
                    File.Move(fullName, fullname2);
                    eventLogger.WriteEntry("move from - " + fullName + " - to - " + fullname2);
                    fullName = fullname2;
                    eventLogger.WriteEntry("new fullName - " + fullName);
                }
                catch (Exception ex)
                {
                    eventLogger.WriteEntry("replace to _mellanox - " + ex.Message);
                }
            }


            else if (PackingSlipBarcode.ToLower().Contains("POINT".ToLower()))
            {
                if (fullName.Contains("_bad"))
                {

                    int num = 0;
                    string numm = "0";
                    try
                    {

                        numm = fileName.Remove(0, fileName.Length - 5);
                        numm = numm.Substring(0, 1);
                    }
                    catch (Exception ex)
                    {
                        eventLogger.WriteEntry(ex.Message);
                    }
                    if (numm == "0")
                    {
                        eventLogger.WriteEntry("returned barcode is invalid: " + PackingSlipBarcode + " move to queue movetoDefQ numm == 0");
                        // If failed to read barcode, the PDF is moved to a location of a DOX-Pro queue
                        dox.movetoDefQ(fullName, 1, "Packing Slip", "bad pointer  -  0 , fail 10 times");
                        return;
                    }
                    else
                    {

                        try
                        {
                            num = int.Parse(numm);
                            num--;
                            System.IO.File.Move(fullName, pathDirectory + "\\" + fileName.Substring(0, fileName.LastIndexOf(".") - 1) + num + ".pdf");
                        }
                        catch (Exception ex)
                        {

                            eventLogger.WriteEntry("fail_bad_pointer " + ex.Message);
                        }
                        return;
                    }

                }
                else
                    System.IO.File.Move(fullName, pathDirectory + "\\" + fileName.Substring(0, fileName.LastIndexOf(".")) + "_bad9.pdf");



                return;

            }
            else if (PackingSlipBarcode.IndexOf("Error") > -1 || PackingSlipBarcode.Length < 15)
            {
                eventLogger.WriteEntry("returned barcode is invalid: " + PackingSlipBarcode + " move to queue movetoDefQ  " + fullName);
                // If failed to read barcode, the PDF is moved to a location of a DOX-Pro queue
                if (File.Exists(fullName))
                {
                    dox.movetoDefQ(fullName, 1, "Packing Slip", "barcode.IndexOf(Error)>-1 || barcode.Length < 15");
                    if (migdal)
                    {
                        PsBad++;
                    }
                    else if (Mellanox)
                    {
                        PsBadMellanox++;
                    }
                    else
                        PsBadOfaqim++;
                }
                return;
            }

            string CheckedPath = fullName;
            if (File.Exists(CheckedPath))
            {
                try
                {
                    if (fileName.Contains("_bad"))
                    {
                        CheckedPath = Properties.Settings.Default.CheckedPackingSlipMIGWithoutBPath + "\\" + fileName.Substring(0, fileName.LastIndexOf(".") - 5) + ".pdf";
                        dateti = "_" + DateTime.Now.ToString().Replace("/", "_").Replace(":", "_").Replace(" AM", "").Replace(" PM", "").Replace(" ", "_");

                        if (File.Exists(CheckedPath))
                            CheckedPath = Properties.Settings.Default.CheckedPackingSlipMIGWithoutBPath + "\\" + fileName.Substring(0, fileName.LastIndexOf(".") - 5) + dateti + ".pdf";

                        //if (File.Exists(CheckedPath))
                        //{
                        //    File.Delete(CheckedPath);
                        //}
                    }
                    else
                    {
                        CheckedPath = Properties.Settings.Default.CheckedPackingSlipMIGWithoutBPath + "\\" + fileName;
                        dateti = "_" + DateTime.Now.ToString().Replace("/", "_").Replace(":", "_").Replace(" AM", "").Replace(" PM", "").Replace(" ", "_");

                        if (File.Exists(CheckedPath))
                            CheckedPath = Properties.Settings.Default.CheckedPackingSlipMIGWithoutBPath + "\\" + fileName.Substring(0, fileName.LastIndexOf(".")) + dateti + ".pdf";
                        //if (File.Exists(CheckedPath))
                        //{
                        //    File.Delete(CheckedPath);problem in convert barcode to flex numbers:
                        //}

                    }

                    System.IO.File.Move(fullName, CheckedPath);
                }
                catch (Exception ex)
                {
                    eventLogger.WriteEntry("fail move the file to CheckedPath" + ex.Message);

                }



            }
            ////////////////////////
            //SECOND LEVEL SIGN THE Packing Slip
            string ResponseSignedPS = dox.createPackingSlipLog(PackingSlipBarcode, CheckedPath);
            if (ResponseSignedPS == "MRB")
            {
                eventLogger.WriteEntry("Success MRB");
                if (migdal)
                {
                    PsGood_signed++;
                    PsGood_archive++;
                }
                else if (Mellanox)
                {
                    PsGoodMellanox_signed++;
                    PsGoodMellanox_archive++;
                }
                else
                {
                    PsGoodOfaqim_signed++;
                    PsGoodOfaqim_archive++;
                }
                return;
            }
            else if (ResponseSignedPS.Contains("MRB") && ResponseSignedPS.Length > 3)
            {
                eventLogger.WriteEntry("not Success MRB");
                if (migdal)
                {
                    PsBad++;
                }
                else if (Mellanox)
                {
                    PsBadMellanox++;
                }
                else
                    PsBadOfaqim++;
                return;

            }

            //
            string fullnameOld = CheckedPath;

            string CheckedPath1 = "";
            if (File.Exists(fullnameOld))
            {
                try
                {
                    n = 0;

                    //Migdal
                    //  if ((Path.GetFileNameWithoutExtension(fullnameOld).Length >= 4 && int.TryParse(Path.GetFileNameWithoutExtension(fullnameOld).Substring(0, 4), out n))
                    //||
                    //(Path.GetFileNameWithoutExtension(fullnameOld).Length >= 5 && Path.GetFileNameWithoutExtension(fullnameOld).StartsWith("BC"))
                    //   )
                    if (migdal)
                    {
                        CheckedPath1 = Settings.Default.CheckedPackingSlipMIGPath + "\\" + PackingSlipBarcode + ".pdf";
                        dateti = "_" + DateTime.Now.ToString().Replace("/", "_").Replace(":", "_").Replace(" AM", "").Replace(" PM", "").Replace(" ", "_");

                        if (File.Exists(CheckedPath1))
                            CheckedPath1 = Settings.Default.CheckedPackingSlipMIGPath + "\\" + PackingSlipBarcode + dateti + ".pdf";
                    }
                    else if (Mellanox) //Mellanox  
                    {
                        CheckedPath1 = Settings.Default.CheckedPackingSlipMIGPath + "\\" + "mellanox_" + PackingSlipBarcode + ".pdf";
                        dateti = "_" + DateTime.Now.ToString().Replace("/", "_").Replace(":", "_").Replace(" AM", "").Replace(" PM", "").Replace(" ", "_");
                        if (File.Exists(CheckedPath1))
                            CheckedPath1 = Settings.Default.CheckedPackingSlipMIGPath + "\\" + "mellanox_" + PackingSlipBarcode + dateti + ".pdf";

                    }
                    else
                    { //Ofakim
                        CheckedPath1 = Settings.Default.CheckedPackingSlipMIGPath + "\\" + "ofk_" + PackingSlipBarcode + ".pdf";
                        dateti = "_" + DateTime.Now.ToString().Replace("/", "_").Replace(":", "_").Replace(" AM", "").Replace(" PM", "").Replace(" ", "_");
                        if (File.Exists(CheckedPath1))
                            CheckedPath1 = Settings.Default.CheckedPackingSlipMIGPath + "\\" + "ofk_" + PackingSlipBarcode + dateti + ".pdf";

                    }
                    System.IO.File.Move(fullnameOld, CheckedPath1);
                }
                catch (Exception ex)
                {
                    eventLogger.WriteEntry("cant move file from: " + fullnameOld + " To: " + CheckedPath1 + " error: " + ex.Message, System.Diagnostics.EventLogEntryType.Warning);
                    dox.movetoDefQ(fullnameOld, 1, "Packing Slip", ex.Message);
                    if (migdal)
                    {
                        PsBad++;
                    }
                    else if (Mellanox)
                    {
                        PsBadMellanox++;
                    }
                    else
                        PsBadOfaqim++;
                    return;
                }
            }
            if (ResponseSignedPS == "Success")
            {
                eventLogger.WriteEntry("Success");
                if (migdal)
                {
                    PsGood_signed++;
                }
                else if (Mellanox)
                {
                    PsGoodMellanox_signed++;
                }
                else
                    PsGoodOfaqim_signed++;
            }
            else
            {
                eventLogger.WriteEntry("not Success");
                if (migdal)
                {
                    PsBad++;
                }
                else if (Mellanox)
                {
                    PsBadMellanox++;
                }
                else
                    PsBadOfaqim++;
            }

        }

        void handleInvoice(string fullName)
        {
            try
            {
                bool retry = true;
                int retries = 0;
                while (retry)
                {
                    try
                    {
                        string[] entityDesc = System.IO.File.ReadAllLines(fullName);
                        retry = false;
                    }
                    catch (Exception ex)
                    {
                        ++retries;
                        if (ex.Message.IndexOf("used by another process") > 0 && retries < 3)
                        {
                            eventLogger.WriteEntry("Cannot access " + fullName + " retry.\n" + ex.Message);
                            Thread.Sleep(2000);
                        }
                        else
                        {
                            eventLogger.WriteEntry("Cannot access " + fullName + "\n" + ex.Message);
                            retry = false;
                        }

                    }
                }

                if (fullName.IndexOf("_err") > 0)
                {
                    FileInfo prmFile = new FileInfo(fullName);
                    prmFile.MoveTo(fullName.Replace("_err", ""));
                    fullName = prmFile.FullName;
                }

                string response = dox.handleInvoice(fullName);
                if (response != "")
                {
                    eventLogger.WriteEntry(response, EventLogEntryType.Error);
                    InvBad++;
                }
                else
                {
                    InvGood++;
                }
            }
            catch (Exception ex)
            {
                eventLogger.WriteEntry(ex.Message, EventLogEntryType.Error);
                InvBad++;
            }
            try
            {
                FileInfo f = new FileInfo(fullName);
                f.MoveTo(Properties.Settings.Default.InvoiceLogPath + "\\" + f.Name + "_err");
                FileInfo inv = new FileInfo(fullName.Replace(".prm", ".pdf"));
                inv.MoveTo(Properties.Settings.Default.InvoiceLogPath + "\\" + inv.Name);
            }
            catch (Exception)
            {
                // Do nothing. This is the correct status, because the file should have been removed.
            }
        }

        void handleOrder99ReturedDocument(string fullName)
        {
            try
            {
                bool retry = true;
                int retries = 0;
                while (retry)
                {
                    try
                    {
                        string[] entityDesc = System.IO.File.ReadAllLines(fullName);
                        retry = false;
                    }
                    catch (Exception ex)
                    {
                        ++retries;
                        if (ex.Message.IndexOf("used by another process") > 0 && retries < 3)
                        {
                            eventLogger.WriteEntry("Cannot access " + fullName + " retry.\n" + ex.Message);
                            Thread.Sleep(2000);
                        }
                        else
                        {
                            eventLogger.WriteEntry("Cannot access " + fullName + "\n" + ex.Message);
                            retry = false;
                        }

                    }
                }
                string response = dox.handleOrder99ReturedDocument(fullName, "");
                if (response != "")
                {
                    eventLogger.WriteEntry(response, EventLogEntryType.Error);

                }

            }
            catch (Exception ex)
            {
                eventLogger.WriteEntry(ex.Message, EventLogEntryType.Error);

            }

        }
        void handleOrder99(string fullName)
        {

            string response = dox.handleOrder99(fullName);
            if (response != "")
            {
                eventLogger.WriteEntry(response, EventLogEntryType.Error);

            }
        }
        void BackUpRelevantFilesForStoreNext()
        {//check every invoice if it is for stornext then we copy it to a differant directory and when storenext file arrives we process it

            System.IO.DirectoryInfo invTray = new System.IO.DirectoryInfo(Properties.Settings.Default.NewInvoicePath);
            System.IO.FileInfo[] allFiles = invTray.GetFiles("*.prm");
            foreach (System.IO.FileInfo file in allFiles)
            {
                try
                {
                    //check if need to use the file for storenext (is company exist in storenext table)

                    string ClientID = "";
                    string ClientName = "";
                    string sInvNo = "";
                    int inv = 0;
                    int Company = 0;

                    // Read parameter file and load parameters
                    string[] entityDesc = System.IO.File.ReadAllLines(file.FullName);

                    foreach (string propLine in entityDesc)
                    {
                        string[] propArr = propLine.Split(':');
                        if (propArr.Length < 2) continue;

                        string propName = propArr[0];
                        string propVal = propArr[1].Substring(1, propArr[1].Length - 2);
                        switch (propName)
                        {
                            case "BAANCompany":
                                Company = int.Parse(propVal);
                                break;
                            case "ClientID":
                                ClientID = propVal;
                                break;
                            case "InvoiceNo":
                                sInvNo = propVal.Substring(propVal.IndexOf("/") + 1);
                                inv = int.Parse(sInvNo);
                                break;
                            case "ClientName":
                                ClientName = propVal;
                                break;
                        }
                    }


                    FlexInvoiceExt FlexInv = new FlexInvoiceExt(Company, inv.ToString());
                    bool isCompanyExist = false;
                    using (OdbcConnection DbConnection = new OdbcConnection(doxParams["BAANDB"]))
                    {
                        openConnection(DbConnection);
                        isCompanyExist = FlexInv.CheckIsCompnayExistInDB(DbConnection, ClientID);
                    }
                    eventLogger.WriteEntry("is Company Exist " + isCompanyExist.ToString());
                    if (isCompanyExist)
                    {
                        eventLogger.WriteEntry("copying file for storenext later ran, client id: " + ClientID);
                        try
                        {
                            file.CopyTo(Properties.Settings.Default.StoreNextPath + @"\" + file.Name);
                        }
                        catch (Exception)
                        {//if fail it is becuase file allready exist

                        }
                        try
                        {
                            //also copy the pdf
                            eventLogger.WriteEntry(invTray.FullName + @"\" + Path.GetFileNameWithoutExtension(file.Name) + ".pdf " + @"C:\Projects\FlexAutoArchiving\InvoiceForStorenext\" + Path.GetFileNameWithoutExtension(file.Name) + ".pdf");
                            File.Copy(invTray.FullName + @"\" + Path.GetFileNameWithoutExtension(file.Name) + ".pdf", Properties.Settings.Default.StoreNextPath + @"\" + Path.GetFileNameWithoutExtension(file.Name) + ".pdf");
                        }
                        catch (Exception)
                        {//if fail it is becuase file allready exist

                        }
                    }
                }
                catch (Exception ex)
                {
                    eventLogger.WriteEntry("Error on dox service function- BackUpRelevantFilesForStoreNext: " + ex.Message, EventLogEntryType.Error);
                }
            }
        }
    }
}
