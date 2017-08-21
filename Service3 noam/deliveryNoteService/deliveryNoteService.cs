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
using System.Web;

namespace deliveryNoteService
{
    public partial class deliveryNoteService : ServiceBase
    {

        const string logName = "DMS Applications Test";
        const string logSource = "Delivery Service Test";
        private DoxHandler dox;
        Dictionary<string, string> doxParams;
        private string today = DateTime.Today.ToShortDateString();
        string connectionString = "Data Source=mignt014;Initial Catalog=DoxPro_Env_Flex;Integrated Security=True";
        System.Timers.Timer invoice_timer = new System.Timers.Timer();
        int num1 = 6;
        System.Timers.Timer ShipmentGetMetadata_timer = new System.Timers.Timer();
        public int naama = 0;

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

        private int numOfDammaged = 0;//Program.svc.NumOfDammaged;
        public int NumOfDammaged
        {
            get
            {
                lock (this)
                {
                    return numOfDammaged;
                }
            }
            set
            {
                lock (this)
                {
                    numOfDammaged = value;
                }
            }
        }
        private string dateOfDammaged1="";
        public string DateOfDammaged1
        {
            get
            {
                lock (this)
                {
                    return dateOfDammaged1;
                }
            }
            set
            {
                lock (this)
                {
                    dateOfDammaged1 = value;
                }
            }
        }
        private string dateOfDammaged2="";
        public string DateOfDammaged2
        {
            get
            {
                lock (this)
                {
                    return dateOfDammaged2;
                }
            }
            set
            {
                lock (this)
                {
                    dateOfDammaged2 = value;
                }
            }
        }
        private string dateOfDammaged3="";
        public string DateOfDammaged3
        {
            get
            {
                lock (this)
                {
                    return dateOfDammaged3;
                }
            }
            set
            {
                lock (this)
                {
                    dateOfDammaged3 = value;
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


        protected override void OnStart(string[] args)
        {
           
          //  EventLog.CreateEventSource("tehila","Application");

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



                RestoreStatistics();
               // ResetStatistics();
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


            ShipmentGetMetadata_timer.Interval = 15000;
            ShipmentGetMetadata_timer.AutoReset = false;
            ShipmentGetMetadata_timer.Elapsed += new System.Timers.ElapsedEventHandler(ShipmentGetMetadata_timer_timer_Elapsed);
            ShipmentGetMetadata_timer.Start();
        }
        //stay
        private void ShipmentGetMetadata_timer_timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {//הוספתי
          //  ResetStatistics();
           // RestoreStatistics();
            ResetStatistics();

            System.Diagnostics.Debugger.Launch();
            ShipmentGetMetadata_timer.Enabled = false;
            try
            {
                using (IDataSupplier dataSupplier = DataManager.GetDataSupplier(DataManager.defaultType, connectionString))
                {
                    dataSupplier.OpenQuery();
                    eventLogger.WriteEntry("before GetData - ShipmentGetMetadata_timer_timer_Elapsed ");

                    DataSet ds = dataSupplier.GetData("SELECT [FileName],[lotsAndMakats],[type] from [shipmentArchive] where status='0' ");

                    if (ds != null && ds.Tables.Count > 0)
                    {
                        DataTable dt = ds.Tables[0];
                        if (dt != null && dt.Rows != null && dt.Rows.Count > 0)
                        {
                            foreach (DataRow dr in dt.Rows)
                            {
                                if (dr != null)
                                {

                                    string type = "";
                                    try
                                    {
                                        string response = dox.setMetaDataForShipment(dr[0].ToString(), dr[1].ToString(), type = dr[2].ToString(), 0);
                                        if (response == "")//Good
                                        {
                                            if (type == "COC")
                                                COCGetData++;
                                            else if (type == "Invoice")
                                                InvoiceGetData++;
                                            else
                                                DeliveryNoteGetData++;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        eventLogger.WriteEntry("Error in response - " + ex.ToString());
                                    }

                                }
                            }
                        }


                    }


                }




            }
            catch (Exception ex)
            {
                eventLogger.WriteEntry("Error in  select from ShipmentGetMetadata_timer_timer_Elapsed- " + ex.Message);
            }
           //Ester
            //קוד חדש לקבצים שלא נמצאו בטבלת הבאן השכתוב להחזירם לסיבוב נוסף
            try
            {
                string ifsuccess = "false";
                using (IDataSupplier dataSupplier = DataManager.GetDataSupplier(DataManager.defaultType, connectionString))
                {
                    dataSupplier.OpenQuery();
                    eventLogger.WriteEntry("before GetData - ShipmentGetMetadata_timer_timer_Elapsed- update from -3 to-1 ");


                    DataSet ds = dataSupplier.GetData("SELECT [FileName],[lotsAndMakats],[type] from [shipmentArchive] where status='-3' and DATEDIFF(HH,dateGetFile,getdate())>72 ");

                    if (ds != null && ds.Tables.Count > 0)
                    {
                        DataTable dt = ds.Tables[0];
                        if (dt != null && dt.Rows != null && dt.Rows.Count > 0)
                        {
                            foreach (DataRow dr in dt.Rows)
                            {
                                if (dr != null)
                                {


                                    string type = dr[2].ToString();


                                    try
                                    {//העברת הבקבצים לתיקית pdfsplit או לתיקית cdi-fromqueue
                                        
                                            System.IO.DirectoryInfo Directory = new System.IO.DirectoryInfo(Properties.Settings.Default.MachsanErrors);

                                            System.IO.FileInfo[] Files = Directory.GetFiles("*.pdf");
                                            foreach (System.IO.FileInfo current_file in Files)
                                            {


                                                string filename = Path.GetFileName(current_file.FullName);

                                                if (dr[0].ToString().Contains("CDI-FromQueue"))
                                                {
                                                    filename = Path.GetFileName(current_file.FullName);
                                                    if (dr[0].ToString().Contains(filename))
                                                    {


                                                        try
                                                        {

                                                            System.IO.File.Move(current_file.FullName, @"\\10.229.8.14\PackingSlips\NewPackingSlipsMIG\CDI-FromQueue\" + filename);
                                                            ifsuccess = "true";
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            eventLogger.WriteEntry("fail in -3" + ex.ToString());
                                                        }
                                                    }

                                                }

                                                else
                                                {
                                                    string fullfilename = "";
                                                    try
                                                    {

                                                        fullfilename = Path.GetFileName(current_file.FullName);
                                                        filename = fullfilename.Substring(filename.LastIndexOf('S'));
                                                    }

                                                    catch (Exception exp)
                                                    {
                                                        eventLogger.WriteEntry("Error in filename - " + "filename-" + fullfilename + exp.ToString());
                                                    }

                                                    if (dr[0].ToString().Contains(filename))
                                                    {
                                                        try
                                                        {
                                                            string PathPdfsplits = type + "\\" + "splitsPdf\\" + filename;
                                                            System.IO.File.Move(current_file.FullName, @"\\10.229.8.14\PackingSlips\NewPackingSlipsMIG\" + PathPdfsplits);
                                                            ifsuccess = "true";


                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            eventLogger.WriteEntry("fail in returningfile" + ex.ToString());
                                                        }
                                                    }
                                                }


                                            }


                                            if (ifsuccess == "true")
                                            {
                                                try
                                                {
                                                    string response = dox.setMetaDataForShipment(dr[0].ToString(), dr[1].ToString(), type = dr[2].ToString(), 3);
                                                    if (response == "")//Good
                                                    {
                                                        eventLogger.WriteEntry("tehila before");
                                                        dox.LotsAndMakats(dr[0].ToString(), dr[3].ToString());
                                                        if (type == "COC")
                                                            COCGetData++;
                                                        else if (type == "Invoice")
                                                            InvoiceGetData++;
                                                        else
                                                            DeliveryNoteGetData++;
                                                    }
                                                }
                                                catch (Exception ex)
                                                {
                                                    eventLogger.WriteEntry("Error in response - " + ex.ToString());
                                                }
                                            }


                                        }
                                    
                                    catch (Exception ex)
                                    {
                                        eventLogger.WriteEntry("Error in response - " + ex.ToString());
                                    }


                                  }
                                
                            }
                        }


                    }


                }




            }

            catch (Exception exe)
            { eventLogger.WriteEntry("fail in update status" + exe.Message); }

            //till here Ester
            eventLogger.WriteEntry("finished working on -3 now continue to status!=3" );

            //archive
            try
            {
                DataSet ds = null;
                string fileName = "";
                ////LIAT AMIR    Properties.Settings.Default.ReturnFiles
                StringCollection sc = Properties.Settings.Default.ReturnFiles;
                List<string> pathForfile = new List<string>();
                foreach (var filePath in sc)
                {

                    pathForfile.Add(filePath);
                }

                foreach (var filePath in pathForfile)
                {
                    eventLogger.WriteEntry("here in " + filePath);
                    System.IO.DirectoryInfo Directory = new System.IO.DirectoryInfo(filePath);
                    System.IO.FileInfo[] Files = Directory.GetFiles("*.pdf");
                    foreach (System.IO.FileInfo current_file in Files)
                    {
                        fileName = Path.GetFileName(current_file.FullName);
                        
                        try
                        {
                            
                             fileName = fileName.Substring(fileName.LastIndexOf("SKMB"), fileName.Length - fileName.LastIndexOf("SKMB"));
                             //if(fileName.Contains("coc_"))
                                 //System.IO.File.Move(filename, renameFile);
                                //refilename=fileName.Substring(fileName.IndexOf('_'));
                            // if (fileName.Contains("invoice_"))
                                 //refilename = fileName.Substring(fileName.IndexOf('_'));

                        }
                           
                        catch { }

                        using (IDataSupplier dataSupplier = DataManager.GetDataSupplier(DataManager.defaultType, connectionString))
                        {
                            try
                            {
                                ds = dataSupplier.GetData("SELECT top 1  fileName,lotsNew,MakatsNew,shippmentNo,supplierID,CompanyNo,SupplierName,SupplierAddress,PurchasingPerson" +
                              ",SupplierPhone,SearchKey,SupplierEmail,[type],dateGetFile  from [shipmentArchive] where status!='-3' and fileName like '%" +
                               fileName + "%' order by dateGetFile desc");
                            }
                            catch
                            {
                                eventLogger.WriteEntry("fail in  ds = dataSupplier.GetData");

                            }

                            DataTable dt = ds.Tables[0];
                            try
                            {
                                DataRow dr = dt.Rows[0];
                                if (dr != null)
                                {
                                    SaveStatistics();
                                    string response = "fail - ExceptionMY ";
                                    try
                                    {
                                        response = dox.archiveShipmentDoc(dr, filePath);
                                        eventLogger.WriteEntry("tehila before");
                                        dox.LotsAndMakats(dr[0].ToString(), dr[3].ToString());

                                    }
                                    catch
                                    {
                                        eventLogger.WriteEntry("fail whern archive 3rd service in  string response = dox.archiveShipmentDoc(dr);");
                                    }
                                    if (response == "")
                                    {

                                        using (IDataSupplier dataSupplier1 = DataManager.GetDataSupplier(DataManager.defaultType, connectionString))
                                        {
                                            dataSupplier.OpenQuery();
                                            dataSupplier.AddParameter("status", "2");
                                            dataSupplier.Execute("update [shipmentArchive] set status=@status where FileName='" + dr["Filename"].ToString() + "'");
                                            eventLogger.WriteEntry("file's status is 2!!!");

                                        }

                                        if (dr["type"].ToString() == "COC")
                                            COCGood++;
                                        else if (dr["type"].ToString() == "Invoice")
                                            InvoiceGood++;
                                        else
                                            DeliveryNoteGood++;
                                    }
                                    else if (response != "fail - ExceptionMY ")
                                    {
                                        if (dr["type"].ToString() == "COC")
                                            COCBad++;
                                        else if (dr["type"].ToString() == "Invoice")
                                            InvoiceBad++;
                                        else
                                            DeliveryNoteBad++;
                                    }
                                }
                            }
                            catch
                            {
                                eventLogger.WriteEntry(" 3rd service cant find in DB for file: " + fileName);
                            }
                            SaveStatistics();
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                eventLogger.WriteEntry("Error in  select from ShipmentArchive_timer_Elapsed- " + ex.Message);
            }
            SaveStatistics();



            ShipmentGetMetadata_timer.Enabled = true;
        }

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
                //InvGood = 0;
                //InvBad = 0;
                //PsBad = 0;
                //PsGood_archive = 0;
                //PsGood_signed = 0;
                //PsBadOfaqim = 0;
                //PsGoodOfaqim_archive = 0;
                //PsGoodOfaqim_signed = 0;
                COCBad = 0;
                COCGood = 0;
                InvoiceBad = 0;
                InvoiceGood = 0;
                DeliveryNoteBad = 0;
                DeliveryNoteGood = 0;
                //PsBadMellanox = 0;
                //PsGoodMellanox_archive = 0;
                //PsGoodMellanox_signed = 0;
                deliveryNoteGetData = 0;
                deliveryNoteGetData = 0;
                COCGetBarcode = 0;
                COCGetData = 0;
                InvoiceGetBarcode = 0;
                InvoiceGetData = 0;

                string folder = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                string path1 = Path.Combine(folder, "statisticsForDammaged.xml");
                try
                {
                    eventLogger.WriteEntry(DateTime.Now.AddMinutes(-5).ToString());
                    XmlDocument doc = new XmlDocument();
                    doc.Load(path1);

                    using (XmlReader reader = XmlReader.Create(path1))
                    {

                        
                        reader.Read();
                        reader.ReadStartElement("Statitics");
                        reader.ReadStartElement("Date");
                       string today1 = reader.ReadString();
                        reader.ReadEndElement();
                     


                        reader.ReadStartElement("NumOfDammaged");
                        numOfDammaged = reader.ReadContentAsInt();
                        reader.ReadEndElement();
                        eventLogger.WriteEntry("go to if");


                        XmlNodeList elemList = doc.GetElementsByTagName("dateOfDammaged1");
                        if (elemList.Count >= 1)
                            dateOfDammaged3 = elemList[elemList.Count - 1].InnerXml;
                        if (elemList.Count >= 2)
                            dateOfDammaged2 = elemList[elemList.Count - 2].InnerXml;
                        if (elemList.Count >= 3)
                            dateOfDammaged1 = elemList[elemList.Count - 3].InnerXml;

                        
                        eventLogger.WriteEntry("dfgfdgdfgfdgfd   " + numOfDammaged.ToString());
                    }
                }


                catch (Exception e)
                {
                    eventLogger.WriteEntry(e.Message + " hi!!  catch!!");
                   


                }
               
                 

               
                

            }
        }
        protected void SaveStatistics()
        {
            lock (this)
            {
                string folder = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
              
               
                
                string path = Path.Combine(folder, "statisticsForShipmment.xml");
                XmlTextWriter writer = new XmlTextWriter(path, null);
                //Write the root element
                writer.WriteStartElement("Statitics");
                //Write sub-elements
                writer.WriteElementString("Date", today);
                // writer.WriteElementString("PsGood_signed", PsGood_signed.ToString());
                // writer.WriteElementString("PsGood_archive", PsGood_archive.ToString());
                // writer.WriteElementString("PsBad", PsBad.ToString());
                // writer.WriteElementString("InvGood", InvGood.ToString());
                // writer.WriteElementString("InvBad", InvBad.ToString());
                //  writer.WriteElementString("PsGoodOfaqim_signed", PsGoodOfaqim_signed.ToString());
                // writer.WriteElementString("PsGoodOfaqim_archive", PsGoodOfaqim_archive.ToString());
                // writer.WriteElementString("PsBadOfaqim", PsBadOfaqim.ToString());
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
               // writer.WriteElementString("numOfDammaged", numOfDammaged.ToString());
                // writer.WriteElementString("PsGoodMellanox_signed", PsGoodMellanox_signed.ToString());
                // writer.WriteElementString("PsGoodMellanox_archive", PsGoodMellanox_archive.ToString());
                // writer.WriteElementString("PsBadMellanox", PsBadMellanox.ToString());

                // end the root element
                writer.WriteEndElement();

                //Write the XML to file and close the writer
                writer.Close();
                
            }
        }
        protected void RestoreStatistics()
        {
            eventLogger.WriteEntry("Tehila!!!!");
            string folder = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string path = Path.Combine(folder, "statisticsForShipmment.xml");
            string path1 = Path.Combine(folder, "statisticsForDammaged.xml");
            using (XmlReader reader = XmlReader.Create(path))
            {

                // Parse the XML document.  ReadString is used to 
                // read the text content of the elements.
                reader.Read();
                reader.ReadStartElement("Statitics");
                reader.ReadStartElement("Date");
                today = reader.ReadString();
                reader.ReadEndElement();
                //reader.ReadStartElement("PsGood_signed");
                //PsGood_signed = reader.ReadContentAsInt();
                //reader.ReadEndElement();
                //reader.ReadStartElement("PsGood_archive");
                //PsGood_archive = reader.ReadContentAsInt();
                //reader.ReadEndElement();

                //reader.ReadStartElement("PsBad");
                //PsBad = reader.ReadContentAsInt();
                //reader.ReadEndElement();

                //reader.ReadStartElement("InvGood");
                //InvGood = reader.ReadContentAsInt();
                //reader.ReadEndElement();

                //reader.ReadStartElement("InvBad");
                //InvBad = reader.ReadContentAsInt();
                //reader.ReadEndElement();




                //reader.ReadStartElement("PsGoodOfaqim_signed");
                //PsGoodOfaqim_signed = reader.ReadContentAsInt();
                //reader.ReadEndElement();

                //reader.ReadStartElement("PsGoodOfaqim_archive");
                //PsGoodOfaqim_archive = reader.ReadContentAsInt();
                //reader.ReadEndElement();
                //reader.ReadStartElement("PsBadOfaqim");
                //PsBadOfaqim = reader.ReadContentAsInt();
                //reader.ReadEndElement();
              

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



               /* //NumOfDammaged
                XmlDocument xml = new XmlDocument();
                xml.LoadXml("C:\\Projects\\SendMailDamgeFiles\\statisticsForShipmment.xml"); // suppose that myXmlString contains "<Names>...</Names>"

                XmlNodeList xnList = xml.SelectNodes("/NumOfDammaged/NumOfDammaged");
                foreach (XmlNode xn in xnList)
                {
                    string firstName = xn["NumOfDammaged"].InnerText;
                   
                    Console.WriteLine("Name: {0} {1}", firstName, lastName);
                }


                string folder1 = "C:\\Projects\\SendMailDamgeFiles";
                string path1 = Path.Combine(folder, "statisticsForShipmment.xml");
                XmlTextWriter writer1 = new XmlTextWriter(path1, null);
                writer1.WriteElementString("NumOfDammaged", NumOfDammaged.ToString());
                // end the root element
                writer1.WriteEndElement();

                //Write the XML to file and close the writer
                writer1.Close();*/











                //reader.ReadStartElement("PsGoodMellanox_signed");
                //PsGoodMellanox_signed = reader.ReadContentAsInt();
                //reader.ReadEndElement();

                //reader.ReadStartElement("PsGoodMellanox_archive");
                //PsGoodMellanox_archive = reader.ReadContentAsInt();
                //reader.ReadEndElement();
                //reader.ReadStartElement("PsBadMellanox");
                //PsBadMellanox = reader.ReadContentAsInt();
                //reader.ReadEndElement();
                reader.ReadEndElement();
                eventLogger.WriteEntry("Theirreaddd");

            }
            try
            {

                XmlDocument doc = new XmlDocument();
                doc.Load(path1);

                //Display all the book titles.
                //XmlNodeList elemList = doc.GetElementsByTagName("dateOfDammaged1");
                //dateOfDammaged1 = elemList[0].InnerXml;
                //eventLogger.WriteEntry("dateOfDammaged1: " + dateOfDammaged1);
               /* for (int i = 0; i < elemList.Count; i++)
                {
                    Console.WriteLine(elemList[i].InnerXml);
                }*/ 

                using (XmlReader reader = XmlReader.Create(path1))
                {
                    eventLogger.WriteEntry(DateTime.Now.AddMinutes(-5).ToString());
                    // Parse the XML document.  ReadString is used to 
                    // read the text content of the elements.
                    eventLogger.WriteEntry("Myreaddd");
                    reader.Read();
                    reader.ReadStartElement("Statitics");
                    reader.ReadStartElement("Date");
                    today = reader.ReadString();
                    reader.ReadEndElement();

                    reader.ReadStartElement("NumOfDammaged");
                    numOfDammaged = reader.ReadContentAsInt();
                    reader.ReadEndElement();
                    eventLogger.WriteEntry("go to if");


                                     
                    

                   // if (numOfDammaged == 1)
                    //{
                      XmlNodeList elemList = doc.GetElementsByTagName("dateOfDammaged1");
                    if(elemList.Count>=1)
                        dateOfDammaged3 = elemList[elemList.Count - 1].InnerXml;
                    if (elemList.Count >= 2)
                      dateOfDammaged2 = elemList[elemList.Count-2].InnerXml;
                    if (elemList.Count >= 3)
                        dateOfDammaged1 = elemList[elemList.Count - 3].InnerXml;
                      
                      //  dateOfDammaged2 = "tehila";
                        
                       

                      

                       // reader.ReadStartElement("dateOfDammaged1");
                        //dateOfDammaged1 = reader.ReadContentAsString();
                       // reader.ReadEndElement();
                  //  }
                   /* if (numOfDammaged == 2)
                    {
                        XmlNodeList elemList = doc.GetElementsByTagName("dateOfDammaged2");
                        dateOfDammaged1 = elemList[0].InnerXml;
                        eventLogger.WriteEntry("dateOfDammaged2: " + dateOfDammaged2);
                       // reader.ReadStartElement("dateOfDammaged2");
                        //dateOfDammaged2 = reader.ReadContentAsString();
                       // reader.ReadEndElement();
                    }
                    if (numOfDammaged == 3)
                    {
                        XmlNodeList elemList = doc.GetElementsByTagName("dateOfDammaged3");
                        dateOfDammaged1 = elemList[0].InnerXml;
                        eventLogger.WriteEntry("dateOfDammaged3: " + dateOfDammaged3);
                       // reader.ReadStartElement("dateOfDammaged3");
                       // dateOfDammaged3 = reader.ReadContentAsString();
                       // reader.ReadEndElement();

       

                    }
                    reader.ReadEndElement();*/
                    eventLogger.WriteEntry("dfgfdgdfgfdgfd   "+numOfDammaged.ToString());
                }
            }
            

            catch(Exception e)
            {
                eventLogger.WriteEntry(e.Message + " hi!!  catch!!");
                XmlDocument xml = new XmlDocument();
                xml.LoadXml(path1); // suppose that myXmlString contains "<Names>...</Names>"
                eventLogger.WriteEntry(path1);
            

                XmlNode xn = xml.SelectSingleNode("/Statitics/Statitics");

                string numOfDammaged11 = xn["NumOfDammaged"].InnerText;
             
                 
            }
          /*  try
            {

                s.Replace("%dateOfDammaged1%", svc.DateOfDammaged1.ToString());

                s = s.Replace("%dateOfDammaged2%", svc.DateOfDammaged2.ToString());

                s = s.Replace("%dateOfDammaged3%", svc.DateOfDammaged3.ToString());
            }
            catch (Exception e)
            {

            }*/
            /*folder = "C:\\Projects\\SendMailDamgeFiles";
             path= Path.Combine(folder, "statisticsForShipmment.xml");
       
            using (XmlReader reader = XmlReader.Create(path))
            {

                // Parse the XML document.  ReadString is used to 
                // read the text content of the elements.
                reader.Read();
                reader.ReadStartElement("Statitics");
                reader.ReadStartElement("Date");
                today = reader.ReadString();
                reader.ReadEndElement();



                reader.ReadStartElement("NumOfDammaged");
                NumOfDammaged = reader.ReadContentAsInt();
                reader.ReadEndElement();
                
                reader.ReadEndElement();

            }*/
          /*  //NumOfDammaged
            XmlDocument xml = new XmlDocument();
            xml.LoadXml("C:\\Projects\\SendMailDamgeFiles\\statisticsForShipmment.xml"); // suppose that myXmlString contains "<Names>...</Names>"

            XmlNodeList xnList = xml.SelectNodes("/NumOfDammaged/NumOfDammaged");
            foreach (XmlNode xn in xnList)
            {
                NumOfDammaged = 50;//Convert.ToInt32(xn["NumOfDammaged"].InnerText);
                
                
            }*/



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
