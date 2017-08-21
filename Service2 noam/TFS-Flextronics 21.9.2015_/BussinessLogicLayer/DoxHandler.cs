//============================================================================================
//
// This module was written as a part of a Windows Service program. The service is responsible
// for identifying the arrival of new documents and invoking this module to create DOX-Pro
// transactions.
//
// Collecting all the details to create DOX-Pro transactions involves reading parameter files
// that are supplied with the documents, and reading BAAN DB records directly.
// The program creates DOX-Pro documets to archive, but also creates the relevant binders for
// them if they do not exist already.
//============================================================================================

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data.Odbc;
using System.Diagnostics;
using System.Xml;
using DataAccessLayer;
using System.Threading;
using iTextSharp.text.pdf;
using System.Data;
using System.Text.RegularExpressions;
using Inlite.ClearImageNet;
using System.Data.SqlClient;
using System.Configuration;
using System.Data.OleDb;
using MySql.Data.MySqlClient;

namespace BussinessLogicLayer
{
    public class DoxHandler
    { 
        DOXAPI.ServiceSoapClient dox;   // DOX-Pro web-service API
        
        string token = string.Empty;    // DOX-Pro login token
        //private OdbcConnection DbConnection;    // BAAN DB access via ODBC

        // Objects in DOX-Pro implementation that are created at initiation and replicated
        // to save time in the processing of a new document. These are document types and 
        // DocType types of key DocType.
        private DOXAPI.DocType flexClientBinderType;
        private DOXAPI.DocType flexSupplierBinderType;
        private DOXAPI.DocType flexPackSlipType;
        private DOXAPI.DocType flexInvoiceType;
        private DOXAPI.DocType flexOrder99Type;
        private DOXAPI.DocType ShipmentDocType;
        private DOXAPI.DocType SupplierInvoiceDocType;
        private DOXAPI.DocType SupplierCOCDocType;
        private DOXAPI.DocType SupplierDeliveryNoteDocType;
        private DOXAPI.DocTypeAttribute flexClientBinder_customerID;
        private DOXAPI.DocTypeAttribute flexSupplierBinder_SupplierID;
        private DOXAPI.DocTypeAttribute flexPackSlip_PackSlipNo;
        private DOXAPI.DocTypeAttribute flexOrder99_Order99No;
        private DOXAPI.DocTypeAttribute ShipmentDoc_ShipmentDocNo;
        private DOXAPI.DocTypeAttribute SupplierInvoiceDoc_ShipmentDocNo;
        private DOXAPI.DocTypeAttribute SupplierCOCDoc_ShipmentDocNo;
        private DOXAPI.DocTypeAttribute SupplierDeliveryNoteDoc_ShipmentDocNo;

        PackingSlipXML inv_xml = new PackingSlipXML();
        public List<PackingSlipXMLEnvelopeLine> lines;
        OdbcDataReader DbReader;

        private string logFilesPath;
        private string logFilesOrd99Path;
        Dictionary<string, string> doxParams;   // Object containing parameters read from application.settings
        System.IO.TextReader testBarcodes;
        System.Diagnostics.EventLog logger;     // Event logging object defined by the Windows service
        string response = string.Empty;
        long binderID;
        private void LogArchive(string FileName, string KindOfDoc, int IsSuccess, int ReasonId, string exMessage = "")
        {
            try
            {
                try
                {
                    using (IDataSupplier dataSupplier = DataManager.GetDataSupplier(DataManager.defaultType, Settings.Default.ConnectionString))
                    {
                        dataSupplier.OpenQuery();
                        dataSupplier.AddParameter("FileName", FileName);
                        dataSupplier.AddParameter("KindOfDoc", KindOfDoc);
                        dataSupplier.AddParameter("IsSuccess", IsSuccess);
                        dataSupplier.AddParameter("ReasonId", ReasonId);
                        dataSupplier.Execute("INSERT INTO [LogArchive]([FileName],[KindOfDoc] ,[IsSuccess],[ReasonId],[date]) \r\n\t\t\t\t     VALUES ( @FileName,@KindOfDoc,@IsSuccess ,@ReasonId, getdate())");
                    }
                }
                catch (Exception ex)
                {
                    logger.WriteEntry("Error in  LogArchive - " + ex.Message);

                }
                ///temp
                if (IsSuccess == 0)
                {
                    try
                    {

                        using (IDataSupplier dataSupplier = DataManager.GetDataSupplier(DataManager.defaultType, Settings.Default.ConnectionString))
                        {
                            dataSupplier.OpenQuery();
                            dataSupplier.AddParameter("FileName", FileName);
                            dataSupplier.AddParameter("KindOfDoc", KindOfDoc);
                            dataSupplier.AddParameter("IsSuccess", IsSuccess);
                            dataSupplier.AddParameter("ReasonId", ReasonId);
                            dataSupplier.AddParameter("exMessage", exMessage);
                            dataSupplier.Execute("INSERT INTO [LogArchiveWithError]([FileName],[KindOfDoc] ,[IsSuccess],[ReasonId],[date],[exMessage]) \r\n\t\t\t\t     VALUES ( @FileName,@KindOfDoc,@IsSuccess ,@ReasonId, getdate(),@exMessage)");
                        }
                    }
                    catch (Exception ex)
                    {
                        logger.WriteEntry("Error in  LogArchiveWithError - " + ex.Message);

                    }
                }
                ////////here

            }
            catch (Exception ex)
            {
                logger.WriteEntry("Error in  LogArchive - " + ex.Message);

            }
        }
        public void LogArchiveShipment(string FileName, string KindOfDoc, int IsSuccess, int ReasonId, string exMessage = "")
        {
            try
            {
                try
                {
                    using (IDataSupplier dataSupplier = DataManager.GetDataSupplier(DataManager.defaultType, Settings.Default.ConnectionString))
                    {
                        dataSupplier.OpenQuery();
                        dataSupplier.AddParameter("FileName", FileName);
                        dataSupplier.AddParameter("KindOfDoc", KindOfDoc);
                        dataSupplier.AddParameter("IsSuccess", IsSuccess);
                        dataSupplier.AddParameter("ReasonId", ReasonId);
                        dataSupplier.Execute("INSERT INTO [LogArchiveShipment]([FileName],[KindOfDoc] ,[IsSuccess],[ReasonId],[date]) \r\n\t\t\t\t     VALUES ( @FileName,@KindOfDoc,@IsSuccess ,@ReasonId, getdate())");
                    }
                }
                catch (Exception ex)
                {
                    logger.WriteEntry("Error in  LogArchiveShipment - " + ex.Message);

                }
                ///temp
                if (IsSuccess == 0)
                {
                    try
                    {

                        using (IDataSupplier dataSupplier = DataManager.GetDataSupplier(DataManager.defaultType, Settings.Default.ConnectionString))
                        {
                            dataSupplier.OpenQuery();
                            dataSupplier.AddParameter("FileName", FileName);
                            dataSupplier.AddParameter("KindOfDoc", KindOfDoc);
                            dataSupplier.AddParameter("IsSuccess", IsSuccess);
                            dataSupplier.AddParameter("ReasonId", ReasonId);
                            dataSupplier.AddParameter("exMessage", exMessage);
                            dataSupplier.Execute("INSERT INTO [LogArchiveWithErrorShipment]([FileName],[KindOfDoc] ,[IsSuccess],[ReasonId],[date],[exMessage]) \r\n\t\t\t\t     VALUES ( @FileName,@KindOfDoc,@IsSuccess ,@ReasonId, getdate(),@exMessage)");
                        }
                    }
                    catch (Exception ex)
                    {
                        logger.WriteEntry("Error in  LogArchiveWithErrorShipment - " + ex.Message);

                    }
                }
                ////////here

            }
            catch (Exception ex)
            {
                logger.WriteEntry("Error in  LogArchiveShipment - " + ex.Message);

            }
        }
        public DoxHandler(Dictionary<string, string> _doxParams, System.Diagnostics.EventLog _logger)
        {
            doxParams = _doxParams;
            logger = _logger;
            //DbConnection = new OdbcConnection(doxParams["BAANDB"]);

            //openConnection();

            logFilesPath = doxParams["BAANLog"];
            try
            {
                dox = new DOXAPI.ServiceSoapClient();                   // Initialize web-service
                doxLogin();                                             // Login to DOX-Pro and keep token
                DOXAPI.DocType[] allTypes = dox.GetAllDocTypes(token);  // Get a list of doc-types from DOX-Pro

                // Find and keep the types used in this program
                foreach (DOXAPI.DocType dt in allTypes)
                {
                    logger.WriteEntry("Types: " + dt.ID + "/" + dt.Name);
                    if (dt.Name == "Customer Binder")
                        flexClientBinderType = dox.GetDocType(token, dt.ID);
                    if (dt.Name == "Customer Packing Slip")
                        flexPackSlipType = dox.GetDocType(token, dt.ID);
                    if (dt.Name == "Customer Invoice")
                        flexInvoiceType = dox.GetDocType(token, dt.ID);
                    if (dt.Name == "Supplier file")
                        flexSupplierBinderType = dox.GetDocType(token, dt.ID);
                    if (dt.Name == "Order99")
                        flexOrder99Type = dox.GetDocType(token, dt.ID);
                    if (dt.Name == "Shipment Doc")
                        ShipmentDocType = dox.GetDocType(token, dt.ID);
                    if (dt.Name == "Supplier Invoice")
                        SupplierInvoiceDocType = dox.GetDocType(token, dt.ID);
                    if (dt.Name == "Supplier COC")
                        SupplierCOCDocType = dox.GetDocType(token, dt.ID);
                    if (dt.Name == "Supplier Delivery Note")
                        SupplierDeliveryNoteDocType = dox.GetDocType(token, dt.ID);
                }

                // Find and keep key fields within types used in this program
                // Attributes are prototypes of fields - so they are fields of the doc-type
                for (int i = 0; i < flexClientBinderType.Attributes.Length; i++)
                {
                    if (flexClientBinderType.Attributes[i].Name == "Customer ID")
                        flexClientBinder_customerID = flexClientBinderType.Attributes[i];
                }
                for (int i = 0; i < flexPackSlipType.Attributes.Length; i++)
                {
                    if (flexPackSlipType.Attributes[i].Name == "Packing Slip No")
                        flexPackSlip_PackSlipNo = flexPackSlipType.Attributes[i];
                }
                for (int i = 0; i < flexSupplierBinderType.Attributes.Length; i++)
                {
                    if (flexSupplierBinderType.Attributes[i].Name == "Supplier No")
                    {
                        logger.WriteEntry("in attribute");
                        flexSupplierBinder_SupplierID = flexSupplierBinderType.Attributes[i];
                    }
                }
                for (int i = 0; i < flexOrder99Type.Attributes.Length; i++)
                {
                    if (flexOrder99Type.Attributes[i].Name == "Order99No")
                    {
                        logger.WriteEntry("in attribute Order99No");
                        flexOrder99_Order99No = flexOrder99Type.Attributes[i];
                    }
                }
                for (int i = 0; i < ShipmentDocType.Attributes.Length; i++)
                {
                    if (ShipmentDocType.Attributes[i].Name == "Shipment Doc No")
                    {
                        logger.WriteEntry("in attribute shipmentDoc");
                        ShipmentDoc_ShipmentDocNo = ShipmentDocType.Attributes[i];
                    }
                }
                for (int i = 0; i < SupplierInvoiceDocType.Attributes.Length; i++)
                {
                    if (ShipmentDocType.Attributes[i].Name == "Shipment Doc No")
                    {
                        logger.WriteEntry("in attribute shipmentDoc Invoice");
                        SupplierInvoiceDoc_ShipmentDocNo = ShipmentDocType.Attributes[i];
                    }
                }
                for (int i = 0; i < SupplierCOCDocType.Attributes.Length; i++)
                {
                    if (ShipmentDocType.Attributes[i].Name == "Shipment Doc No")
                    {
                        logger.WriteEntry("in attribute shipmentDoc COC");
                        SupplierCOCDoc_ShipmentDocNo = ShipmentDocType.Attributes[i];
                    }
                }
                for (int i = 0; i < SupplierDeliveryNoteDocType.Attributes.Length; i++)
                {
                    if (ShipmentDocType.Attributes[i].Name == "Shipment Doc No")
                    {
                        logger.WriteEntry("in attribute shipmentDoc Delivery Note");
                        SupplierDeliveryNoteDoc_ShipmentDocNo = ShipmentDocType.Attributes[i];
                    }
                }
               
                // If a type or fields does not exist, stop the program
                if (flexClientBinderType == null ||
                    flexPackSlipType == null ||
                    flexInvoiceType == null ||
                    flexClientBinder_customerID == null ||
                    flexPackSlip_PackSlipNo == null ||
                    flexOrder99Type == null ||
                    flexSupplierBinderType == null ||
                    ShipmentDocType == null ||
                    SupplierInvoiceDocType == null
                    || SupplierCOCDocType == null
                    || SupplierDeliveryNoteDocType == null)
                {
                    throw (new Exception("Flex-Dox Types not detected"));
                }
            }
                
            catch (Exception ex)
            {
                throw (ex);
            }
            logger.WriteEntry("before doxParams BarcodeFile ");
            // A debug barcodes file
            if (doxParams["BarcodeFile"] != String.Empty)
                testBarcodes = new StreamReader(doxParams["BarcodeFile"]);
            logger.WriteEntry("after doxParams BarcodeFile ");
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

        private void openMysqlConnection(MySqlConnection DbConnection)
        {
            try
            {
                if (DbConnection.State != System.Data.ConnectionState.Open)
                {
                    DbConnection.Open();
                }
            }
            catch (Exception ex)
            {
                throw (new Exception("Could not connect to DB with " + doxParams["BAANDB"] + "\n" + ex.Message));
            }
        }

        private void initalizeInvoiceAndCustomer(string prmFilename, FlexCustomer cst, FlexInvoice inv)
        {

            // The two following objects are simulating the fields of the DOX-Pro doc-types for
            // a customer binder and an invoice.

            logger.WriteEntry("in initalizeInvoiceAndCustomer ");

            // Read parameter file and load parameters
            string[] entityDesc = System.IO.File.ReadAllLines(prmFilename);

            foreach (string propLine in entityDesc)
            {
                string[] propArr = propLine.Split(':');
                if (propArr.Length < 2) continue;

                string propName = propArr[0];
                string propVal = propArr[1].Substring(1, propArr[1].Length - 2);
                switch (propName)
                {
                    case "BAANCompany":
                        cst.Company = propVal;
                        inv.Company = propVal;
                        break;
                    case "ClientID":
                        cst.ClientID = propVal;
                        inv.ClientID = propVal;
                        break;
                    case "ClientName":
                        cst.ClientName = propVal;
                        break;
                    case "ClientAddress":
                        cst.ClientAddress = propVal;
                        break;
                    case "InvoiceNo":
                        string sInvNo = propVal.Substring(propVal.IndexOf("/") + 1);
                        inv.InvoiceNo = int.Parse(sInvNo);
                        break;
                    case "InvoiceDate":
                        inv.IssueDate = DateTime.ParseExact(propVal, "dd-MM-yyyy", null);
                        break;
                    case "PackingSlips":
                        inv.PackingSlips = propVal;
                        break;
                    case "DocumentFile":    // The actual printed invoice file
                        inv.Filename = doxParams["NewInvoicePath"] + "/" + propVal;
                        break;
                }
            }

        }
        public string handlePsXml(string fullname, string xmlPath)
        {
            string line = "";
            string response = "";
            long packingSlipNo = 0;
            string packingSlipFromTxt = "";

            bool tryReadFromTxt = false;
            bool tryReadFromTxt2 = false;
            string barcode = string.Empty;
            string DirectoryCompany = "";
            int company = 0;
            try
            {
                logger.WriteEntry("try getBarcodes from file - " + fullname);


                // barcode = getBarcodes(fullname, true, true);
                if (File.Exists(fullname.Remove(fullname.LastIndexOf('.')) + ".txt"))
                {
                    using (FileStream fs = new FileStream(fullname.Remove(fullname.LastIndexOf('.')) + ".txt", FileMode.Open, FileAccess.Read))
                    {
                        using (StreamReader reader = new StreamReader(fs))
                        {
                            while ((line = reader.ReadLine()) != null)
                            {
                                if (line.IndexOf("Packing Slip No.") != -1)
                                {
                                    // packingSlipFromTxt = new String(line.ToCharArray().Where(c => Char.IsDigit(c)).ToArray());
                                    packingSlipFromTxt = line.Remove(0, line.LastIndexOf(":") + 2).Replace(" ", "");
                                    logger.WriteEntry("packingSlipString" + packingSlipFromTxt);
                                    tryReadFromTxt = long.TryParse(packingSlipFromTxt.ToString(), out packingSlipNo);
                                    DirectoryCompany = Path.GetDirectoryName(fullname).Remove(0, Path.GetDirectoryName(fullname).LastIndexOf(@"\") + 1);
                                    tryReadFromTxt2 = int.TryParse(DirectoryCompany.ToString(), out company);
                                    if (tryReadFromTxt && tryReadFromTxt2)
                                        return response = getDataFromBaan(xmlPath, company.ToString(), packingSlipFromTxt.ToString());

                                }
                            }
                            reader.Close();
                        }
                        fs.Close();
                    }
                }
                if (!tryReadFromTxt || !tryReadFromTxt2)
                {

                    barcode = getBarcodes(fullname, true, true);

                    if (barcode.IndexOf("Error") > -1 || barcode.Length < 15)
                    {
                        barcode = getBarcodesClearImage(fullname);
                        if (barcode.IndexOf("Error") > -1 || barcode.Length < 15)
                        {
                            return "Error reading the barcode - " + barcode;
                        }
                    }
                    try
                    {
                        long tmp = long.Parse(barcode.Replace(" ", "")); // will cause exception if not all characters are digits
                        string Company = barcode.Substring(0, 3);
                        string CustomerID = barcode.Substring(3, 6);
                        string PackingSlipNo = barcode.Substring(9);
                        return response = getDataFromBaan(xmlPath, Company, PackingSlipNo);
                    }
                    catch (Exception ex)
                    {
                        // If barcode reading yielded wrong data, the PDF is moved to manual archiving
                        return "Barcode data mismatch " + barcode + "  ,error - "
                            + ex.Message;
                    }
                }
                return "Error in get PackingSlip Number";

            }
            catch (Exception e)
            {

                return "fail getBarcodes from file - " + fullname + " ,Error - " + e.Message;

            }



        }
        public string handleInvoice(string prmFilename)
        {
            // A new invoice is delivered for archiving by sending a parameter file (text) and a pdf file
            // (the printed invoice) to a given location. The service finds them and calls this routine 
            // for creating a transaction to archive the invoice.

            // The two following objects are simulating the fields of the DOX-Pro doc-types for
            // a customer binder and an invoice.
            FlexCustomer cst = new FlexCustomer(flexClientBinderType, flexClientBinder_customerID);
            FlexInvoice inv = new FlexInvoice(flexInvoiceType);

            // Read parameter file and load parameters
            string[] entityDesc = System.IO.File.ReadAllLines(prmFilename);

            foreach (string propLine in entityDesc)
            {
                string[] propArr = propLine.Split(':');
                if (propArr.Length < 2) continue;

                string propName = propArr[0];
                string propVal = propArr[1].Substring(1, propArr[1].Length - 2);
                switch (propName)
                {
                    case "BAANCompany":
                        cst.Company = propVal;
                        inv.Company = propVal;
                        break;
                    case "ClientID":
                        cst.ClientID = propVal;
                        break;
                    case "ClientName":
                        cst.ClientName = propVal;
                        break;
                    case "ClientAddress":
                        cst.ClientAddress = propVal;
                        break;
                    case "InvoiceNo":
                        string sInvNo = propVal.Substring(propVal.IndexOf("/") + 1);
                        inv.InvoiceNo = int.Parse(sInvNo);
                        break;
                    case "InvoiceDate":
                        inv.IssueDate = DateTime.ParseExact(propVal, "dd-MM-yyyy", null);
                        break;
                    case "PackingSlips":
                        inv.PackingSlips = propVal;
                        break;
                    case "DocumentFile":    // The actual printed invoice file
                        inv.Filename = doxParams["NewInvoicePath"] + "/" + propVal;
                        break;
                }
            }
            // Login to DOX and save session token
            doxLogin();

            // If there is no customer binder yet (first invoice for this customer), then create
            // a new binder on DOX-Pro. For existing customers, some details may be updated frm BAAN.
            response = updateOrCreateBinder(cst);

            if (response != "")
            {
                return response;
            }

            string retMsg = "";
            if (inv.good)
            {
                long newDocID;

                // Get customer binder and invoice document as DOX-Pro objects
                DOXAPI.Binder bin = cst.asIDBinder();
                DOXAPI.Document doc = inv.asDocument(cst.FullID, cst.ClientName, inv.Filename);

                // This is the archiving transaction
                response = dox.Archive(token, doc, bin, "Invoices", false);

                // Successfull archiving returns the new entity ID. Otherwise, an error message is returned.
                if (long.TryParse(response, out newDocID))
                {
                    // If successfull, update other DOX-Pro documents
                    string[] packSlips = inv.PackingSlips.Split(',');
                    updatePackSlipsInvoiceNo(packSlips, cst.Company, inv.InvoiceNo);

                    Logger.Log((int)doc.DocType.ID, Logger.Operations.ArchiveDocument, Logger.Statuses.OK, (int)newDocID, inv.Filename, string.Empty, doc.Title, string.Empty);
                }
                else
                {
                    handleInvoiceError(prmFilename, inv.Filename, "Error creating invoice " + inv.InvoiceNo + "\n" + response);
                    Logger.Log((int)doc.DocType.ID, Logger.Operations.ArchiveDocument, Logger.Statuses.Error, -1, inv.Filename, string.Empty, doc.Title, string.Empty);
                    retMsg = "Error creating invoice " + inv.InvoiceNo + "\n" + response;
                    return retMsg;
                }
            }
            else
            {
                handleInvoiceError(prmFilename, inv.Filename, "Incomplete details");
                Logger.Log(18, Logger.Operations.ArchiveDocument, Logger.Statuses.Error, -1, inv.Filename, string.Empty, string.Empty, "incomplete details");
                retMsg = "Invoice " + inv.InvoiceNo + " could no be archived - incomplete details";
                return retMsg;
            }

            try
            { 
                //copying files is a temp change for checking a problem - please remove after
                string copyingFolderName = @"\\localhost\Projects\BackUpInvoices\";
                logger.WriteEntry(copyingFolderName + Path.GetFileNameWithoutExtension(inv.Filename) + Path.GetExtension(inv.Filename), System.Diagnostics.EventLogEntryType.Information);
                if (!(File.Exists(copyingFolderName + Path.GetFileNameWithoutExtension(inv.Filename) + Path.GetExtension(inv.Filename))))
                {
                    System.IO.File.Copy(inv.Filename, copyingFolderName + Path.GetFileNameWithoutExtension(inv.Filename) + Path.GetExtension(inv.Filename));
                    System.IO.File.Copy(prmFilename, copyingFolderName + Path.GetFileNameWithoutExtension(prmFilename) + Path.GetExtension(prmFilename));
                }
            }
            catch (Exception e)
            {
                logger.WriteEntry("Failed To Back up File to BackUpInvoices this is the error: " + e.Message, System.Diagnostics.EventLogEntryType.Warning);
            }
            try
            {


                System.IO.File.Delete(inv.Filename);
                System.IO.File.Delete(prmFilename);
            }
            catch (Exception e)
            {
                retMsg = "Exception during PDF or PRM file removal: " + e.Message;
            }
            return retMsg;
        }
        public string RunStoreNext(string prmFilename)
        {
            string retMsg = "";
            try
            {
                FlexCustomer cst = new FlexCustomer(flexClientBinderType, flexClientBinder_customerID);
                logger.WriteEntry("finished initalized new customer", System.Diagnostics.EventLogEntryType.Information);
                FlexInvoice inv = new FlexInvoice(flexInvoiceType);
                logger.WriteEntry("finished initalized new invoice", System.Diagnostics.EventLogEntryType.Information);
                initalizeInvoiceAndCustomer(prmFilename, cst, inv);
                //check if files exits
                //check if allready ran today

                logger.WriteEntry("reached storenext part " + doxParams["StorenextRanFile"], System.Diagnostics.EventLogEntryType.Information);

                if ((String.Compare(doxParams["EnableStorenextIntegration"], "true", true) != 0))
                {
                    logger.WriteEntry("Storenext flag is not enableded", System.Diagnostics.EventLogEntryType.Information);
                    // Remove files
                    try
                    {
                        System.IO.File.Delete((System.IO.Path.Combine(doxParams["StoreNextPath"], System.IO.Path.GetFileName(inv.Filename))));
                        System.IO.File.Delete(prmFilename);
                    }
                    catch (Exception e)
                    {
                        retMsg = "Exception during PDF or PRM file removal: " + e.Message;
                    }
                }
                else
                {
                    try
                    {

                        logger.WriteEntry("start processing for storenext " + inv.Company + ", invoice " + inv.InvoiceNo, System.Diagnostics.EventLogEntryType.Information);
                        using (OdbcConnection DbConnection = new OdbcConnection(doxParams["BAANDB"]))
                        {
                            openConnection(DbConnection);
                            FlexInvoiceExt inv_xml2 = new FlexInvoiceExt(int.Parse(inv.Company), inv.InvoiceNo.ToString());
                            logger.WriteEntry("after new FlexInvoiceExt", EventLogEntryType.Information);
                            if (inv_xml2.FetchFromDB(DbConnection, true, logger))
                            {
                                string xml_path = System.IO.Path.Combine(doxParams["StorenextXMLLandingFolder"], System.IO.Path.GetFileName(inv.Filename).Replace(".pdf", ".xml"));
                                //  string xml_path = System.IO.Path.Combine("C:\\backXmlForStorenext\\", System.IO.Path.GetFileName(inv.Filename).Replace(".pdf", ".xml"));
                                logger.WriteEntry("xml_path - " + xml_path);

                                logger.WriteEntry("stroenext pack slip: " + inv.PackingSlips, EventLogEntryType.Information);
                                try
                                {
                                    if (inv_xml2.SerializeToXML(xml_path, inv.PackingSlips, true, logger))
                                    {
                                        //if (inv_xml2.SerializeToXML(xml_path2, inv.PackingSlips, true, logger))
                                        //{
                                        logger.WriteEntry("in if (inv_xml2.SerializeToXML(xml_path, inv.PackingSlips, true))");
                                        try
                                        {
                                            System.IO.File.Move(System.IO.Path.Combine(doxParams["StoreNextPath"], System.IO.Path.GetFileName(inv.Filename)), System.IO.Path.Combine(doxParams["StorenextPDFLandingFolder"], System.IO.Path.GetFileName(inv.Filename)));
                                            System.IO.File.Delete(prmFilename);
                                            return "";
                                        }
                                        catch { }

                                        //}
                                        //else
                                        //{
                                        //    retMsg = "Cannot serialize invoice " + inv.InvoiceNo.ToString() + " to XML file \"" + xml_path + "\"";
                                        //}
                                    }
                                    else
                                    {
                                        retMsg = "Cannot serialize invoice " + inv.InvoiceNo.ToString() + " to XML file \"" + xml_path + "\"";
                                    }
                                }
                                catch (Exception ex)
                                {
                                    logger.WriteEntry(ex.Message + "/n/n/n/n/ " + ex.ToString());
                                }
                            }
                            else
                            {
                                retMsg = "Invoice was not found or connection to DB exhausted problem";
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        logger.WriteEntry(e.ToString());
                        retMsg = "Exception during XML message generating: " + e.Message;
                    }
                }
            }
            catch (Exception e)
            {
                logger.WriteEntry("Error in RunStorenext: " + e.ToString(), System.Diagnostics.EventLogEntryType.Error);
                throw e;
            }
            return retMsg;
        }
        private void handleInvoiceError(string prmFilename, string pdfFilename, string errMsg)
        {
            try
            {
                String logFileName = doxParams["InvoiceLogPath"] + "/" + System.IO.Path.GetFileName(prmFilename).Replace(".prm", ".log");

                File.Copy(prmFilename, logFileName, true);

                FileInfo logFI = new FileInfo(logFileName);
                StreamWriter invLogFile = logFI.AppendText();
                invLogFile.WriteLine();
                invLogFile.WriteLine(errMsg);
                invLogFile.Close();

                System.IO.FileInfo pdf = new FileInfo(pdfFilename);
                pdf.CopyTo(doxParams["InvoiceLogPath"] + "/" + pdf.Name);
                //Now we can delete original files
                System.IO.File.Delete(pdfFilename);
                System.IO.File.Delete(prmFilename);
            }
            catch (Exception ex)
            {
                Logger.Log(0, Logger.Operations.ArchiveDocument, Logger.Statuses.Error, -2, string.Empty, string.Empty, string.Empty, ex.Message);
            }
        }


        public string handleArchivePackingSlip(string fullname)
        {
            // A new packing slip is copied to a given location on the server by a scanner or other means.
            // The file is a PDF, and it has a barcode that contains key fields. This routine uses the OCR 
            // module to read the barcode and generate an archiving transaction to DOX-Pro.
            try
            {

                bool f = false;
                logger.WriteEntry("Handling " + fullname);
                string filename = Path.GetFileName(fullname);
                string response;
                doxLogin();
                string barcode = "";
                string pathBarCode = Path.GetFileNameWithoutExtension(fullname);
                logger.WriteEntry("pathBarCode - " + pathBarCode);
                if (pathBarCode.Contains("ofk_"))
                {
                    barcode = pathBarCode.Replace("ofk_", "");
                    pathBarCode = barcode;
                }
                else if (pathBarCode.Contains("mellanox_"))
                {
                    barcode = pathBarCode.Replace("mellanox_", "");
                    pathBarCode = barcode;
                }
                if (pathBarCode.Contains("_"))
                {
                    barcode = pathBarCode.Substring(0, pathBarCode.IndexOf("_"));
                    pathBarCode = barcode;
                }
                else if (barcode == "")
                    barcode = Path.GetFileNameWithoutExtension(fullname);
                logger.WriteEntry("get barcode " + barcode, System.Diagnostics.EventLogEntryType.Information);

                if (barcode.IndexOf("M") > 0) //its MRB Return Document
                {

                    response = handleOrder99ReturedDocument(fullname, barcode);
                    logger.WriteEntry("its MRB Return Document response - " + response);
                    return response + "MRB";
                }
                // The packing slip calls simulates DOX-Pro fields for the packing slip doc-type
                logger.WriteEntry("barcode - " + barcode);
                FlexPackingSlip ps = new FlexPackingSlip(flexPackSlipType, flexPackSlip_PackSlipNo);
                try
                {
                    long tmp = long.Parse(barcode.Replace(" ", "")); // will cause exception if not all characters are digits
                    ps.Company = barcode.Substring(0, 3);
                    ps.CustomerID = barcode.Substring(barcode.IndexOf(" "), barcode.Length - 6 - barcode.IndexOf(" "));
                    ps.PackingSlipNo = barcode.Remove(0, barcode.Length - 6);
                }
                catch (Exception ex)
                {
                    // If barcode reading yielded wrong data, the PDF is moved to manual archiving
                    logger.WriteEntry("Barcode data mismatch - fail");
                    movetoDefQ(fullname, 1, "Packing Slip", ex.Message);
                    Logger.Log(16, Logger.Operations.ArchiveDocument, Logger.Statuses.MovedToManual, -1, fullname, barcode, string.Empty, string.Empty);
                    return "Barcode data mismatch " + barcode;
                }
                ps.Filename = fullname;

                logger.WriteEntry("Barcode:" + barcode + " c:" + ps.Company + " cus:" + ps.CustomerID + " ps:" + ps.PackingSlipNo);
                string itemList = string.Empty;
                string order = string.Empty;

                // BAAN DB is accessed via ODBC for fetching related data
                using (OdbcConnection DbConnection = new OdbcConnection(doxParams["BAANDB"]))
                {
                    openConnection(DbConnection);
                    using (OdbcCommand DbCommand = DbConnection.CreateCommand())
                    {

                        DbCommand.CommandText = String.Format(doxParams["ItemsQ"], ps.Company, ps.PackingSlipNo);
                        OdbcDataReader DbReader;
                        // logger.WriteEntry(DbCommand.CommandText, System.Diagnostics.EventLogEntryType.Information);
                        try
                        {
                            DbReader = DbCommand.ExecuteReader();
                        }
                        catch (Exception ex)
                        {
                            logger.WriteEntry(DbCommand.CommandText + "\n" + ex.Message, System.Diagnostics.EventLogEntryType.Error);
                            movetoDefQ(fullname, 2, "Packing Slip", ex.Message);
                            Logger.Log(16, Logger.Operations.ArchiveDocument, Logger.Statuses.MovedToManual, -1, fullname, barcode, string.Empty, ex.Message);
                            return "Error retreiving items";
                        }

                        bool dbSucess = false;
                        while (DbReader.Read())
                        {
                            dbSucess = true;
                            itemList += DbReader.GetString(0).Trim() + ",";
                            ps.IssueDate = DbReader.GetDateTime(1);
                            order = DbReader.GetString(2);
                        }
                        ps.ItemList = itemList.TrimEnd(',');
                        DbReader.Close();
                        if (!dbSucess)
                        {
                            logger.WriteEntry(DbCommand.CommandText + " - The PS isn't exists in Baan");
                            movetoDefQ(fullname, 2, "packing slip");
                            Logger.Log(16, Logger.Operations.ArchiveDocument, Logger.Statuses.MovedToManual, -1, fullname, barcode, string.Empty, "Error retrieving packing slip details for psno " + ps.Company + "/" + ps.PackingSlipNo + ". Sent to manual queue");
                            return "Error retrieving packing slip details for psno " + ps.Company + "/" + ps.PackingSlipNo + ". Sent to manual queue";
                        }

                        DbCommand.CommandText = String.Format(doxParams["CustQ"], ps.Company, order);
                        try
                        {
                            DbReader = DbCommand.ExecuteReader();
                        }
                        catch (Exception ex)
                        {
                            logger.WriteEntry(DbCommand.CommandText + "\n" + ex.Message, System.Diagnostics.EventLogEntryType.Error);
                            movetoDefQ(fullname, 6, "Packing Slip", ex.Message);
                            Logger.Log(16, Logger.Operations.ArchiveDocument, Logger.Statuses.MovedToManual, -1, fullname, barcode, string.Empty, ex.Message);
                            return "Error getting customer order";
                        }
                        while (DbReader.Read())
                        {
                            ps.CustomerOrderNo = DbReader.GetString(0);
                        }
                        DbReader.Close();
                    }
                }
                // This objhect simulates the customer binder in DOX-Pro
                FlexCustomer cst = new FlexCustomer(flexClientBinderType, flexClientBinder_customerID);
                cst.Company = ps.Company;
                cst.ClientID = ps.CustomerID;

                // If there is no customer binder yet (first invoice for this customer), then create
                // a new binder on DOX-Pro. For existing customers, some details may be updated frm BAAN.
                try
                {
                    response = updateOrCreateBinder(cst);
                }
                catch (Exception ex)
                {
                    response = "Could not establish DOX-API call: " + ex.Message;
                }
                if (response != "")
                {
                    movetoDefQ(fullname, 5, "Packing Slip", response);
                    return response;
                }

                // Get customer binder and packing slip document as DOX-Pro objects
                DOXAPI.Binder clientBinder = cst.asIDBinder();
                DOXAPI.Document docPS = ps.asDocument();

                // Try to fetch a related DOX-Pro object (invoice) to get its number
                string invNo = findInvoicefor(ps.Company, ps.PackingSlipNo);
                if (invNo != "")
                {
                    DOXFields.SetField(docPS, "Invoice No", invNo);
                    logger.WriteEntry("Found matching invoice: " + invNo);
                }

                // Archive a new packing slip in DOX-Pro
                response = dox.Archive(token, docPS, clientBinder, "Packing Slips", false);
                logger.WriteEntry("response after archive function : " + response);
                long docID;
                if (long.TryParse(response, out docID))
                {
                    // If successfull, create a log that is read by a BAAN job for marking the reciept of a
                    // packing slip in BAAN
                    /////    createPackingSlipLog(ps);
                    System.IO.File.Delete(fullname);
                    LogArchive(fullname, "Packing Slip", 1, 7);
                    Logger.Log((int)docPS.DocType.ID, Logger.Operations.ArchiveDocument, Logger.Statuses.OK, (int)docID, fullname, barcode, docPS.Title, string.Empty);
                }
                else
                {
                    movetoDefQ(fullname, 3, "Packing Slip", response);
                    Logger.Log(16, Logger.Operations.ArchiveDocument, Logger.Statuses.MovedToManual, -1, fullname, barcode, string.Empty, response);
                    return "Packing Slip No. " + ps.PackingSlipNo + " could not be archived: " + response
                                + ".\nFile moved to default queue";
                }
                return "";
            }
            catch (Exception ex)
            {
                Logger.Log(16, Logger.Operations.ArchiveDocument, Logger.Statuses.Error, -1, fullname, string.Empty, string.Empty, ex.Message);
                logger.WriteEntry("Exception in handle packing slip " + ex.Message);
                movetoDefQ(fullname, 3, "Packing Slip", ex.Message);
                return "Exception while handling packing slip " + ex.Message;
            }
        }
        public string GetBarcodesForPS(string fullname)
        {
            string barcode = "";
            try
            {
                barcode = getBarcodes(fullname, true, true);
            }
            catch (Exception ex)
            {
                // PackingSlipBarcode.ToLower().Contains("POINT".ToLower()
                logger.WriteEntry("catch fail read barcode 2 times - this is returend barcode " + barcode + " Exc: " + ex.Message);
                barcode = ex.Message + " - Error";
            }
            if (barcode.ToLower().Contains("POINT".ToLower()) || barcode == "" || barcode == "Error" || barcode.Length < 15)
            {
                barcode = getBarcodesClearImage(fullname);
            }
            return barcode;
        }

        public string handlePackingSlip(string fullname)
        {
            // A new packing slip is copied to a given location on the server by a scanner or other means.
            // The file is a PDF, and it has a barcode that contains key fields. This routine uses the OCR 
            // module to read the barcode and generate an archiving transaction to DOX-Pro.
            //try
            //{
            ////    Thread.Sleep(500);
            ////    bool f = false;
            ////    string error = "";
            ////    logger.WriteEntry("Handling " + fullname);
            ////    string filename = Path.GetFileName(fullname);
            ////    string response;
            ////    //doxLogin();
            ////    string barcode = "";
            ////    //logger.WriteEntry("try to do Path.GetFileNameWithoutExtension(fullname).StartsWith(BC)", System.Diagnostics.EventLogEntryType.Information);
            ////    if (Path.GetFileNameWithoutExtension(fullname).StartsWith("BC"))
            ////    {
            ////        f = true;
            ////        barcode = "400  " + Path.GetFileNameWithoutExtension(fullname).Substring(2);
            ////        logger.WriteEntry("barcode " + barcode, System.Diagnostics.EventLogEntryType.Warning);


            ////    }
            ////    else if (doxParams["BarcodeFile"] != String.Empty) // Barcodes test file
            ////    {
            ////        try
            ////        {
            ////            barcode = testBarcodes.ReadLine();
            ////            logger.WriteEntry(" barcode = testBarcodes.ReadLine(); " + barcode);
            ////        }
            ////        catch (Exception e)
            ////        {
            ////            error = "  :  " + e.Message + " \n " + e.ToString();
            ////            barcode = "Error";

            ////        }
            ////    }
            ////    else  // Real barcodes
            ////    {

            ////    }


            ////   // logger.WriteEntry("barcode.IndexOf(M):" + barcode.IndexOf("M"), System.Diagnostics.EventLogEntryType.Information);
            ////    if (barcode.IndexOf("M") > 0) //its MRB Return Document
            ////    {

            ////        response = handleOrder99ReturedDocument(fullname);
            ////        return response;
            ////    }
            ////    // The packing slip calls simulates DOX-Pro fields for the packing slip doc-type

            ////    Thread.Sleep(100);

            ////    ps.Filename = fullname;
            ////    try
            ////    {
            ////        createPackingSlipLog(ps);
            ////    }
            ////    catch(Exception ex)
            ////    {
            ////        logger.WriteEntry("Not Creating PS log at " + barcode + " because: " + ex.Message, System.Diagnostics.EventLogEntryType.Warning);
            ////    }
            ////    string CheckedPath = "";
            ////    string fullnameOld = fullname;
            ////    try
            ////    {
            ////        int n = 0;

            ////        //Migdal
            ////        if ((Path.GetFileNameWithoutExtension(fullname).Length >= 4 && int.TryParse(Path.GetFileNameWithoutExtension(fullname).Substring(0, 4), out n))
            ////      ||
            ////      (Path.GetFileNameWithoutExtension(fullname).Length >= 5 && Path.GetFileNameWithoutExtension(fullname).StartsWith("BC"))
            ////    )
            ////            CheckedPath = Settings.Default.CheckedPackingSlipMIGPath +"\\"+ barcode + ".pdf";

            ////        else //Ofakim
            ////            CheckedPath = Settings.Default.CheckedPackingSlipMIGPath + "\\"+"ofk_" + barcode + ".pdf";
            ////        System.IO.File.Move(fullname, CheckedPath);
            ////    }
            ////    catch (Exception ex)
            ////    {
            ////        logger.WriteEntry("cant move file from: "+ fullname + " To: " + CheckedPath + " error: " + ex.Message, System.Diagnostics.EventLogEntryType.Warning);
            ////        movetoDefQ(fullnameOld, 1, "Packing Slip", ex.Message);
            ////    }

            //    return "Sucsses";

            //}
            //catch(Exception ex)
            //{
            //    movetoDefQ(fullname, 1, "Packing Slip", ex.Message);
            //    logger.WriteEntry("Error because: " + ex.Message, System.Diagnostics.EventLogEntryType.Warning);
            //    return "NotSuccess "+ex.Message;
            //}
            try
            {
                Thread.Sleep(300); bool f = false;
                string error = "";
                logger.WriteEntry("Handling " + fullname);
                string filename = Path.GetFileName(fullname);
                string response;
                //doxLogin();
                string barcode = "";
                //logger.WriteEntry("try to do Path.GetFileNameWithoutExtension(fullname).StartsWith(BC)", System.Diagnostics.EventLogEntryType.Information);
                if (Path.GetFileNameWithoutExtension(fullname).StartsWith("BC"))
                {
                    f = true;
                    barcode = "400  " + Path.GetFileNameWithoutExtension(fullname).Substring(2);
                    logger.WriteEntry("barcode " + barcode, System.Diagnostics.EventLogEntryType.Warning);


                }
                else if (doxParams["BarcodeFile"] != String.Empty) // Barcodes test file
                {
                    try
                    {
                        barcode = testBarcodes.ReadLine();
                        logger.WriteEntry(" barcode = testBarcodes.ReadLine(); " + barcode);
                    }
                    catch (Exception e)
                    {
                        error = "  :  " + e.Message + " \n " + e.ToString();
                        barcode = "Error";

                    }
                }
                else  // Real barcodes
                {
                    try
                    {
                        barcode = getBarcodes(fullname, true, true);
                    }
                    catch (Exception ex)
                    {
                        logger.WriteEntry("fail read barcode 2 times - this is returend barcode " + barcode + " Exc: " + ex.Message);
                        barcode = "Error";
                    }
                }

                if (barcode.IndexOf("Error") > -1 || barcode.Length < 15)
                {
                    barcode = getBarcodesClearImage(fullname);
                    if (barcode.IndexOf("Error") > -1 || barcode.Length < 15)
                    {
                        logger.WriteEntry("returned barcode is invalid: " + barcode + " move to queue movetoDefQ  ");
                        // If failed to read barcode, the PDF is moved to a location of a DOX-Pro queue
                        movetoDefQ(fullname, 1, "Packing Slip", "barcode.IndexOf(Error)>-1 || barcode.Length < 15");
                        if (barcode.Length < 15)
                            error = "the barcode.Length < 15 ";
                        return "handlePackingSlip" + error + " - the error the barcode -" + barcode;
                    }
                }
                // logger.WriteEntry("barcode.IndexOf(M):" + barcode.IndexOf("M"), System.Diagnostics.EventLogEntryType.Information);
                if (barcode.ToLower().IndexOf("M".ToLower()) > 0) //its MRB Return Document
                {

                    response = handleOrder99ReturedDocument(fullname, barcode);
                    return response;
                }
                // The packing slip calls simulates DOX-Pro fields for the packing slip doc-type

                Thread.Sleep(100);
                FlexPackingSlip ps = new FlexPackingSlip(flexPackSlipType, flexPackSlip_PackSlipNo);
                try
                {
                    long tmp = long.Parse(barcode.Replace(" ", "")); // will cause exception if not all characters are digits
                    //400  5678911212
                    ps.Company = barcode.Substring(0, 3);
                    //if (f == false || (barcode.Length < 14 && f))
                    //{
                    //    ps.CustomerID = barcode.Substring(3, 6);
                    //    ps.PackingSlipNo = barcode.Substring(9);
                    //}
                    //else
                    //{
                    //    ps.CustomerID = barcode.Substring(3, 7);
                    //    ps.PackingSlipNo = barcode.Substring(10);
                    //}
                    ps.CustomerID = (barcode.Substring(barcode.IndexOf(" "), barcode.Length - 6 - barcode.IndexOf(" ")));
                    ps.PackingSlipNo = barcode.Remove(0, barcode.Length - 6);
                }
                catch (Exception ex)
                {
                    logger.WriteEntry("barcode - " + barcode + " ,problem in convert barcode to flex numbers: " + ex.Message);
                    movetoDefQ(fullname, 1, "Packing Slip", ex.Message);
                    Logger.Log(16, Logger.Operations.ArchiveDocument, Logger.Statuses.MovedToManual, -1, fullname, barcode, string.Empty, string.Empty);
                    return "Barcode data mismatch " + barcode;
                }
                ps.Filename = fullname;
                //try
                //{
                //    createPackingSlipLog(ps);
                //}
                //catch (Exception ex)
                //{
                //    logger.WriteEntry("Not Creating PS log at " + barcode + " because: " + ex.Message, System.Diagnostics.EventLogEntryType.Warning);
                //}
                string CheckedPath = "";
                string fullnameOld = fullname;
                try
                {
                    int n = 0;

                    //Migdal
                    if ((Path.GetFileNameWithoutExtension(fullname).Length >= 4 && int.TryParse(Path.GetFileNameWithoutExtension(fullname).Substring(0, 4), out n))
                  ||
                  (Path.GetFileNameWithoutExtension(fullname).Length >= 5 && Path.GetFileNameWithoutExtension(fullname).StartsWith("BC"))
                )
                        CheckedPath = Settings.Default.CheckedPackingSlipMIGPath + "\\" + barcode + ".pdf";

                    else //Ofakim
                        CheckedPath = Settings.Default.CheckedPackingSlipMIGPath + "\\" + "ofk_" + barcode + ".pdf";
                    System.IO.File.Move(fullname, CheckedPath);
                }
                catch (Exception ex)
                {
                    logger.WriteEntry("cant move file from: " + fullname + " To: " + CheckedPath + " error: " + ex.Message, System.Diagnostics.EventLogEntryType.Warning);
                    movetoDefQ(fullnameOld, 1, "Packing Slip", ex.Message);
                }

                return "Sucsses";

            }
            catch (Exception ex)
            {
                movetoDefQ(fullname, 1, "Packing Slip", ex.Message);
                logger.WriteEntry("Error because: " + ex.Message, System.Diagnostics.EventLogEntryType.Warning);
                return "NotSuccess " + ex.Message;
            }
        }

        private string getBarcodes(string filename, bool fromTop, bool sum15)
        {
            /*if (fromTop)
                dox.SetBarCodeClipRectangle(0.0, 0.05, 0.0, 0.25);
            else
                dox.SetBarCodeClipRectangle(0.0, 0.75, 0.0, 0.95);
            */

            Abbyy10.FineReader reader = new Abbyy10.FineReader();
            //TODO - changed hardcoded path to the proeprty
            reader.AbbyyPath = @"C:\Program Files (x86)\ABBYY SDK\10\FineReader Engine\Bin\";//System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            reader.ClippingRectangle = new Abbyy10.Rectangle(0, 0, 0, 0);

            //   if (reader.ClippingRectangle == null)
            //      reader.ClippingRectangle = new Rectangle(0.0, 0.0, 1.0, 1.0);

            List<string> barcodes = reader.GetBarCodes(filename);
            if (barcodes.Count == 0 || barcodes[0] == "" || barcodes[0].Length < 5 || barcodes[0].Substring(0, 5) == "Error" || (sum15 == true && barcodes[0].Length < 15) || (sum15 == true && (barcodes[0].Length > 17 || !barcodes[0].StartsWith("4"))))
            {

                barcodes = ExecuteCommand(filename);
                if (barcodes.Count == 0 || barcodes[0].Contains("Error") || barcodes[0].Length < 15)
                {
                    barcodes = ExecuteCommandOcr(filename);
                }
            }

            if (barcodes.Count == 0)//Length 
            {
                logger.WriteEntry("No barcode recognized on " + filename, System.Diagnostics.EventLogEntryType.Warning);
                return "";
            }
            logger.WriteEntry("barcodes.Count " + barcodes.Count + "\n" + filename + "\n" + barcodes[0], System.Diagnostics.EventLogEntryType.Information);
            if (barcodes[0] == "")
            {
                logger.WriteEntry("Empty barcode on " + filename, System.Diagnostics.EventLogEntryType.Warning);
                return "";
            }
            if (barcodes[0].Length < 5)
            {
                logger.WriteEntry("Too short barcode in " + filename + "\n" + barcodes[0], System.Diagnostics.EventLogEntryType.Warning);
                return "";
            }
            if (barcodes[0].Substring(0, 5) == "Error")
            {
                logger.WriteEntry("Error reading barcode in " + filename + "\n" + barcodes[0], System.Diagnostics.EventLogEntryType.Warning);
                return "";
            }

            return barcodes[0];
        }

        private List<string> ExecuteCommandOcr(string filename)
        {

            List<string> barcode = new List<string>();
            string error = "";
            try
            {

                string pdf = @"""" + filename + @"""";
                try
                {

                    FileStream fs1 = new FileStream("barcode.bat", FileMode.OpenOrCreate, FileAccess.Write);
                    StreamWriter writer = new StreamWriter(fs1);
                    writer.WriteLine(@"cd C:\Program Files\DMS Inc\DeliveryNotesSetup\GetBarcodeFromFileCMD");

                    writer.WriteLine(@"OcrGetBarcodes.exe  " + pdf);
                    writer.Close();
                }
                catch (Exception ex)
                {

                    //barcode.Clear();
                    barcode.Add("Error");
                    error = "Error";
                    return barcode;
                }
                System.Diagnostics.Process proc = new System.Diagnostics.Process();
                proc.StartInfo.FileName = "barcode.bat";
                proc.StartInfo.Verb = "runas";
                proc.StartInfo.WorkingDirectory = @"";
                proc.StartInfo.CreateNoWindow = true;
                proc.StartInfo.UseShellExecute = false;
                proc.StartInfo.RedirectStandardError = true;
                proc.StartInfo.RedirectStandardOutput = true;
                proc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                proc.Start();
                string output = proc.StandardOutput.ReadToEnd();
                proc.Close();
                logger.WriteEntry("output Ocr -" + output);


                if (error == "")
                {
                    try
                    {

                        string[] lines = output.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
                        string bercodeOcr = "", PsNum2 = "", PsNum1 = "";
                        foreach (string item in lines)
                        {


                            if (item.Contains("400M ") || item.Contains("400 ") || item.Contains("421 "))
                            {
                                bercodeOcr = item.Remove(0, item.IndexOf("4") > 0 ? item.IndexOf("4") : 0);
                                if (item.Contains("400M "))
                                    bercodeOcr = bercodeOcr.Replace("400M ", "400M  ");
                                else if (item.Contains("400 "))
                                    bercodeOcr = bercodeOcr.Replace("400 ", "400  ");
                                else if (item.Contains("421 "))
                                    bercodeOcr = bercodeOcr.Replace("421 ", "421  ");

                                PsNum1 = bercodeOcr;
                                logger.WriteEntry("PsNum1-" + PsNum1);
                                if (PsNum2 != "" && PsNum1.Contains(PsNum2))
                                {

                                    barcode.Add(PsNum1);
                                    logger.WriteEntry(PsNum1);
                                    break;

                                }


                            }
                            else if (item.Contains("Packing Slip No.: "))
                            {
                                bercodeOcr = item.Replace("Packing Slip No.: ", "").ToString();
                                PsNum2 = bercodeOcr;
                                logger.WriteEntry("PsNum2-" + PsNum2);

                                if (PsNum1 != "" && PsNum1.Contains(PsNum2))
                                {
                                    PsNum1 = PsNum1.Remove(0, PsNum1.IndexOf("4") > 0 ? PsNum1.IndexOf("4") : 0);
                                    barcode.Add(PsNum1);
                                    logger.WriteEntry(PsNum1);
                                    break;

                                }



                            }
                        }


                    }
                    catch (Exception ex)
                    {

                        error = "Error";

                        barcode.Add("Error");
                        return barcode;


                    }
                    if (File.Exists(filename.Remove(filename.LastIndexOf('.')) + ".txt"))
                        File.Delete(filename.Remove(filename.LastIndexOf('.')) + ".txt");
                }
                else if (output.Contains("N/A") || barcode.Count == 0)
                {
                    error = "Error";// error2 = "Error";
                }


            }
            catch (Exception ex)
            {
                error = "Error";

                barcode.Add("Error");
                return barcode;

            }
            return barcode;
        }

        public string createPackingSlipLog(string barcode, string fullname)
        {
            if (barcode.ToLower().IndexOf("M".ToLower()) > 0) //its MRB Return Document
            {
                logger.WriteEntry("in   if (barcode.IndexOf(M) > 0) " + barcode);
                response = handleOrder99ReturedDocument(fullname, barcode);
                return response
                    + "MRB";
            }
            //bool f = false;
            FlexPackingSlip ps = new FlexPackingSlip(flexPackSlipType, flexPackSlip_PackSlipNo);
            try
            {
                long tmp = long.Parse(barcode.Replace(" ", "")); // will cause exception if not all characters are digits
                //400  5678911212
                ps.Company = barcode.Substring(0, 3);
                string barcode2 = barcode;
                barcode2 = barcode2.Remove(0, 3);
                //if ((barcode.Length < 14))
                //{
                //    ps.CustomerID = barcode.Substring(3, 6);
                //    ps.PackingSlipNo = barcode.Substring(9);
                //}
                //else
                //{
                //    ps.CustomerID = barcode.Substring(3, 7);
                //    ps.PackingSlipNo = barcode.Substring(10);
                //}

                //ps.Company = barcode.Substring(0, 3);
                //    ps.CustomerID = barcode.Substring(barcode.LastIndexOf(" "), barcode.Length - 6 - barcode.IndexOf(" ") - 1);
                ps.CustomerID = barcode2.Substring(0, (barcode2.Length - 6));
                //    ps.PackingSlipNo = barcode.Remove(0, barcode.Length - 6);
                ps.PackingSlipNo = barcode2.Remove(0, barcode2.Length - 6);
            }

            catch (Exception ex)
            {
                logger.WriteEntry("problem in convert barcode to flex numbers: " + ex.Message);
                movetoDefQ(fullname, 1, "Packing Slip", ex.Message);
                Logger.Log(16, Logger.Operations.ArchiveDocument, Logger.Statuses.MovedToManual, -1, fullname, barcode, string.Empty, string.Empty);
                return "Barcode data mismatch " + barcode;
            }

            // using is required so the file will not be locked
            try
            {
                string logPath = logFilesPath.Replace("{ccc}", ps.Company);
                logger.WriteEntry("Creating PS log at " + logPath + "/" + ps.PackingSlipNo.Trim() + ".csv");
                using (FileStream fs = File.Create(logPath + "/" + ps.PackingSlipNo.Trim() + ".csv"))
                {
                }
            }
            catch (Exception ex)
            {
                // if the file already exists, it's OK
                logger.WriteEntry("PS log not created: " + ex.Message, System.Diagnostics.EventLogEntryType.Warning);
                return "PS log not created: " + ex.Message;
            }
            return "Success";

        }

        private void createOrd99Log(FlexOrder99 Ord99)
        {
            // using is required so the file will not be locked
            try
            {
                logFilesOrd99Path = doxParams["BAANOrd99Log"];
                string logPath = logFilesOrd99Path.Replace("{ccc}", Ord99.Company);
                logger.WriteEntry("Creating order 99 log at " + logFilesOrd99Path + "/" + Ord99.Order99No.Trim() + ".csv");
                using (FileStream fs = File.Create(logPath + "/" + Ord99.Order99No.Trim() + ".csv"))
                {
                }
            }
            catch (Exception ex)
            {
                // if the file already exists, it's OK
                logger.WriteEntry("order 99 log not created: " + ex.Message, System.Diagnostics.EventLogEntryType.Warning);
            }

        }

        public void movetoDefQ(string filename, int reason, string kindOfDoc, string exMessage = "")
        {
            logger.WriteEntry("Packing slip file\\order 99\\ Supplier Invoice " + filename + " moved to manual queue", System.Diagnostics.EventLogEntryType.Warning);
            try
            {
                logger.WriteEntry("move " + filename + " to " + doxParams["DefectedQueuePath"] + "\\" + Path.GetFileName(filename), System.Diagnostics.EventLogEntryType.Information);
                System.IO.File.Move(filename, doxParams["DefectedQueuePath"] + "\\" + Path.GetFileName(filename));
                LogArchive(filename, kindOfDoc, 0, reason, exMessage);
            }
            catch (Exception ex)
            {
                logger.WriteEntry("Failed to move to manual queue with original fileName(retry with differant name) with this error: " + ex.Message, System.Diagnostics.EventLogEntryType.Error);

            }
            string newFileName = "";
            try
            {
                newFileName = Path.GetFileNameWithoutExtension(filename) + "_" + DateTime.Now.ToString("yyyy-dd-M--HH-mm-ss") + Path.GetExtension(filename);
                System.IO.File.Move(filename, doxParams["DefectedQueuePath"] + "\\" + newFileName);
            }
            catch (Exception ex)
            {
                logger.WriteEntry("Failed 2ND TIME to move to manual queue with original fileName" + newFileName + " with this error: " + ex.Message, System.Diagnostics.EventLogEntryType.Error);

            }
        }

        private string findInvoicefor(string company, string psno)
        {
            string invNumbers = "";
            // The following lines perform a search on DOX-Pro entities for a specific document.
            // 1. Define search fileds array by number of search criterions
            DOXAPI.SearchField[] fields = new DOXAPI.SearchField[2];
            // 2. For each one, create a search DocType
            DOXAPI.SearchField hasPS = new DOXAPI.SearchField();
            // 3. Set its search propeorties
            hasPS.FieldName = "Packing Slips";
            hasPS.SearchType = DOXAPI.SearchTypes.Partial;
            hasPS.FieldValue = psno;
            // 4. Put in search fields array
            fields[0] = hasPS;

            // 2.
            DOXAPI.SearchField inCompany = new DOXAPI.SearchField();
            // 3.
            inCompany.FieldName = "Invoice No";
            inCompany.SearchType = DOXAPI.SearchTypes.StartWith;
            inCompany.FieldValue = company;
            // 4.
            fields[1] = inCompany;

            // 5. Perform search
            DOXAPI.TreeItemWithDocType[] invoices = dox.FindTreeItemWithDocType(token, fields, flexInvoiceType.DocTypeId);
            foreach (DOXAPI.TreeItemWithDocType ti in invoices)
            {
                // Fetch entity from result set
                DOXAPI.TreeItemWithDocType inv = dox.GetTreeItemWithDocType(token, ti);
                // Read a DocType from entity
                object theInv = DOXFields.GetField(inv, "Invoice No");
                if (theInv != null) invNumbers += theInv.ToString();

            }
            return invNumbers.TrimEnd(',');
        }
        private void FindInvoice(string company, string inv_no)
        {
            doxLogin();
            DOXAPI.SearchField[] fields = new DOXAPI.SearchField[1];

            DOXAPI.SearchField field = new DOXAPI.SearchField();
            field.FieldName = "Invoice No";
            field.SearchType = DOXAPI.SearchTypes.StartWith;
            field.FieldValue = company + inv_no;
            fields[0] = field;

            DOXAPI.TreeItemWithDocType[] invoices = dox.FindTreeItemWithDocType(token, fields, flexInvoiceType.DocTypeId);
            foreach (DOXAPI.TreeItemWithDocType ti in invoices)
            {
                // Fetch entity from result set
                DOXAPI.TreeItemWithDocType inv = dox.GetTreeItemWithDocType(token, ti);
                string url = dox.GetDocumentURL(token, inv.ID);
                Console.WriteLine("ID: {0}, url \"{1}\"", inv.ID, url);
            }

        }
        private string updateOrCreateBinder(FlexCustomer cst)
        {
            try
            {
                // Try to fetch an entity from DOX-Pro
                DOXAPI.TreeItemWithDocType customerBinder = dox.GetTreeItemWithDocType(token, cst.asFetchItem());
                // If it doesn't exist...
                if (customerBinder == null)
                {
                    logger.WriteEntry("Creating binder for " + cst.ClientID);
                    // Create a new binder on DOX-Pro
                    response = dox.CreateBinder(token, cst.asBinder(), "Customers\\" + cst.Company, flexClientBinderType.DividerSets[0]);
                    // If the call failed...
                    if (!long.TryParse(response, out binderID))
                    {
                        logger.WriteEntry("Error creating binder for " + cst.FullID + "\n" + response);
                        if (response.IndexOf("Not logged in") > 0) // The token has expired
                        {
                            doxLogin();
                            response = dox.CreateBinder(token, cst.asBinder(), "Customers\\" + cst.Company, flexClientBinderType.DividerSets[0]);
                            if (!long.TryParse(response, out binderID))
                            {
                                Logger.Log((int)cst.asBinder().DocType.ID, Logger.Operations.CreateBinder, Logger.Statuses.Error, -1, string.Empty, string.Empty, cst.ClientName, response);
                                return "Error creating binder for " + cst.FullID + "\n" + response;
                            }
                        }
                    }
                    Logger.Log((int)cst.asBinder().DocType.ID, Logger.Operations.CreateBinder, Logger.Statuses.OK, (int)binderID, string.Empty, string.Empty, cst.ClientName, string.Empty);
                }
                else
                {
                    using (OdbcConnection DbConnection = new OdbcConnection(doxParams["BAANDB"]))
                    {
                        openConnection(DbConnection);
                        // Read data from BAAN DB
                        cst.refreshCustomerDetails(DbConnection, doxParams);

                        // Update the binder entity
                        response = cst.updateCustomerFields(customerBinder);
                        if (response != "")
                        {
                            return response;
                        }
                        // Save updates to DOX-Pro
                        response = dox.UpdateTreeItemWithDocType(token, customerBinder);
                        if (response != "Item updated." && response != String.Empty)
                        {
                            Logger.Log((int)cst.asBinder().DocType.ID, Logger.Operations.CreateBinder, Logger.Statuses.Error, (int)customerBinder.ID, string.Empty, string.Empty, customerBinder.Title, response);
                            return "Error updating binder for " + cst.FullID + "\n" + response;
                        }
                        Logger.Log((int)customerBinder.DocType.ID, Logger.Operations.CreateBinder, Logger.Statuses.OK, (int)customerBinder.ID, string.Empty, string.Empty, customerBinder.Title, string.Empty);
                    }
                }
                return "";
            }
            catch (Exception ex)
            {
                Logger.Log((int)cst.asBinder().DocType.ID, Logger.Operations.CreateBinder, Logger.Statuses.Error, -1, string.Empty, string.Empty, string.Empty, response);
                logger.WriteEntry("Error in updateorcreatebinder " + ex.Message);
                return "Error in updateorcreatebinder " + ex.Message;
            }
        }

        private void updatePackSlipsInvoiceNo(string[] slips, string company, int invoiceNo)
        {
            logger.WriteEntry("looking for " + slips.Length + " ps numbers");
            foreach (string psno in slips)
            {
                // The following lines perform a search on DOX-Pro entities for a specific document.
                // 1. Define search fileds array by number of search criterions
                DOXAPI.SearchField[] fields = new DOXAPI.SearchField[1];
                // 2. For each one, create a search DocType
                DOXAPI.SearchField psnoField = new DOXAPI.SearchField();
                // 3. Set its search propeorties
                psnoField.FieldName = "Packing Slip No";
                psnoField.SearchType = DOXAPI.SearchTypes.Exact;
                psnoField.FieldValue = company + psno;
                // 4. Put in search fields array
                fields[0] = psnoField;
                logger.WriteEntry("looking for " + company + psno);
                // 5. Perform search
                DOXAPI.TreeItemWithDocType[] tis = dox.FindTreeItemWithDocType(token, fields, flexPackSlipType.DocTypeId);
                if (tis.Length == 0)
                {
                    logger.WriteEntry("Packing slip " + company + "/" + psno + " not found");
                    return;
                }

                // Fetch entity from search results
                DOXAPI.TreeItemWithDocType ti = dox.GetTreeItemWithDocType(token, tis[0]);
                // Update a DocType in the entity
                DOXFields.SetField(ti, "Invoice No", company + invoiceNo);
                // Store back in DOX-Pro
                response = dox.UpdateTreeItemWithDocType(token, ti);

                if (response != "Item updated." && response != String.Empty)
                    logger.WriteEntry("Packing slip " + company + "/" + psno + " could not be updated with the invoice number\n" + response, System.Diagnostics.EventLogEntryType.Warning);
            }
        }
        public string handleOrder99(string fullName)
        {
            try
            {
                //get parameres from file name
                string FileName = Path.GetFileNameWithoutExtension(fullName);
                FlexOrder99 Ord99;
                try
                {
                    //now open an empty order99 document
                    Ord99 = new FlexOrder99(flexOrder99Type, flexOrder99_Order99No);
                    Ord99.Company = FileName.Substring(0, 3);
                    Ord99.SupplierID = FileName.Substring(4, 6);
                    Ord99.Order99No = FileName.Substring(10, 6);
                    Ord99.Filename = doxParams["Ord99EmptyPDFDocument"];//empty document becuase we dont have the returned document yet

                }
                catch (Exception e)
                {

                    //TODO remove marks here
                    movetoDefQ(fullName, 6, "Order99", e.Message);
                    Logger.Log(16, Logger.Operations.ArchiveDocument, Logger.Statuses.MovedToManual, -1, fullName, FileName, string.Empty, string.Empty);
                    logger.WriteEntry("FileName mismatch in Order99: " + FileName + ", Error: " + e.ToString());
                    return "FileName mismatch in Order99  " + FileName;
                }
                //load supplier
                FlexSupplier Supplier = new FlexSupplier(flexSupplierBinderType, flexSupplierBinder_SupplierID);
                Supplier.SupplierNo = Ord99.SupplierID;
                Supplier.CompanyNo = Ord99.Company;
                try
                {
                    using (OdbcConnection DbConnection = new OdbcConnection(doxParams["BAANDB"]))
                    {
                        openConnection(DbConnection);

                        OdbcCommand DbCommand = DbConnection.CreateCommand();
                        DbCommand.CommandTimeout = 3600;
                        OdbcDataReader DbReader;
                        DbCommand.CommandText = String.Format("select b020.t_nama,b020.t_namc, b020.t_refs, b020.t_telp, b020.t_seak,b040.t_email  from " +
                              "  "
                            + " baandb.ttccom020{0} as b020  full outer join baandb.ttccom040{0} as b040 on  b040.t_suno=b020.t_suno" +
                              " where   b040.t_actv=1  and b040.t_suno like '%{1}%'", "400", Supplier.SupplierNo);

                        try
                        {
                            DbReader = DbCommand.ExecuteReader();
                        }
                        catch (Exception ex)
                        {
                            return "Error " + ex.Message;
                        }

                        if (DbReader.Read())
                        {

                            Supplier.CompanyNo = "400";
                            Supplier.SupplierName = DbReader.GetString(0).Trim();
                            Supplier.SupplierAddress = DbReader.GetString(1).Trim();
                            Supplier.PurchasingPerson = DbReader.GetString(2).Trim();
                            Supplier.SupplierPhone = DbReader.GetString(3).Trim();
                            Supplier.SearchKey = DbReader.GetString(4).Trim();
                            Supplier.SupplierEmail = DbReader.GetString(5).Trim();
                            DbReader.Close();
                        }
                    }
                }
                catch (Exception ex)
                {

                }


                //check if order99 exist - if not create it
                // Try to fetch an entity from DOX-Pro            
                DOXAPI.TreeItemWithDocType Order99Binder = dox.GetTreeItemWithDocType(token, Ord99.asFetchItem());
                DOXAPI.TreeItemWithDocType SuppBinder = dox.GetTreeItemWithDocType(token, Supplier.asFetchItem());
                if ((Order99Binder != null) && (SuppBinder != null)) //case order with the supplier allready exists
                {
                    logger.WriteEntry("order " + Ord99.Order99No + " with supplier: " + Ord99.SupplierID + " allready exists");
                    File.Delete(fullName);
                    return "";
                }
                else
                {
                    updateOrCreateSuppBinder(Supplier);
                    // Get Supplier binder and Order 99 document as DOX-Pro objects
                    DOXAPI.Binder supplierBinder = Supplier.asIDBinder();
                    DOXAPI.Document docOrd99 = Ord99.asDocument();

                    response = dox.Archive(token, docOrd99, supplierBinder, "Order 99", false);
                    logger.WriteEntry("after Archive order99 (empty ord99) this is the response: " + response);

                    long docID;
                    if (long.TryParse(response, out docID))
                    {
                        System.IO.File.Delete(fullName);
                        logger.WriteEntry("Order 99 no. " + Ord99.Order99No + " archived with ID=" + docID);
                        LogArchive(fullName, "Order99", 1, 7);
                        Logger.Log((int)docOrd99.DocType.ID, Logger.Operations.ArchiveDocument, Logger.Statuses.OK, (int)docID, fullName, FileName, docOrd99.Title, string.Empty);
                    }
                    else
                    {
                        movetoDefQ(fullName, 3, "Order99", response);
                        Logger.Log(16, Logger.Operations.ArchiveDocument, Logger.Statuses.MovedToManual, -1, fullName, FileName, string.Empty, response);
                        return "order99 No. " + Ord99.Order99No + " could not be archived: " + response
                                    + ".\nFile moved to default queue";
                    }

                }
                return "";
            }
            catch (Exception ex)
            {
                Logger.Log(16, Logger.Operations.ArchiveDocument, Logger.Statuses.Error, -1, fullName, string.Empty, string.Empty, ex.Message);
                logger.WriteEntry("Exception in handle Order 99(Not the Returned Document) " + ex.Message);
                movetoDefQ(fullName, 3, "Order99", ex.Message);
                return "Exception in handle Order 99(Not the Returned Document):  " + ex.Message;
            }

        }
        public void updateSupplier(string id_supplier)
        {
            FlexSupplier Supplier = new FlexSupplier(flexSupplierBinderType, flexSupplierBinder_SupplierID);
            Supplier.SupplierNo = id_supplier;

            try
            {
                using (OdbcConnection DbConnection = new OdbcConnection(doxParams["BAANDB"]))
                {
                    openConnection(DbConnection);

                    OdbcCommand DbCommand = DbConnection.CreateCommand();
                    DbCommand.CommandTimeout = 3600;
                    OdbcDataReader DbReader = null;
                    DbCommand.CommandText = String.Format("select b020.t_nama,b020.t_namc, b020.t_refs, b020.t_telp, b020.t_seak,b040.t_email  from " +
                          "  "
                        + " baandb.ttccom020{0} as b020  full outer join baandb.ttccom040{0} as b040 on  b040.t_suno=b020.t_suno" +
                          " where   b040.t_actv=1  and b040.t_suno like '%{1}%'", "400", Supplier.SupplierNo);

                    try
                    {
                        DbReader = DbCommand.ExecuteReader();
                    }
                    catch (Exception ex)
                    {
                        logger.WriteEntry("Error " + ex.Message);
                    }

                    if (DbReader.Read())
                    {

                        Supplier.CompanyNo = "400";
                        Supplier.SupplierName = DbReader.GetString(0).Trim();
                        Supplier.SupplierAddress = DbReader.GetString(1).Trim();
                        Supplier.PurchasingPerson = DbReader.GetString(2).Trim();
                        Supplier.SupplierPhone = DbReader.GetString(3).Trim();
                        Supplier.SearchKey = DbReader.GetString(4).Trim();
                        Supplier.SupplierEmail = DbReader.GetString(5).Trim();
                        DbReader.Close();
                        logger.WriteEntry("UPDATE SUPPLIER Supplier - " + Supplier.SupplierName);
                        response = updateOrCreateSuppBinder(Supplier);
                        logger.WriteEntry("response - " + response);
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }
        public string handleOrder99ReturedDocument(string fullname, string barcode)
        {
            string error = "";
            // A new order99 is copied to a given location on the server by a scanner or other means.
            // The file is a PDF, and it has a barcode that contains key fields. This routine uses the OCR 
            // module to read the barcode and generate an archiving transaction to DOX-Pro.
            try
            {
                logger.WriteEntry("Handling Order99 Returned document: " + fullname);
                string filename = Path.GetFileNameWithoutExtension(fullname);
                doxLogin();
                // string barcode;

                // The order99 calls simulates DOX-Pro fields for the order99 doc-type
                FlexOrder99 Ord99 = new FlexOrder99(flexOrder99Type, flexOrder99_Order99No);
                if (barcode == "")
                {
                    try
                    {
                        barcode = getBarcodes(fullname, true, true);
                    }

                    catch (Exception e)
                    {
                        barcode = "Error";
                        error = " : " + e.Message
                            + "  " + Environment.NewLine + e.ToString();
                    }

                    if (barcode.IndexOf("Error") > -1 || barcode.Length < 15)
                    {
                        barcode = getBarcodesClearImage(fullname);
                        if (barcode.Equals("Error") || barcode.Length < 15)
                        {
                            // If failed to read barcode, the PDF is moved to a location of a DOX-Pro queue
                            movetoDefQ(fullname, 1, "Order99", error);
                            return "Error while reading barcode " + error + " - the error";
                        }

                    }
                }
                try
                {// -BARCODE FOR Supplier
                    Ord99.Company = barcode.Substring(0, 3);
                    Ord99.SupplierID = barcode.Substring(6, 4);
                    Ord99.Order99No = barcode.Substring(10, 6).Trim();
                    Ord99.SupplierID = Ord99.SupplierID.TrimStart('0');//remove leading zeros
                    Ord99.SupplierID = Ord99.SupplierID.Trim();
                    logger.WriteEntry("Barcode:" + barcode + " Order99No:" + Ord99.Order99No + " supplier:" + Ord99.SupplierID + "  Company:" + Ord99.Company);

                }
                catch (Exception e)
                {
                    // If barcode reading yielded wrong data, the PDF is moved to manual archiving
                    movetoDefQ(fullname, 1, "Order99", e.Message);
                    Logger.Log(16, Logger.Operations.ArchiveDocument, Logger.Statuses.MovedToManual, -1, fullname, barcode, string.Empty, string.Empty);
                    return "Barcode data mismatch in Order99 Returned Document " + barcode;
                }
                Ord99.Filename = fullname;
                // This object simulates the supplier binder in DOX-Pro
                FlexSupplier supp = new FlexSupplier(flexSupplierBinderType, flexSupplierBinder_SupplierID);
                //TODO - Fix Suuplier Number

                supp.SupplierNo = Ord99.SupplierID;
                supp.CompanyNo = Ord99.Company;


                try
                {
                    using (OdbcConnection DbConnection = new OdbcConnection(doxParams["BAANDB"]))
                    {
                        openConnection(DbConnection);

                        OdbcCommand DbCommand = DbConnection.CreateCommand();
                        DbCommand.CommandTimeout = 3600;
                        OdbcDataReader DbReader;
                        DbCommand.CommandText = String.Format("select b020.t_nama,b020.t_namc, b020.t_refs, b020.t_telp, b020.t_seak,b040.t_email  from " +
                              "  "
                            + " baandb.ttccom020{0} as b020  full outer join baandb.ttccom040{0} as b040 on  b040.t_suno=b020.t_suno" +
                              " where   b040.t_actv=1  and b040.t_suno like '%{1}%'", "400", supp.SupplierNo);

                        try
                        {
                            DbReader = DbCommand.ExecuteReader();
                        }
                        catch (Exception ex)
                        {
                            return "Error " + ex.Message;
                        }

                        if (DbReader.Read())
                        {

                            supp.CompanyNo = "400";
                            supp.SupplierName = DbReader.GetString(0).Trim();
                            supp.SupplierAddress = DbReader.GetString(1).Trim();
                            supp.PurchasingPerson = DbReader.GetString(2).Trim();
                            supp.SupplierPhone = DbReader.GetString(3).Trim();
                            supp.SearchKey = DbReader.GetString(4).Trim();
                            supp.SupplierEmail = DbReader.GetString(5).Trim();
                            DbReader.Close();
                        }
                    }
                }
                catch (Exception ex)
                {

                }










                logger.WriteEntry("Barcode:" + barcode + " c:" + Ord99.Order99No + " supplier:" + Ord99.SupplierID + " Order99No:" + Ord99.Order99No + " Company:" + Ord99.Company);
                string response;

                // If there is no supp binder yet (first order99 for this supplier), then create
                // a new binder on DOX-Pro. For existing supplier, some details may be updated frm BAAN.

                response = updateOrCreateSuppBinder(supp);

                logger.WriteEntry("before deleteing org document");
                //we have the org order 99 (not the returened document but the org from db with empty document - if we do have it we'll delete it because it contains an empty document - and refile the scanned document
                DOXAPI.TreeItemWithDocType Order99Org = dox.GetTreeItemWithDocType(token, Ord99.asFetchItem());
                if (Order99Org != null)
                {
                    dox.DeleteDocument(token, Order99Org.ID);
                }

                // Get supplier binder and packing slip document as DOX-Pro objects
                DOXAPI.Binder supplierBinder = supp.asIDBinder();
                DOXAPI.Document docOrd99 = Ord99.asDocument();


                // Archive a new order99 in DOX-Pro
                response = dox.Archive(token, docOrd99, supplierBinder, "Order 99", false);
                long docID;
                if (long.TryParse(response, out docID))
                {
                    createOrd99Log(Ord99);
                    System.IO.File.Delete(fullname);
                    logger.WriteEntry("Order 99 Returned Document no. " + Ord99.Order99No + " archived with ID=" + docID);
                    LogArchive(fullname, "Order99", 1, 7);
                    Logger.Log((int)docOrd99.DocType.ID, Logger.Operations.ArchiveDocument, Logger.Statuses.OK, (int)docID, fullname, barcode, docOrd99.Title, string.Empty);
                }
                else
                {

                    movetoDefQ(fullname, 3, "Order99", response);
                    Logger.Log(16, Logger.Operations.ArchiveDocument, Logger.Statuses.MovedToManual, -1, fullname, barcode, string.Empty, response);
                    return "Order 99 Returned Document no.  " + Ord99.Order99No + " could not be archived: " + response
                                + ".\nFile moved to default queue";
                }
                return "";
            }
            catch (Exception ex)
            {
                Logger.Log(16, Logger.Operations.ArchiveDocument, Logger.Statuses.Error, -1, fullname, string.Empty, string.Empty, ex.Message);
                logger.WriteEntry("Exception in handle order99 " + ex.Message);
                movetoDefQ(fullname, 3, "Order99", ex.Message);
                return "Exception while handling order99 " + ex.Message;
            }
        }

        private string getBarcodesClearImage(string fullname)
        {
            List<string> barcodesShip = new List<string>();
            try
            {
                BarcodeReader reader = new BarcodeReader();
                reader.Code39 = true; reader.Code128 = true;
                Inlite.ClearImageNet.Barcode[] barcodes = reader.Read(fullname);

                if (barcodes == null || barcodes.Length == 0)
                    barcodesShip.Add("Error");
                else
                {
                    foreach (Inlite.ClearImageNet.Barcode barcode in barcodes)
                    {
                        barcodesShip.Add(barcode.Text);
                    }

                }
                if (barcodesShip.Count == 0)
                    barcodesShip.Add("Error");
                logger.WriteEntry("In getBarcodesClearImage - the barcode :" + barcodesShip[0]);
                return barcodesShip[0];
            }
            catch (Exception ex)
            {
                barcodesShip.Add("Error");
                logger.WriteEntry("In getBarcodesClearImage - the barcode :" + barcodesShip[0]);
                return barcodesShip[0];
            }

        }

        //public string handleShipmentDoc(string fullname)
        //{
        //    // A new supplier invoice is copied to a given location on the server by a scanner or other means.
        //    // The file is a PDF, and it has a barcode that contains key fields. This routine uses the OCR 
        //    // module to read the barcode and generate an archiving transaction to DOX-Pro.
        //    try
        //    {
        //        logger.WriteEntry("Handling Supplier Shipment Doc NoneStorenext: " + fullname);
        //        string filename = Path.GetFileNameWithoutExtension(fullname);
        //        doxLogin();


        //        // The order99 calls simulates DOX-Pro fields for the order99 doc-type
        //        SupplierShipmentDoc SuppShipDoc = new SupplierShipmentDoc(ShipmentDocType, ShipmentDoc_ShipmentDocNo);
        //        SuppShipDoc.Filename = fullname;
        //        string barcode = string.Empty;
        //        try
        //        {////LIATLIATLIAT -BARCODE FOR Supplier
        //            SuppShipDoc.Company = "400";
        //            SuppShipDoc.SupplierID = "2134";
        //            SuppShipDoc.LotNo = "5548";// barcode.Substring(4, 4);
        //        }
        //        catch (Exception e)
        //        {
        //            // If barcode reading yielded wrong data, the PDF is moved to manual archiving
        //            movetoDefQ(fullname, 6, ".", e.Message);
        //            Logger.Log(16, Logger.Operations.ArchiveDocument, Logger.Statuses.MovedToManual, -1, fullname, barcode, string.Empty, string.Empty);
        //            return "Barcode data mismatch in not Storenext Supplier Returned Document " + barcode;
        //        }
        //        using (OdbcConnection DbConnection = new OdbcConnection(doxParams["BAANDB"]))
        //        {
        //            openConnection(DbConnection);
        //            // Read data from BAAN DB
        //            // SuppShipDoc.ShippmentNo = SuppShipDoc.GetShippmentNo(DbConnection,SuppShipDoc.Company, SuppShipDoc.LotNo, SuppShipDoc.Makat);
        //        }
        //        if (SuppShipDoc.ShippmentNo == "0")
        //        {

        //            // If barcode reading yielded wrong data, the PDF is moved to manual archiving
        //            movetoDefQ(fullname, 4, "Shipment");
        //            Logger.Log(16, Logger.Operations.ArchiveDocument, Logger.Statuses.MovedToManual, -1, fullname, barcode, string.Empty, string.Empty);
        //            return "couldnt find shipment number to makat " + SuppShipDoc.Makat + " and lot: " + SuppShipDoc.LotNo;
        //        }
        //        // This object simulates the supplier binder in DOX-Pro
        //        FlexSupplier supp = new FlexSupplier(flexSupplierBinderType, flexSupplierBinder_SupplierID);
        //        //TODO - Fix Suuplier Number
        //        supp.SupplierNo = SuppShipDoc.SupplierID;
        //        supp.CompanyNo = SuppShipDoc.Company;


        //        logger.WriteEntry(" supplier:" + SuppShipDoc.SupplierID + " ShippmentNo:" + SuppShipDoc.ShippmentNo);
        //        string response;

        //        // If there is no supp binder yet (first order99 for this supplier), then create
        //        // a new binder on DOX-Pro. For existing supplier, some details may be updated frm BAAN.

        //        response = updateOrCreateSuppBinder(supp);

        //        // Get supplier binder and packing slip document as DOX-Pro objects
        //        DOXAPI.Binder supplierBinder = supp.asIDBinder();
        //        DOXAPI.Document docSuppShip = SuppShipDoc.asDocument();


        //        // Archive a new order99 in DOX-Pro
        //        response = dox.Archive(token, docSuppShip, supplierBinder, "Shipments", false);
        //        long docID;
        //        if (long.TryParse(response, out docID))
        //        {
        //            System.IO.File.Delete(fullname);
        //            LogArchive(fullname, "Shipments", 1, 7);
        //            logger.WriteEntry("Shipment Doc Document no. " + SuppShipDoc.ShippmentNo + " archived with ID=" + docID);
        //            Logger.Log((int)docSuppShip.DocType.ID, Logger.Operations.ArchiveDocument, Logger.Statuses.OK, (int)docID, fullname, barcode, docSuppShip.Title, string.Empty);
        //        }
        //        else
        //        {

        //            movetoDefQ(fullname, 3, "Shipments", response);
        //            Logger.Log(16, Logger.Operations.ArchiveDocument, Logger.Statuses.MovedToManual, -1, fullname, barcode, string.Empty, response);
        //            return "Order 99 Returned Document no.  " + SuppShipDoc.ShippmentNo + " could not be archived: " + response
        //                        + ".\nFile moved to default queue";
        //        }
        //        return "";
        //    }
        //    catch (Exception ex)
        //    {
        //        Logger.Log(16, Logger.Operations.ArchiveDocument, Logger.Statuses.Error, -1, fullname, string.Empty, string.Empty, ex.Message);
        //        logger.WriteEntry("Exception in handle order99 " + ex.Message);
        //        movetoDefQ(fullname, 3, "order99", ex.Message);
        //        return "Exception while handling order99 " + ex.Message;
        //    }
        //}
       
        
        public void movetoDefQShipment(string filename, int reason, string kindOfDoc, string renameBegining, string exMessage = "")
        {
            logger.WriteEntry("Shipment " + filename + " moved to manual queue", System.Diagnostics.EventLogEntryType.Warning);
            try
            {
                logger.WriteEntry("move " + filename + " to " + @"\\mignt080\media\DoxArchive\scaner\envFlex\MachsanErrors\" + Path.GetFileName(filename), System.Diagnostics.EventLogEntryType.Information);
                string renameFile = "";
                if (!filename.Contains(renameBegining))
                    renameFile = filename.Substring(0, filename.LastIndexOf('\\')) + "\\" + renameBegining + filename.Remove(0, filename.LastIndexOf('\\') + 1);
                else
                    renameFile = filename.Substring(0, filename.LastIndexOf('\\')) + "\\" + filename.Remove(0, filename.LastIndexOf('\\') + 1);

                System.IO.File.Move(filename, renameFile);
                
                    System.IO.File.Move(renameFile, @"\\mignt080\media\DoxArchive\scaner\envFlex\MachsanErrors\" + Path.GetFileName(renameFile));
                LogArchive(filename, kindOfDoc, 0, reason, exMessage);//
            }
            catch (Exception ex)
            {
                try
                {
                    string renameFile = "";
                    if (!filename.Contains(renameBegining))
                        renameFile = filename.Substring(0, filename.LastIndexOf('\\')) + "\\" + renameBegining + "_" + DateTime.Now.ToString("yyyy-dd-M--HH-mm-ss") + filename.Remove(0, filename.LastIndexOf('\\') + 1);
                    else
                        renameFile = filename.Substring(0, filename.LastIndexOf('\\')) + "\\" + "_" + DateTime.Now.ToString("yyyy-dd-M--HH-mm-ss") + filename.Remove(0, filename.LastIndexOf('\\') + 1);

                    System.IO.File.Move(filename, renameFile);
                    
                        System.IO.File.Move(renameFile, @"\\mignt080\media\DoxArchive\scaner\envFlex\MachsanErrors\" + Path.GetFileName(renameFile));

                }
                catch { }
                logger.WriteEntry("Failed to move to manual queue with original fileName(retry with differant name) with this error: " + ex.Message, System.Diagnostics.EventLogEntryType.Error);

            }
            string newFileName = "";
            try
            {
                string renameFile = "";
                if (!filename.Contains(renameBegining))
                    renameFile = filename.Substring(0, filename.LastIndexOf('\\')) + "\\" + renameBegining + filename.Remove(0, filename.LastIndexOf('\\') + 1);
                else
                    renameFile = filename.Substring(0, filename.LastIndexOf('\\')) + "\\" + filename.Remove(0, filename.LastIndexOf('\\') + 1);
                if (File.Exists(renameFile))
                {
                    File.Delete(renameFile);
                    System.IO.File.Move(filename, renameFile);//
                    newFileName = Path.GetFileNameWithoutExtension(renameFile) + "_" + DateTime.Now.ToString("yyyy-dd-M--HH-mm-ss") + Path.GetExtension(renameFile);
                    
                        System.IO.File.Move(renameFile, @"\\mignt080\media\DoxArchive\scaner\envFlex\MachsanErrors\" + newFileName);//
                }

            }
            catch (Exception ex)
            {
                logger.WriteEntry("Failed 2ND TIME to move to manual queue with original fileName" + newFileName + " with this error: " + ex.Message, System.Diagnostics.EventLogEntryType.Error);

            }
        }
        public string handleShipmentDoc(string fullname, string filenameWithoutExt, string typeFile)
        {

            string current_page = "";
            bool last = false;
            int i = 1, j = 1;
            bool isNotFirst = false;
            string barcodeString = "Error";
            int returnGood = 0, returnBad = 0;
            try
            {

                List<string> barcodes;
                List<string> barcodes2;
                pdfSplit pdfs = new pdfSplit();
                string ppath = fullname;
                PdfReader pdfReader = new PdfReader(ppath);
                int numberOfPages = pdfReader.NumberOfPages;
                pdfReader.Close();
                //while not finish to extract all pdf correct ship
                logger.WriteEntry("Handling Supplier Shipment Doc NoneStorenext: " + fullname);
                if (!Directory.Exists(Path.GetDirectoryName(fullname) + "\\splitsPdf"))
                    Directory.CreateDirectory(Path.GetDirectoryName(fullname) + "\\splitsPdf");
                if (!Directory.Exists(Path.GetDirectoryName(fullname) + "\\splitsPdfForGetBarcodes"))
                    Directory.CreateDirectory(Path.GetDirectoryName(fullname) + "\\splitsPdfForGetBarcodes");
                string filename = Path.GetFileNameWithoutExtension(fullname);
                string lastPage = "";

                while (i <= numberOfPages && !last)
                {
                    current_page = "";
                    isNotFirst = false;
                    barcodes2 = new List<string>();
                    doxLogin();
                    barcodeString = "Error";
                    barcodes2 = new List<string>();
                    while (barcodeString == "Error" && i <= numberOfPages)
                    {
                        string newFileNamei = Path.GetDirectoryName(fullname) + "\\splitsPdfForGetBarcodes\\" + filename + "_" + i + ".pdf";
                        logger.WriteEntry("newFileNamei - " + newFileNamei);
                        pdfs.ExtractPage(fullname, newFileNamei, i);
                        try
                        {
                            int count = 0;
                            barcodes = new List<string>();

                            barcodes = ExecuteCommandShip(newFileNamei, Path.GetFileNameWithoutExtension(newFileNamei));
                          
                            
                            foreach (var item in barcodes)
                            {
                                int flager = 0;
                                string s = item;
                                if (item.Contains("/"))
                                {
                                    s = item.Substring(0, item.IndexOf("/"));
                                    if (s.Length == 10)
                                        flager = 1;
                                }

                                long Letter;

                                if (long.TryParse(s, out Letter))
                                {
                                    if (((item.Length == 10 && s.Length == 10) || (flager == 1 && s.Length == 10)) && item.Substring(0, 2) == (DateTime.Today.Year % 100).ToString())
                                        count++;
                                }
                            }
                            if (!isNotFirst)//save the first time
                            {
                                barcodes2.AddRange(barcodes);
                                isNotFirst = true;
                            }
                            //barcodes.Count <= 1
                            else if (barcodes != null && barcodes.Count > 0 && barcodes[0] == "Error" || count == 0)
                            {
                                barcodeString = "Error";
                                if (barcodes[0] != "Error")
                                {
                                    barcodes2.AddRange(barcodes);
                                }
                            }
                            else
                            {
                                barcodeString = "";

                            }
                        }
                        catch (Exception ex)
                        {
                            logger.WriteEntry("Error in excute barcode - " + ex.Message);

                        }
                        lastPage = newFileNamei;
                        File.Delete(newFileNamei);
                        i++;
                    }
                    i -= 1;
                    try
                    {
                        string newFileName = "";
                        if ((barcodeString == ""))//all options
                        {
                            int ii = 0, jj = 0;
                            jj = j > i ? j - 1 : j;
                            ii = i - 1 < j ? i : i - 1;
                            newFileName = Path.GetDirectoryName(fullname) + "\\splitsPdf\\" + filename + "_" + jj + "_" + ii + ".pdf";

                            current_page = newFileName;
                            if (jj == ii)
                                pdfs.ExtractPage(fullname, newFileName, ii);
                            else
                                pdfs.ExtractPages(fullname, newFileName, jj, ii);


                        }
                        else if (barcodeString == "Error" || i == numberOfPages)
                        {
                            int jj = j > i ? j - 1 : j;
                            newFileName = Path.GetDirectoryName(fullname) + "\\splitsPdf\\" + filename + "_" + jj + "_" + numberOfPages + ".pdf";

                            current_page = newFileName;
                            pdfs.ExtractPages(fullname, newFileName, jj, numberOfPages);


                            last = true;
                        }
                        else
                        {
                            int ii = 0, jj = 0;
                            jj = j > i ? j - 1 : j;
                            ii = i - 1 < j ? i : i - 1;
                            newFileName = Path.GetDirectoryName(fullname) + "\\splitsPdf\\" + filename + "_" + jj + "_" + ii + ".pdf";
                            current_page = newFileName;
                            if (jj == ii)

                                pdfs.ExtractPage(fullname, newFileName, ii);
                            else
                                pdfs.ExtractPages(fullname, newFileName, jj, ii);
                        }

                        try
                        {

                            int count = 0;
                            StringBuilder builder = new StringBuilder();
                            builder.Append("");
                            foreach (string value in barcodes2)
                            {
                                int flager = 0;
                                string s = value;
                                if (value.Contains("/"))
                                {
                                    s = value.Substring(0, value.IndexOf("/"));
                                    if (s.Length == 10)
                                        flager = 1;
                                }

                                long Letter;

                                if (long.TryParse(s, out Letter))
                                {
                                    if (((value.Length == 10 && s.Length == 10) || (flager == 1 && s.Length == 10)) && value.Substring(0, 2) == (DateTime.Today.Year % 100).ToString())
                                        count++;
                                }
                                builder.Append(value);
                                builder.Append(',');
                            }
                            // && count == 1
                            if (builder.ToString() != "" && !builder.ToString().Contains("Error") && count >= 1)
                            {
                                logger.WriteEntry("going to  LotsAndMakats" + newFileName+" lots: "+( builder.ToString().EndsWith(",") ? builder.ToString().Substring(0, builder.ToString().Length - 1) : builder.ToString()));
                                string response=LotsAndMakats(newFileName, builder.ToString().EndsWith(",") ? builder.ToString().Substring(0, builder.ToString().Length - 1) : builder.ToString(), typeFile);
                                logger.WriteEntry("after  LotsAndMakats" + fullname);
                               // if (response.Equals(""))
                               // {
                                    using (IDataSupplier dataSupplier = DataManager.GetDataSupplier(DataManager.defaultType, "Data Source=localhost;user id=dmsdoxpassword=xodsmd123!;Initial Catalog=DoxPro_Env_Flex;Integrated Security=True"))
                                    {


                                        dataSupplier.OpenQuery();
                                        dataSupplier.AddParameter("FileName", newFileName);
                                        dataSupplier.AddParameter("lotsAndMakats", builder.ToString().EndsWith(",") ? builder.ToString().Substring(0, builder.ToString().Length - 1) : builder.ToString());
                                        dataSupplier.AddParameter("status", "0");
                                        dataSupplier.AddParameter("type", typeFile);
                                        dataSupplier.AddParameter("dateGetFile", DateTime.Now.ToString("MM-dd-yyyy HH:mm:ss"));
                                        dataSupplier.Execute("INSERT INTO [shipmentArchive]([FileName],[lotsAndMakats] ,[status],[type],[dateGetFile]) \r\n\t\t\t\t     VALUES ( @FileName,@lotsAndMakats,@status,@type,@dateGetFile)");

                                    }

                                    returnGood++;
                               // }
                                /*else
                                {
                                    using (IDataSupplier dataSupplier = DataManager.GetDataSupplier(DataManager.defaultType, Settings.Default.ConnectionString))
                                    {
                                        dataSupplier.OpenQuery();
                                        dataSupplier.AddParameter("FileName", newFileName);
                                        dataSupplier.AddParameter("lotsAndMakats", builder.ToString().EndsWith(",") ? builder.ToString().Substring(0, builder.ToString().Length - 1) : builder.ToString());
                                        dataSupplier.AddParameter("status", "-3");
                                        dataSupplier.AddParameter("type", typeFile);
                                        dataSupplier.AddParameter("dateGetFile", DateTime.Now.ToString("MM-dd-yyyy HH:mm:ss"));
                                        dataSupplier.Execute("INSERT INTO [shipmentArchive]([FileName],[lotsAndMakats] ,[status],[type],[dateGetFile]) \r\n\t\t\t\t     VALUES ( @FileName,@lotsAndMakats,@status,@type,@dateGetFile)");
                                    }
                                    returnBad++;
                                }*/
                            }
                            else
                            {
                                using (IDataSupplier dataSupplier = DataManager.GetDataSupplier(DataManager.defaultType, Settings.Default.ConnectionString))
                                {
                                    dataSupplier.OpenQuery();
                                    dataSupplier.AddParameter("FileName", newFileName);
                                    dataSupplier.AddParameter("lotsAndMakats", builder.ToString().EndsWith(",") ? builder.ToString().Substring(0, builder.ToString().Length - 1) : builder.ToString());
                                    dataSupplier.AddParameter("status", "-0");
                                    dataSupplier.AddParameter("type", typeFile);
                                    dataSupplier.AddParameter("dateGetFile", DateTime.Now.ToString("MM-dd-yyyy HH:mm:ss"));
                                    dataSupplier.Execute("INSERT INTO [shipmentArchive]([FileName],[lotsAndMakats] ,[status],[type],[dateGetFile]) \r\n\t\t\t\t     VALUES ( @FileName,@lotsAndMakats,@status,@type,@dateGetFile)");
                                }
                                returnBad++;
                                logger.WriteEntry("the barcode is invalid");
                                // If barcode reading yielded wrong data, the PDF is moved to manual archiving
                                string renameFile = typeFile.ToLower().Contains("COC".ToLower()) ? "COC_" : typeFile.ToLower().Contains("Invoice".ToLower()
                                    ) ? "Invoice_" : "DeliveryNote_";
                                LogArchiveShipment(newFileName, typeFile, 0, 1, "No Barcodes - Or Read Barcode = ERROR");
                                movetoDefQShipment(newFileName, 6, ".", renameFile, "the barcode is invalid");

                            }

                        }
                        catch (Exception ex)
                        {
                            logger.WriteEntry("Error in  insert to shipmentArchive- " + ex.Message);

                        }
                        current_page = newFileName;
                        if (File.Exists(lastPage))
                            File.Delete(lastPage);
                        logger.WriteEntry("lastPage" + lastPage);
                        j = i;

                        if (barcodes2.Count < 1)
                        {
                            if (newFileName == "")
                            {
                                int ii = 0, jj = 0;
                                jj = j > i ? j - 1 : j;
                                ii = i - 1 < j ? i : i - 1;
                                if (last)
                                {
                                    newFileName = Path.GetDirectoryName(fullname) + "\\splitsPdf\\" + filename + "_" + (j > i ? j - 1 : j) + "_" + numberOfPages + ".pdf";
                                    if (jj == ii)
                                        pdfs.ExtractPage(fullname, newFileName, ii);
                                    else
                                        pdfs.ExtractPages(fullname, newFileName, jj, numberOfPages);

                                }
                                else
                                {
                                    newFileName = Path.GetDirectoryName(fullname) + "\\splitsPdf\\" + filename + "_" + (j > i ? j - 1 : j) + "_" + numberOfPages + ".pdf";
                                    if (jj == ii)
                                        pdfs.ExtractPage(fullname, newFileName, ii);
                                    else
                                        pdfs.ExtractPages(fullname, newFileName, jj, numberOfPages);
                                }

                            }
                            logger.WriteEntry("The barcode is invalid");
                            string renameFile = typeFile.ToLower().Contains("COC".ToLower()) ? "COC_" : typeFile.ToLower().Contains("Invoice".ToLower()) ? "Invoice_" : "DeliveryNote_";
                            LogArchiveShipment(newFileName, typeFile, 0, 1, "The barcode is invalid");
                            movetoDefQShipment(newFileName, 6, ".", renameFile, "the barcode is invalid");
                            try
                            {
                                using (IDataSupplier dataSupplier = DataManager.GetDataSupplier(DataManager.defaultType, Settings.Default.ConnectionString))
                                {
                                    dataSupplier.OpenQuery();
                                    dataSupplier.AddParameter("FileName", newFileName);
                                    dataSupplier.AddParameter("lotsAndMakats", "");

                                    dataSupplier.AddParameter("status", "-0");
                                    dataSupplier.AddParameter("type", typeFile);
                                    dataSupplier.Execute("INSERT INTO [shipmentArchive]([FileName],[lotsAndMakats] ,[status],[type]) \r\n\t\t\t\t     VALUES ( @FileName,@lotsAndMakats,@status,@type)");
                                }
                                returnBad++;
                            }
                            catch (Exception ex)
                            {
                                logger.WriteEntry("Error in  insert to shipmentArchive- " + ex.Message);

                            }
                            j = i;
                            continue;

                        }
                        j = i;
                    }
                    catch (Exception e)
                    {
                        returnBad++;
                        logger.WriteEntry("exception - " + e.Message);
                        return returnGood + "_" + returnBad;
                        //string renameFile = typeFile.ToLower().Contains("COC".ToLower()) ? "COC_" : typeFile.ToLower().Contains("Invoice".ToLower()
                        //     ) ? "Invoice_" : "DeliveryNote_";
                        //LogArchiveShipment(current_page, typeFile, 0, 1, e.Message);
                        //movetoDefQShipment(current_page, 6, ".", renameFile, e.Message);
                        //Logger.Log(16, Logger.Operations.ArchiveDocument, Logger.Statuses.MovedToManual, -1, current_page, barcodes2[0], string.Empty, string.Empty);

                    }
                    j = i;


                }

            }


            catch (Exception ex)
            {


                // Add by batya 16.2.2017 damaged file
                // create directory for damaged file and put into all file (COC or invoice) that a PDFreader can't succses open- damaged.
                // the same level COC directory and invoices directory
              

                string result = "";
                string DirFile = "";
                string[] aaa = fullname.Split('\\');

                foreach (var item in aaa)
                {
                    if (item == aaa[aaa.Length - 2])
                        DirFile += item + "\\";
                    else
                        if (item != aaa[aaa.Length - 1])
                        {
                            result += item + "\\";
                            DirFile += item + "\\";
                        }
                }

                logger.WriteEntry("The url damaged file:" + fullname);

                if (!Directory.Exists((Path.GetDirectoryName(result) + "\\damaged")))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(result) + "\\damaged");
                    logger.WriteEntry("create damaged directory - sucssfull");
                }

                try
                {
                    System.IO.File.Copy(fullname, Path.GetDirectoryName(result) + "\\damaged\\" + Path.GetFileName(fullname));
                    System.IO.File.Delete(fullname);
                    logger.WriteEntry("move damaged file to damaged directory - sucssfull");
                }
                    
                  catch (Exception)

                    {
                        logger.WriteEntry("The file damaged not moved");   
                    }
                // end change batya
                



                Logger.Log(16, Logger.Operations.ArchiveDocument, Logger.Statuses.Error, -1, fullname, string.Empty, string.Empty, ex.Message);
                logger.WriteEntry("Exception in handle Shipment " + ex.Message);
                //string renameFile = typeFile.ToLower().Contains("COC".ToLower()) ? "COC_" : typeFile.ToLower().Contains("Invoice".ToLower()
                //                ) ? "Invoice_" : "DeliveryNote_";
                //LogArchiveShipment(fullname, typeFile, 0, 1, ex.Message);
                //movetoDefQShipment(fullname, 3, "Shipment", renameFile, ex.Message);
                //returnBad++;
                return returnGood + "_" + returnBad;
            }
            logger.WriteEntry("finish archive the shipment docs");
            if (!Directory.Exists((Path.GetDirectoryName(fullname) + "\\done")))
                Directory.CreateDirectory(Path.GetDirectoryName(fullname) + "\\done");
            if (File.Exists(Path.GetDirectoryName(fullname) + "\\done\\" + Path.GetFileName(fullname)))
                System.IO.File.Move(fullname, Path.GetDirectoryName(fullname) + "\\done\\" + Path.GetFileName(fullname).Replace(".pdf", "_" + DateTime.Now.ToString("yyyy-dd-M--HH-mm-ss") + ".pdf"));
            else
                System.IO.File.Move(fullname, Path.GetDirectoryName(fullname) + "\\done\\" + Path.GetFileName(fullname));

            return returnGood + "_" + returnBad;
        }
        //End Ayala 13.04.2015
        public string handleStorenextInvoiceDoc(string DataFileName, string PdfFileName, string Company)
        {
            //storenext has 2 shipment document - txt file and pdf. the txt file has data

            try
            {

                if (!File.Exists(PdfFileName))
                {
                    if (!File.Exists(doxParams["SuppInvErrQ"] + "\\" + Path.GetFileName(PdfFileName)))
                    {
                        System.IO.File.Copy(DataFileName, doxParams["SuppInvErrQ"] + "\\" + Path.GetFileName(DataFileName));
                    }
                    movetoSuppInvErrorQ(DataFileName);
                    Logger.Log(16, Logger.Operations.ArchiveDocument, Logger.Statuses.MovedToManual, -1, DataFileName, DataFileName, string.Empty, string.Empty);
                    return "Storenext Supplier Invoice Doc Pdf file  " + PdfFileName + "  DOES NOT exists ";

                }
                logger.WriteEntry("Handling Storenext Supplier Invoice Doc: " + DataFileName);
                doxLogin();
                // Read the file and display it line by line.
                string line;
                string[] Data;
                System.IO.StreamReader Datafile = new System.IO.StreamReader(DataFileName);
                //TODO remove // from moveToDefQ
                try
                {
                    line = Datafile.ReadLine();
                    Data = line.Split(',');
                }
                catch (Exception e)
                {
                    Datafile.Close();

                    movetoSuppInvErrorQ(DataFileName);
                    try
                    {
                        System.IO.File.Copy(PdfFileName, doxParams["SuppInvErrQ"] + "\\" + Path.GetFileName(PdfFileName));
                    }
                    catch (Exception)
                    {
                        //do nothing - file allreday exists
                    }
                    movetoDefQ(PdfFileName, 6, "invoice", e.Message);
                    Logger.Log(16, Logger.Operations.ArchiveDocument, Logger.Statuses.MovedToManual, -1, DataFileName, DataFileName, string.Empty, string.Empty);
                    return "Storenext Supplier Invoice Doc Error reading data from text file " + DataFileName + " with Error: " + e.ToString();
                }

                Datafile.Close();

                //preapre for archiving - get relevant data
                SupplierInvoice SuppInvDoc = new SupplierInvoice(SupplierInvoiceDocType, SupplierInvoiceDoc_ShipmentDocNo);

                string Temp_SupplierID = "";//in storenext file we have memi supplier we need to do db query to get flex supplier
                try
                //TODO fix locations here
                {
                    SuppInvDoc.Company = Company;
                    Temp_SupplierID = Data[2].Trim();//if exist shipment number archive it to shipment if no archive it to invoices
                    SuppInvDoc.InvoiceNo = Data[3].Trim();
                }
                catch (Exception e)
                {
                    movetoSuppInvErrorQ(DataFileName);
                    try
                    {
                        System.IO.File.Copy(PdfFileName, doxParams["SuppInvErrQ"] + "\\" + Path.GetFileName(PdfFileName));
                    }
                    catch (Exception)
                    {
                        //do nothing - file allreday exists
                    }
                    movetoDefQ(PdfFileName, 6, "invoice", e.Message);
                    Logger.Log(16, Logger.Operations.ArchiveDocument, Logger.Statuses.MovedToManual, -1, DataFileName, line, string.Empty, string.Empty);
                    return "Storenext Supplier Invoice Doc Error reading data from text file  " + DataFileName + " with Error: " + e.ToString(); ;
                }
                #region GetFlexSupplierID
                SuppInvDoc.SupplierID = "";
                try
                {
                    if (Temp_SupplierID != "")
                    {
                        using (OdbcConnection DbConnection = new OdbcConnection(doxParams["BAANDB"]))
                        {
                            openConnection(DbConnection);
                            SuppInvDoc.SupplierID = FindFlexStorenextSupplierID(Temp_SupplierID, SuppInvDoc.Company);
                        }
                    }
                    else
                    {
                        throw new Exception("Exception in Storenext Invoice Document - Supplier ID doesnt exist in storenext data file");
                    }
                    if (SuppInvDoc.SupplierID == "")
                        throw new Exception("Exception in Storenext Invoice Document - Flex Supplier doesnt exist in DB for Supplier Memi:" + Temp_SupplierID);
                }
                catch (Exception e)
                {
                    movetoSuppInvErrorQ(DataFileName);
                    try
                    {
                        System.IO.File.Copy(PdfFileName, doxParams["SuppInvErrQ"] + "\\" + Path.GetFileName(PdfFileName));
                    }
                    catch (Exception)
                    {
                        //do nothing - file allreday exists
                    }
                    movetoDefQ(PdfFileName, 6, "invoice");
                    Logger.Log(16, Logger.Operations.ArchiveDocument, Logger.Statuses.MovedToManual, -1, DataFileName, line, string.Empty, string.Empty);
                    return "Storenext Supplier Invoice Doc Excpetion Message: " + e.Message;
                }
                #endregion// GetFlexSupplierID


                // This object simulates the supplier binder in DOX-Pro
                FlexSupplier supp = new FlexSupplier(flexSupplierBinderType, flexSupplierBinder_SupplierID);
                //TODO - Fix Suuplier Number
                supp.SupplierNo = SuppInvDoc.SupplierID;
                supp.CompanyNo = SuppInvDoc.Company;


                logger.WriteEntry("SupplierNo:" + SuppInvDoc.SupplierID + " InvoiceNo:" + SuppInvDoc.InvoiceNo);
                string response;



                response = updateOrCreateSuppBinder(supp);
                // Get supplier binder and packing slip document as DOX-Pro objects
                DOXAPI.Binder supplierBinder = supp.asIDBinder();
                SuppInvDoc.Filename = PdfFileName;
                DOXAPI.Document docSuppInv = SuppInvDoc.asDocument();
                logger.WriteEntry("before Storenext supplier invoice Doc Archiving");
                // Archive a new order99 in DOX-Pro
                response = dox.Archive(token, docSuppInv, supplierBinder, "Invoices", false);
                long docID;
                if (long.TryParse(response, out docID))
                {
                    System.IO.File.Delete(PdfFileName);
                    System.IO.File.Delete(DataFileName);
                    LogArchive(PdfFileName, "Invoices", 1, 7);
                    logger.WriteEntry("Storenext Invoice Doc Document no. " + SuppInvDoc.InvoiceNo + " archived with ID=" + docID);
                    Logger.Log((int)docSuppInv.DocType.ID, Logger.Operations.ArchiveDocument, Logger.Statuses.OK, (int)docID, DataFileName, line, supplierBinder.Title, string.Empty);

                }
                else
                {

                    movetoSuppInvErrorQ(DataFileName);
                    try
                    {
                        System.IO.File.Copy(PdfFileName, doxParams["SuppInvErrQ"] + "\\" + Path.GetFileName(PdfFileName));
                    }
                    catch (Exception)
                    {
                        //do nothing - file allreday exists
                    }
                    movetoDefQ(PdfFileName, 3, "invoice");
                    Logger.Log(16, Logger.Operations.ArchiveDocument, Logger.Statuses.MovedToManual, -1, DataFileName, line, string.Empty, response);
                    return "Storenext Invoice Doc  no.  " + SuppInvDoc.InvoiceNo + " could not be archived: " + response
                                + ".\nFile moved to default queue";
                }
                return "";
            }
            catch (Exception ex)
            {
                Logger.Log(16, Logger.Operations.ArchiveDocument, Logger.Statuses.Error, -1, DataFileName, string.Empty, string.Empty, ex.Message);
                logger.WriteEntry("Exception in handle Storenext Supplier Invoice Doc " + ex.Message);
                movetoSuppInvErrorQ(DataFileName);
                try
                {
                    System.IO.File.Copy(PdfFileName, doxParams["SuppInvErrQ"] + "\\" + Path.GetFileName(PdfFileName));
                }
                catch (Exception)
                {
                    //do nothing - file allreday exists
                }
                movetoDefQ(PdfFileName, 6, "invoice", ex.Message);
                return "Exception while handling Storenext Supplier Invoice Doc " + ex.Message;
            }
        }
        private void movetoSuppInvErrorQ(string filename)
        {
            logger.WriteEntry("Storenext Supplier Invoice file " + filename + " moved to manual queue: " + doxParams["SuppInvErrQ"], System.Diagnostics.EventLogEntryType.Warning);
            if (!(File.Exists(doxParams["SuppInvErrQ"] + "\\" + Path.GetFileName(filename))))
            {
                System.IO.File.Move(filename, doxParams["SuppInvErrQ"] + "\\" + Path.GetFileName(filename));
            }
            else //allready exists we need just to delete
            {
                File.Delete(filename);
            }
        }
        private string updateOrCreateSuppBinder(FlexSupplier supp)
        {
            // Try to fetch an entity from DOX-Pro
            logger.WriteEntry("check if supplier exist " + supp.SupplierNo);
            DOXAPI.TreeItemWithDocType SuppBinder = dox.GetTreeItemWithDocType(token, supp.asFetchItem());
            // If it doesn't exist...
            if (SuppBinder == null)
            {
                CreateSupplierInDox(supp);
                logger.WriteEntry("supplier doesnt exist so Creating binder for " + supp.SupplierNo);
                // Create a new binder on DOX-Pro
                response = dox.CreateBinder(token, supp.asBinder(), "Suppliers", flexSupplierBinderType.DividerSets[0]);
                // If the call failed...
                logger.WriteEntry("token after create binder " + token);
                if (!long.TryParse(response, out binderID))
                {
                    logger.WriteEntry("Error creating binder for " + supp.SupplierName + "\n" + response);
                    if (response.IndexOf("Not logged in") > 0) // The token has expired
                    {
                        doxLogin();
                        response = dox.CreateBinder(token, supp.asBinder(), "Suppliers\\" + supp.SupplierName, flexSupplierBinderType.DividerSets[0]);
                        if (!long.TryParse(response, out binderID))
                        {
                            Logger.Log((int)supp.asBinder().DocType.ID, Logger.Operations.CreateBinder, Logger.Statuses.Error, -1, string.Empty, string.Empty, supp.SupplierName, response);
                            return "Error creating binder for " + supp.SupplierName + "\n" + response;
                        }
                    }
                }
                Logger.Log((int)supp.asBinder().DocType.ID, Logger.Operations.CreateBinder, Logger.Statuses.OK, (int)binderID, string.Empty, string.Empty, supp.SupplierName, string.Empty);
            }
            else
            {
                logger.WriteEntry("supplier exist now updating data " + supp.SupplierNo);
                using (OdbcConnection DbConnection = new OdbcConnection(doxParams["BAANDB"]))
                {
                    openConnection(DbConnection);
                    // Read data from BAAN DB                        
                    supp.GetSupplierDetails(DbConnection, doxParams);

                    // Update the binder entity
                    response = supp.updateSupplierFields(SuppBinder);
                    if (response != "")
                    {
                        return response;
                    }
                    // Save updates to DOX-Pro
                    response = dox.UpdateTreeItemWithDocType(token, SuppBinder);
                    if (response != "Item updated." && response != String.Empty)
                    {
                        Logger.Log((int)supp.asBinder().DocType.ID, Logger.Operations.CreateBinder, Logger.Statuses.Error, (int)SuppBinder.ID, string.Empty, string.Empty, SuppBinder.Title, response);
                        return "Error updating binder for " + supp.SupplierName + "\n" + response;
                    }
                    Logger.Log((int)SuppBinder.DocType.ID, Logger.Operations.CreateBinder, Logger.Statuses.OK, (int)SuppBinder.ID, string.Empty, string.Empty, SuppBinder.Title, string.Empty);
                }
            }
            return "";
        }
        private string CreateSupplierInDox(FlexSupplier supp)
        {
            using (OdbcConnection DbConnection = new OdbcConnection(doxParams["BAANDB"]))
            {
                openConnection(DbConnection);
                // Read data from BAAN DB
                logger.WriteEntry("before GetSupplierDetails  ");
                supp.GetSupplierDetails(DbConnection, doxParams);
                logger.WriteEntry("befor updateorcreateSuuplierbinder ");
                // Update the binder entity
                DOXAPI.TreeItemWithDocType suppBinder = dox.GetTreeItemWithDocType(token, supp.asFetchItem());
                response = supp.updateSupplierFields(suppBinder);
                if (response != "")
                {
                    return response;
                }
                // Save updates to DOX-Pro
                response = dox.UpdateTreeItemWithDocType(token, suppBinder);
                if (response != "Item updated." && response != String.Empty)
                {
                    Logger.Log((int)supp.asBinder().DocType.ID, Logger.Operations.CreateBinder, Logger.Statuses.Error, (int)suppBinder.ID, string.Empty, string.Empty, suppBinder.Title, response);
                    return "Error updating binder for " + supp.SupplierName + "\n" + response;
                }
                Logger.Log((int)suppBinder.DocType.ID, Logger.Operations.CreateBinder, Logger.Statuses.OK, (int)suppBinder.ID, string.Empty, string.Empty, suppBinder.Title, string.Empty);
                return "";
            }
        }
        public string FindFlexStorenextSupplierID(string Temp_SupplierID, string company)
        {

            using (OdbcConnection DbConnection = new OdbcConnection(doxParams["BAANDB"]))
            {
                openConnection(DbConnection);
                OdbcCommand DbCommand = DbConnection.CreateCommand();
                DbCommand.CommandText = String.Format(doxParams["FlexStorenextSuppQuery"], company, Temp_SupplierID);
                System.Diagnostics.EventLog.WriteEntry("GetFlexSupplierID", "the query: " + DbCommand.CommandText, System.Diagnostics.EventLogEntryType.Information);

                OdbcDataReader reader;

                reader = DbCommand.ExecuteReader();



                if (reader.Read())
                {
                    string suno = reader.GetString(0);
                    reader.Close();
                    return suno;

                }
                else
                {
                    reader.Close();
                    return "";//to do remove number here
                }
            }

        }
        private void doxLogin()
        {
            //if (String.IsNullOrEmpty(token))
            {
                token = dox.Login("baanint", "fl3x8aan1n7", doxParams["DoxEnv"]);
            }
        }
        //The function to get the data from baan for packing slip - to xml
        public string getDataFromBaan(string fullname, string company, string PackingSlipNo)
        {
            int numLines = 0;
            lines = new List<PackingSlipXMLEnvelopeLine>();
            inv_xml = new PackingSlipXML();
            if (inv_xml.Envelope == null)
            {
                inv_xml.Envelope = new PackingSlipXMLEnvelope();
            }
            if (inv_xml.Envelope.Header == null)
            {
                inv_xml.Envelope.Header = new PackingSlipXMLEnvelopeHeader();
            }
            int j = 0, counter = 0, i = 0, counter2 = 0;
            try
            {
                using (OdbcConnection DbConnection = new OdbcConnection("DSN=BAAN"))
                {
                    openConnection(DbConnection);
                    using (OdbcCommand DbCommand = DbConnection.CreateCommand())
                    {
                        //check if the customer of the packing slip exists in the table of the customers-ttccom810 
                        DbCommand.CommandText = String.Format(
                        "select baan810.t_fnum from  baandb.ttccom810{0} as baan810 where " +//Table of Customers
                        "baan810.t_cuno in (select t_cuno from baandb.ttdsls040{0} where t_orno in (select distinct t_orno from baandb.ttdsls045{0} as DSLS045 " +
                         " where DSLS045.t_dino={1})) ", company, PackingSlipNo);

                        try
                        {
                            DbReader = DbCommand.ExecuteReader();

                        }
                        catch (Exception ex)
                        {
                            return "Exception in getting Data From BAAN : " + ex.Message + "\n";
                        }
                        if (!DbReader.Read())
                        {
                            return "The Customer of Packing Slip no " + PackingSlipNo + " is not exists in BAAN \n";

                        }
                        DbReader.Close();
                        //get the data from tables in BAAN for the Packing Slip
                        //     DbCommand.CommandText = String.Format(
                        //       "select distinct baan000.t_send, baan000.t_nama," +//1,2
                        //       "baan810.t_fnum," +//3
                        //       "baan045.t_dino,baan045.t_ddat, baan045.t_pono, baan045.t_item, baan045.t_dqua,"//4,5,6,7,8
                        //       + "baan010.t_nama, baan010.t_namc, baan010.t_namd, baan010.t_name,baan010.t_ccty,"//9,10,11,12,13
                        //       + "baan040.t_refa,baan040.t_odat,baan040.t_orno,"//14,15,16
                        //       + "baan041.t_epos, baan041.t_cups, baan041.t_txta, baan041.t_revi,"//17,18,19,20
                        //       + "baan950.t_item," +
                        //        "baan001.t_dsca"
                        //       + " from baandb.ttccom000{0} as baan000, baandb.ttccom810{0} as baan810, baandb.ttdsls045{0} as baan045, baandb.ttccom010{0} as baan010," +
                        //       " baandb.ttdsls040{0} as baan040, baandb.ttdsls041{0} as baan041  , baandb.ttiitm950{0} as  baan950, baandb.ttiitm001{0} as baan001  "
                        //     + " Where baan045.t_dino={1} and (baan001.t_item= baan045.t_item and baan041.t_item= baan045.t_item and " +
                        //    " baan040.t_orno= baan045.t_orno and (baan041.t_pono= baan045.t_pono and baan041.t_orno= baan045.t_orno) and baan041.t_item= baan950.t_item and " +
                        //   " baan045.t_item= baan950.t_item and baan001.t_item= baan950.t_item and " +
                        //   " baan810.t_cuno in (select t_cuno from baandb.ttdsls040{0} where t_orno in (select distinct t_orno from baandb.ttdsls045{0} as DSLS045 where DSLS045.t_dino={1})) and " +
                        //    " baan010.t_cuno in (select t_cuno from baandb.ttdsls040{0} where t_orno in (select distinct t_orno from baandb.ttdsls045{0} as DSLS045 where DSLS045.t_dino={1})) " +
                        //")", company, PackingSlipNo);
                        DbCommand.CommandText = String.Format(
                       "select distinct baan000.t_send, baan000.t_nama," +//0,1
                       "baan810.t_fnum," +//2
                       "baan045.t_dino,baan045.t_ddat, baan045.t_pono, baan045.t_item, baan045.t_dqua,"//3,4,5,6,7
                       + "baan010.t_nama, baan010.t_namc, baan010.t_namd, baan010.t_name,baan010.t_ccty,"//8,9,10,11,12
                       + "baan040.t_refa,baan040.t_odat,baan040.t_orno,"//13,14,15
                       + "baan041.t_epos, baan041.t_cups, baan041.t_txta, baan041.t_revi,"//16,17,18,19
                       + "baan950.t_mitm," +//20
                       " baan001.t_dsca"//21
                       + " from baandb.ttccom000{0} as baan000, baandb.ttccom810{0} as baan810, baandb.ttdsls045{0} as baan045, baandb.ttccom010{0} as baan010," +
                        " baandb.ttdsls040{0} as baan040, baandb.ttdsls041{0} as baan041  , baandb.ttiitm950{0} as  baan950, baandb.ttiitm001{0} as baan001  "
                       + " Where baan045.t_dino={1} and (baan001.t_item=baan041.t_item and " +
                     " baan040.t_orno= baan045.t_orno and (baan041.t_pono= baan045.t_pono and baan041.t_orno= baan045.t_orno) and " +
                     "  baan950.t_cuno =baan810.t_cuno and baan950.t_mnum=\"999\" and baan950.t_item=baan001.t_item and " +
                     " baan810.t_cuno in (select t_cuno from baandb.ttdsls040{0} where t_orno in (select distinct t_orno from baandb.ttdsls045{0} as DSLS045 where DSLS045.t_dino={1})) and " +
                     " baan010.t_cuno in (select t_cuno from baandb.ttdsls040{0} where t_orno in (select distinct t_orno from baandb.ttdsls045{0} as DSLS045 where DSLS045.t_dino={1})) " +
                     ")", company, PackingSlipNo);

                        try
                        {
                            DbReader = DbCommand.ExecuteReader();

                        }
                        catch (Exception ex)
                        {
                            return "Exception in getting Data From BAAN : " + ex.Message + "\n";
                        }
                        float f = 0;
                        Int16 lineno = 0;
                        string qty = "";
                        while (DbReader.Read())
                        {
                            //The first time
                            if (counter == 0 && DbReader.GetString(0).Replace(" ", string.Empty) != string.Empty)
                            {

                                inv_xml.Envelope.Sender = DbReader.GetString(0).Replace(",", "").Replace(" ", "");
                                inv_xml.Envelope.Header.SupplierName = DbReader.GetString(1);
                                inv_xml.Envelope.Receiver = DbReader.GetString(2);
                                inv_xml.Envelope.Header.PackingSlipNo = DbReader.GetString(3);
                                inv_xml.Envelope.Header.DeliveryDate = DbReader.GetDateTime(4).ToString("dd/MM/yyyy");
                                //  f = 0;
                                //  bool b = float.TryParse(DbReader.GetFloat(7).ToString(), out f);

                                qty = "";
                                qty = DbReader.GetDecimal(7).ToString();
                                lineno = DbReader.GetInt16(5);
                                inv_xml.Envelope.Header.CompanyName = DbReader.GetString(8);
                                inv_xml.Envelope.Header.Address = DbReader.GetString(9) + " " + DbReader.GetString(10);
                                inv_xml.Envelope.Header.City = DbReader.GetString(11);
                                inv_xml.Envelope.Header.Country = DbReader.GetString(12);
                                inv_xml.Envelope.Header.Reference = new PackingSlipXMLEnvelopeHeaderReference[1];
                                inv_xml.Envelope.Header.Reference[0] = new PackingSlipXMLEnvelopeHeaderReference();
                                inv_xml.Envelope.Header.Reference[0].RefNo = DbReader.GetString(13);
                                inv_xml.Envelope.Header.Reference[0].RefType = PackingSlipXMLEnvelopeHeaderReferenceRefType.purchaseOrder;
                                inv_xml.Envelope.Header.Reference[0].RefDate = DbReader.GetDateTime(14).ToString("dd/MM/yyyy");

                                i++;
                                j++;
                                counter++;

                            }
                            if (counter >= 1)
                            {
                                if (qty == "")
                                {
                                    qty = DbReader.GetDecimal(7).ToString();
                                    //  bool b = float.TryParse(DbReader.GetFloat(7).ToString(), out f);

                                }

                                PackingSlipXMLEnvelopeLine line = new PackingSlipXMLEnvelopeLine();
                                line.ItemBarcode = DbReader.GetString(6);
                                if (qty.IndexOf('.') != -1 && (qty.IndexOf('.') + 3) <= qty.Length)

                                    line.UnitsQty = qty.Substring(0, qty.IndexOf('.') + 3);

                                else
                                    line.UnitsQty = qty;
                                // line.UnitsQty = Math.Round(f, 2).ToString();
                                string LineNo = DbReader.GetString(16);
                                if (lineno == 0)
                                    lineno = DbReader.GetInt16(5);
                                line.LineNo = lineno.ToString();
                                line.SupplierLineNo = LineNo;
                                line.UnitsQtyMea = "EACH";
                                line.Comments = DbReader.GetString(18);
                                line.Revision = DbReader.GetString(19);
                                line.CustomerBarcode = DbReader.GetString(20);
                                line.CustomerItemDescription = DbReader.GetString(21);
                                line.ItemDescription = DbReader.GetString(21);
                                line.Reference = new PackingSlipXMLEnvelopeLineReference[1];
                                line.Reference[0] = new PackingSlipXMLEnvelopeLineReference();
                                line.Reference[0].RefNo = DbReader.GetString(13);
                                line.Reference[0].SupplierRefNo = LineNo;
                                line.Reference[0].RefDate = DbReader.GetDateTime(14).ToString("dd/MM/yyyy");
                                line.Reference[0].RefType = PackingSlipXMLEnvelopeLineReferenceRefType.purchaseOrder;
                                numLines++;
                                lines.Add(line);
                                f = 0;
                                qty = "";
                                lineno = 0;
                                counter++;
                            }
                            counter2++;
                        }
                        DbReader.Close();
                        //get the data from tables in BAAN for the Packing Slip
                        //if there is no CustomerBarcode
                        //so the CustomerBarcode will be as the ItemBarcode
                        //(There is no data in ttiitm950 and 0 rows returned )
                        if (counter2 != 0)
                        {
                            //The pdf file name
                            inv_xml.Envelope.Header.SNAttachName = System.IO.Path.GetFileName(fullname).Replace(".xml", ".pdf");
                            //The Time - create the xml 
                            inv_xml.Envelope.Header.PackingSlipDate = DateTime.Now.ToString("dd/MM/yyyy");
                            inv_xml.Envelope.Details = new PackingSlipXMLEnvelopeLine[numLines + 1];
                            //Put The lines in the Details
                            inv_xml.Envelope.Details = lines.ToArray();
                            //Try write the xml file
                            return SerializeToXML(fullname).ToString();
                        }
                        else
                        {
                            DbCommand.CommandText = String.Format(
                                                  "select distinct baan000.t_send, baan000.t_nama," +//0,1
                                               "baan810.t_fnum," +//2
                                               "baan045.t_dino,baan045.t_ddat, baan045.t_pono, baan045.t_item, baan045.t_dqua,"//3,4,5,6,7
                                               + "baan010.t_nama, baan010.t_namc, baan010.t_namd, baan010.t_name,baan010.t_ccty,"//8,9,10,11,12
                                               + "baan040.t_refa,baan040.t_odat,baan040.t_orno,"//13,14,15
                                               + "baan041.t_epos, baan041.t_cups, baan041.t_txta, baan041.t_revi,"//16,17,18,19
                                               + "baan001.t_dsca"//20
                                               + " from baandb.ttccom000{0} as baan000, baandb.ttccom810{0} as baan810, baandb.ttdsls045{0} as baan045, baandb.ttccom010{0} as baan010," +
                                               " baandb.ttdsls040{0} as baan040, baandb.ttdsls041{0} as baan041  , baandb.ttiitm001{0} as baan001  "
                                             + " Where baan045.t_dino={1} and (baan001.t_item= baan045.t_item and baan041.t_item= baan045.t_item and " +
                                              " baan040.t_orno= baan045.t_orno and (baan041.t_pono= baan045.t_pono and baan041.t_orno= baan045.t_orno) and  " +
                                              " baan010.t_cuno in (select t_cuno from baandb.ttdsls040{0} where t_orno in (select distinct t_orno from baandb.ttdsls045{0} as DSLS045 where DSLS045.t_dino={1})) and " +
                                               " baan810.t_cuno in (select t_cuno from baandb.ttdsls040{0} where t_orno in (select distinct t_orno from baandb.ttdsls045{0} as DSLS045 where DSLS045.t_dino={1})) " +
                                              ")", company, PackingSlipNo);

                            try
                            {
                                DbReader = DbCommand.ExecuteReader();

                            }
                            catch (Exception ex)
                            {
                                return "Exception in getting Data From BAAN : " + ex.Message + "\n";
                            }
                            j = 0; counter = 0; i = 0; counter2 = 0;

                            f = 0;
                            lineno = 0;
                            qty = "";
                            while (DbReader.Read())
                            {

                                //The first time
                                if (counter == 0 && DbReader.GetString(0).Replace(" ", string.Empty) != string.Empty)
                                {

                                    inv_xml.Envelope.Sender = DbReader.GetString(0).Replace(",", "").Replace(" ", "");
                                    inv_xml.Envelope.Header.SupplierName = DbReader.GetString(1);
                                    inv_xml.Envelope.Receiver = DbReader.GetString(2);
                                    inv_xml.Envelope.Header.PackingSlipNo = DbReader.GetString(3);
                                    inv_xml.Envelope.Header.DeliveryDate = DbReader.GetDateTime(4).ToString("dd/MM/yyyy");
                                    // f = 0;
                                    // bool b = float.TryParse(DbReader.GetFloat(7).ToString(), out f);
                                    qty = "";
                                    qty = DbReader.GetDecimal(7).ToString();

                                    lineno = DbReader.GetInt16(5);
                                    inv_xml.Envelope.Header.CompanyName = DbReader.GetString(8);
                                    inv_xml.Envelope.Header.Address = DbReader.GetString(9) + " " + DbReader.GetString(10);
                                    inv_xml.Envelope.Header.City = DbReader.GetString(11);
                                    inv_xml.Envelope.Header.Country = DbReader.GetString(12);
                                    inv_xml.Envelope.Header.Reference = new PackingSlipXMLEnvelopeHeaderReference[1];
                                    inv_xml.Envelope.Header.Reference[0] = new PackingSlipXMLEnvelopeHeaderReference();
                                    inv_xml.Envelope.Header.Reference[0].RefNo = DbReader.GetString(13);
                                    inv_xml.Envelope.Header.Reference[0].RefType = PackingSlipXMLEnvelopeHeaderReferenceRefType.purchaseOrder;
                                    inv_xml.Envelope.Header.Reference[0].RefDate = DbReader.GetDateTime(14).ToString("dd/MM/yyyy");

                                    i++;
                                    j++;
                                    counter++;

                                }
                                if (counter >= 1)
                                {





                                    if (qty == "")
                                    {
                                        qty = DbReader.GetDecimal(7).ToString();
                                        //  bool b = float.TryParse(DbReader.GetFloat(7).ToString(), out f);

                                    }

                                    PackingSlipXMLEnvelopeLine line = new PackingSlipXMLEnvelopeLine();
                                    line.ItemBarcode = DbReader.GetString(6);
                                    if (qty.IndexOf('.') != -1 && (qty.IndexOf('.') + 3) <= qty.Length)

                                        line.UnitsQty = qty.Substring(0, qty.IndexOf('.') + 3);
                                    //     line.UnitsQty = Math.Round(decimal.Parse(qty), 2).ToString();
                                    //    logger.WriteEntry("  line.UnitsQty  - " + line.UnitsQty);
                                    else
                                        line.UnitsQty = qty;
                                    // line.UnitsQty = Math.Round(f, 2).ToString();


                                    string LineNo = DbReader.GetString(16);
                                    if (lineno == 0)
                                        lineno = DbReader.GetInt16(5);
                                    line.LineNo = lineno.ToString();
                                    line.SupplierLineNo = LineNo;
                                    line.UnitsQtyMea = "EACH";
                                    line.Comments = DbReader.GetString(18);
                                    line.Revision = DbReader.GetString(19);
                                    //  line.CustomerBarcode = DbReader.GetString(6);
                                    line.CustomerBarcode = " ";
                                    line.CustomerItemDescription = DbReader.GetString(20);
                                    line.ItemDescription = DbReader.GetString(20);
                                    line.Reference = new PackingSlipXMLEnvelopeLineReference[1];
                                    line.Reference[0] = new PackingSlipXMLEnvelopeLineReference();
                                    line.Reference[0].RefNo = DbReader.GetString(13);
                                    line.Reference[0].SupplierRefNo = LineNo;
                                    line.Reference[0].RefDate = DbReader.GetDateTime(14).ToString("dd/MM/yyyy");
                                    line.Reference[0].RefType = PackingSlipXMLEnvelopeLineReferenceRefType.purchaseOrder;
                                    numLines++;
                                    lines.Add(line);
                                    f = 0;
                                    qty = "";
                                    lineno = 0;
                                    counter++;
                                }
                                counter2++;

                            }


                        }

                    }

                }

                if (counter2 != 0)
                {
                    inv_xml.Envelope.Header.SNAttachName = System.IO.Path.GetFileName(fullname).Replace(".xml", ".pdf");
                    inv_xml.Envelope.Header.PackingSlipDate = DateTime.Now.ToString("dd/MM/yyyy");
                    inv_xml.Envelope.Details = new PackingSlipXMLEnvelopeLine[numLines + 1];
                    inv_xml.Envelope.Details = lines.ToArray();
                    return SerializeToXML(fullname).ToString();
                }
                else
                {
                    return "There is no data in BAAN for Packing Slip " + PackingSlipNo;
                }
            }

            catch (Exception nr)
            {
                return "Exception in getting Data From BAAN : " + nr.Message + "\n";

            }

        }

        private string SerializeToXML(string path, bool debug = true)
        {

            try
            {

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
            catch (Exception e)
            {
                return "Exception in create xml file for Packing Slip  -  " + e.Message;
            }
            return "true";
        }
        //Ayala 13.04.2015
        private List<string> ExecuteCommand(string filename)
        {
            List<string> barcode = new List<string>();
            string error = "";
            try
            {

                string pdf = @"""" + filename + @"""";
                try
                {
                    logger.WriteEntry("try to create barcode.bat file ");
                    FileStream fs1 = new FileStream("barcode.bat", FileMode.OpenOrCreate, FileAccess.Write);
                    StreamWriter writer = new StreamWriter(fs1);
                    writer.WriteLine(@"cd C:\Program Files\DMS Inc\DeliveryNotesSetup\GetBarcodeFromFileCMD");
                    // logger.WriteEntry(@"ReadBarCodesFromFile.exe  " + pdf);
                    writer.WriteLine(@"ReadBarCodesFromFile.exe  " + pdf);
                    writer.Close();
                }
                catch (Exception ex)
                {
                    logger.WriteEntry("error in create bat file " + ex.Message);
                    barcode.Clear();
                    barcode.Add("Error");
                    error = "Error";
                }
                System.Diagnostics.Process proc = new System.Diagnostics.Process();
                proc.StartInfo.FileName = "barcode.bat";
                proc.StartInfo.Verb = "runas";
                proc.StartInfo.WorkingDirectory = @"";
                proc.StartInfo.CreateNoWindow = true;
                proc.StartInfo.UseShellExecute = false;
                proc.StartInfo.RedirectStandardError = true;
                proc.StartInfo.RedirectStandardOutput = true;
                proc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                proc.Start();
                string output = proc.StandardOutput.ReadToEnd();
                proc.Close();
                logger.WriteEntry("output - " + output);
                if (output.Contains("N/A"))
                {
                    error = "Error";// error2 = "Error";
                }
                if (error == "")
                {
                    try
                    {
                        // error = output.IndexOf("==>") != -1;
                        error = output.Remove(0, output.IndexOf("==>") + 4);
                        error = error.IndexOf("C") != -1 ? error.Substring(0, error.IndexOf("C")) : error;
                        if (error.Length >= 15)
                            if (error.Contains('M'))
                                error = error.Substring(0, 17);
                            else
                                error = error.Substring(0, 16);
                        //one space at end
                        if (error.LastIndexOf(" ") > 6)
                            error = error.Substring(0, error.LastIndexOf(" "));
                        string b = string.Empty;


                        for (int i = 0; i < error.Length; i++)
                        {
                            if (Char.IsDigit(error[i]) || i <= 6)
                                b += error[i];
                        }
                        logger.WriteEntry("error = b;  " + b);
                        error = b;

                        logger.WriteEntry("recognize barcode - " + error);

                    }
                    catch (Exception ex)
                    {
                        logger.WriteEntry("error by get barcode the second way " + ex.Message);
                        error = "Error";
                        //     error2 = "Error";

                    }
                }


            }
            catch (Exception ex)
            {
                logger.WriteEntry("error by get barcode the second way " + ex.Message);
                error = "Error";
                //  error2 = "Error";
            }
            if (error == "" || error.Contains("=") || error.Contains(".pdf"))
            {
                error = "Error";
                // error2 = "Error";
                logger.WriteEntry("error by get barcode the second way " + error);

            }
            else
                logger.WriteEntry("success get barcode the second way " + error);
            barcode.Clear();
            barcode.Add(error);

            return barcode;


        }

        private List<string> ExecuteCommandShip(string filename, string onlyNameFile)
        {
            List<string> barcode = new List<string>();
            string error = "";
            //try
            //{
            int count = 0;
            string pdf = @"""" + filename + @"""";
            try
            {
                BarcodeReader reader = new BarcodeReader();
                reader.Code39 = true; reader.Code128 = true;
                Inlite.ClearImageNet.Barcode[] barcodes = reader.Read(filename);

                if (barcodes != null && barcodes.Length > 0)
                {
                    logger.WriteEntry("barcodes.Length - " + barcodes.Length);
                        
                    foreach (Inlite.ClearImageNet.Barcode oBarcode in barcodes)
                    {
                        barcode.Add(oBarcode.Text);
                        logger.WriteEntry("oBarcode.Text - " + oBarcode.Text);

                        int flager = 0;
                        string s = oBarcode.ToString();
                        if (s.Contains("/"))
                        {
                            s = s.Substring(0, s.IndexOf("/"));
                            if (s.Length == 10)
                                flager = 1;
                        }

                        long Letter;

                        if (long.TryParse(s, out Letter))
                        {


                            if (((oBarcode.Length >= 10 && s.Length == 10) || (flager == 1 && s.Length == 10)) && oBarcode.Text.Substring(0, 2) == (DateTime.Today.Year % 100).ToString())
                                count++;

                        }
                    
                    
                    }

                }

                if (barcode.Count > 0 && count == 0)
                {
                    logger.WriteEntry("barcodes.Count - " + barcodes.Count()
                        + " ,barcode.Count - " + barcode.Count);
                    try
                    {
                        try
                        {
                            try
                            {
                                logger.WriteEntry("try to create barcode.bat file ");
                                FileStream fs1 = new FileStream("barcode.bat", FileMode.OpenOrCreate, FileAccess.Write);
                                StreamWriter writer = new StreamWriter(fs1);
                                writer.WriteLine(@"cd C:\Program Files\DMS Inc\DeliveryNotesSetup\GetBarcodeFromFileCMD");

                                writer.WriteLine(@"ReadBarCodesFromFile.exe  " + pdf);
                                writer.Close();
                            }
                            catch (Exception ex)
                            {
                                logger.WriteEntry("error in create bat file " + ex.Message);

                                barcode.Add("Error");
                                error = "Error";
                                return barcode;
                            }
                            System.Diagnostics.Process proc = new System.Diagnostics.Process();
                            proc.StartInfo.FileName = "barcode.bat";
                            proc.StartInfo.Verb = "runas";
                            proc.StartInfo.WorkingDirectory = @"";
                            proc.StartInfo.CreateNoWindow = true;
                            proc.StartInfo.UseShellExecute = false;
                            proc.StartInfo.RedirectStandardError = true;
                            proc.StartInfo.RedirectStandardOutput = true;
                            proc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                            proc.Start();
                            string output = proc.StandardOutput.ReadToEnd();
                            proc.Close();
                            logger.WriteEntry("output - " + output);

                            if (output.Contains("N/A"))
                            {
                                error = "Error";// error2 = "Error";
                            }
                            if (error == "")
                            {
                                try
                                {
                                    logger.WriteEntry("beforelines");
                                    try
                                    {
//                                        string[] lines = output.Remove(0, output.IndexOf("==>")).Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
                                        string[] lines = output.Remove(0, output.IndexOf("system32>")).Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
                                        logger.WriteEntry("afterlines ");
                                        foreach (var item in lines)
                                        {
                                            logger.WriteEntry("in lines ");

                                            if (item.IndexOf("==>") >= 0)
                                            {
                                                logger.WriteEntry("in in lines ");
                                                barcode.Add(item.Remove(0, item.IndexOf("==>") + 4));

                                                logger.WriteEntry("barcode - " + item.Remove(0, item.IndexOf("==>") + 4));
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        logger.WriteEntry("miryam exception in lines " + ex.Message);
                                    }


                                }
                                catch (Exception ex)
                                {
                                    foreach (String s in barcode)
                                    {logger.WriteEntry("b " + s);
                                        
                                    }
                                    logger.WriteEntry("error by get barcode the second way1 " + ex.Message);
                                    error = "Error";
                                    //  barcode.Clear();
                                    barcode.Add("Error");
                                    return barcode;


                                }
                            }
                            if (barcode.Count == 0)
                                barcode.Add("Error");

                        }
                        catch (Exception ex)
                        {
                            logger.WriteEntry("error by get barcode the second way2 " + ex.Message);
                            error = "Error";

                            barcode.Add("Error");
                            return barcode;

                        }

                        return barcode;


                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("error by get barcode the second way " + ex.Message);
                        error = "Error";

                        barcode.Add("Error");
                        return barcode;

                    }
                }
                else if (barcode.Count == 0) { barcode.Add("Error"); }

            }
            catch (Exception ex)
            {
                Console.WriteLine("error by get barcode the second way " + ex.Message);
                error = "Error";

                barcode.Add("Error");
                return barcode;

            }
            if (barcode.Count == 0)
                barcode.Add("Error");
            return barcode;

        }
        public string setMetaDataForShipment(string fullname, string barcodesLotAndMakats, string typeFile)
        {

            try
            {
                List<string> barcodes2 = new List<string>();
                string[] lines = barcodesLotAndMakats.Split(',');
                foreach (string line in lines)
                {
                    barcodes2.Add(line);
                }
                using (MySqlConnection DbConnection = new MySqlConnection(Settings.Default.ConnectionStringForBaan))
                {
                    logger.WriteEntry("in get getDataForShipment " + DbConnection.ConnectionString.ToString());
                    openMysqlConnection(DbConnection);
                    try
                    {
                        return getDataForShipment(DbConnection, barcodes2.Count > 2, barcodes2, typeFile, fullname);
                    }
                    catch (Exception ex)
                    {
                        logger.WriteEntry("Error " + ex.Message);
                        try
                        {
                            using (IDataSupplier dataSupplier = DataManager.GetDataSupplier(DataManager.defaultType, Settings.Default.ConnectionString))
                            {
                                dataSupplier.OpenQuery();
                                dataSupplier.AddParameter("FileName", fullname);
                                dataSupplier.AddParameter("status", "-1");
                                dataSupplier.AddParameter("ErrorReason", ex.Message);
                                dataSupplier.Execute("update [shipmentArchive] set status=@status,ErrorReason=@ErrorReason where FileName=@FileName ");
                            }
                            logger.WriteEntry("Error in Get Data From Baan -1");
                            string renameFile = typeFile.ToLower().Contains("COC".ToLower()) ? "COC_" : typeFile.ToLower().Contains("Invoice".ToLower()) ? "Invoice_" : "DeliveryNote_";
                            LogArchiveShipment(fullname, typeFile, 0, 2, ex.Message);
                            movetoDefQShipment(fullname, 6, ".", renameFile, " Error in Archive the doc");
                        }
                        catch (Exception ex2)
                        {
                            logger.WriteEntry("Error in getDataForShipment- " + ex2.Message);
                        }
                        return "Error - " + ex.Message;
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteEntry("Error " + ex.Message);
                try
                {
                    using (IDataSupplier dataSupplier = DataManager.GetDataSupplier(DataManager.defaultType, Settings.Default.ConnectionString))
                    {
                        dataSupplier.OpenQuery();
                        dataSupplier.AddParameter("FileName", fullname);
                        dataSupplier.AddParameter("status", "-1");
                        dataSupplier.AddParameter("ErrorReason", ex.Message);
                        dataSupplier.Execute("update [shipmentArchive] set status=@status,ErrorReason=@ErrorReason where FileName=@FileName ");
                    }
                    logger.WriteEntry("Error in Get Data From Baan -1");
                    string renameFile = typeFile.ToLower().Contains("COC".ToLower()) ? "COC_" : typeFile.ToLower().Contains("Invoice".ToLower()) ? "Invoice_" : "DeliveryNote_";
                    LogArchiveShipment(fullname, typeFile, 0, 2, ex.Message);
                    movetoDefQShipment(fullname, 6, ".", renameFile, " Error in Archive the doc");
                }
                catch (Exception ex2)
                {
                    logger.WriteEntry("Error in getDataForShipment- " + ex2.Message);
                }
                return "Error - " + ex.Message;
            }
        }
        //tehila add 3.8.2017
        public string RunStoredProcLots(string shipmentNo, string barcodesLotAndMakats, string filename,string type)
        {
            logger.WriteEntry(" in RunStoredProcLots- " );
             SqlConnection conn = null;
            SqlDataReader rdr = null;

            

            try
            {
                conn = new SqlConnection("Data Source=mignt014;Initial Catalog=DoxPro_Env_Flex;Integrated Security=True");
                conn.Open();
                SqlCommand cmd = new SqlCommand("dbo.[AddUpdateLOT]", conn);
                logger.WriteEntry("11: " + cmd.ToString());
                 cmd.Parameters.Add(new SqlParameter("@lots", barcodesLotAndMakats));
                cmd.Parameters.Add(new SqlParameter("@shipmentNo", shipmentNo));
                logger.WriteEntry("22: " );
                logger.WriteEntry("lots: " + barcodesLotAndMakats+ "  shipmentNo: " + shipmentNo);
                logger.WriteEntry("33: ");
                 cmd.CommandType = CommandType.StoredProcedure;
                 logger.WriteEntry("44: ");
                rdr = cmd.ExecuteReader();
                logger.WriteEntry("55: ");
                /*while (rdr.Read())
                {
                    Console.WriteLine(
                        "Product: {0,-25} Price: ${1,6:####.00}",
                        rdr["TenMostExpensiveProducts"],
                        rdr["UnitPrice"]);
                }*/
                return "";
            }
            catch (Exception e)
            {
                logger.WriteEntry("error RunStoredProcLots " + e.Message);
               /* string name = filename.Substring(filename.LastIndexOf('\\'));
                try
                {
                    System.IO.File.Move(filename, doxParams["MachsanErrors"] + type + "_" + name.Substring(1));
                    filename = doxParams["MachsanErrors"] + type + "_" + name.Substring(1);
                }

                catch (Exception we)
                {
                    logger.WriteEntry(we.Message);
                }*/
                return "error RunStoredProcLots " + e.Message;

            }
            finally
            {
                if (conn != null)
                {
                    conn.Close();
                }
                if (rdr != null)
                {
                    rdr.Close();
                }
            }
        }
        public string RunStoredProcMakats(string shipmentNo, string barcodesLotAndMakats, string filename,string type)
        {
            SqlConnection conn = null;
            SqlDataReader rdr = null;


            try
            {
                conn = new SqlConnection("Data Source=mignt014;Initial Catalog=DoxPro_Env_Flex;Integrated Security=True");
                conn.Open();
                SqlCommand cmd = new SqlCommand("dbo.[AddUpdateMAKAT]", conn);
                cmd.Parameters.Add(new SqlParameter("@makats", barcodesLotAndMakats));
                cmd.Parameters.Add(new SqlParameter("@shipmentNo", shipmentNo));
                cmd.CommandType = CommandType.StoredProcedure;
                rdr = cmd.ExecuteReader();
                return "";

            }
            catch (Exception e)
            {
                logger.WriteEntry("error RunStoredProcMakats "+e.Message);
                /*string name = filename.Substring(filename.LastIndexOf('\\'));
                try
                {
                    System.IO.File.Move(filename, doxParams["MachsanErrors"] + type + "_" + name.Substring(1));
                    filename = doxParams["MachsanErrors"] + type + "_" + name.Substring(1);
                }

                catch (Exception we)
                {
                    logger.WriteEntry(we.Message);
                }*/
                return "error RunStoredProcMakats " + e.Message;
            }
            finally
            {
                if (conn != null)
                {
                    conn.Close();
                }
                if (rdr != null)
                {
                    rdr.Close();
                }
            }
        }
      

        private string LotsAndMakats(string filename, string barcodesLotAndMakats,string type)
        {
            string shipmentNo = "";
            string mayLot = "";
            string mayMakat = "";
            try
            {
                List<string> barcodes2 = new List<string>();
                string[] lines = barcodesLotAndMakats.Split(',');
                foreach (string line in lines)
                {
                    barcodes2.Add(line);
                }
                using (MySqlConnection DbConnection = new MySqlConnection(Settings.Default.ConnectionStringForBaan))
                {
                    logger.WriteEntry("in get getDataForShipment " + DbConnection.ConnectionString.ToString());
                    openMysqlConnection(DbConnection);
                    try
                    {
                        logger.WriteEntry("going to get shipmentNo");
                        shipmentNo = getDataForShipment(DbConnection, barcodes2.Count > 2, barcodes2, "", filename);
                        logger.WriteEntry("shipmentNo is:  "+shipmentNo);
                    }
                    catch(Exception e) {
                        logger.WriteEntry(e.Message);
                        return e.Message;

                    }
                }
               
                if (!(shipmentNo.Equals("")))
                {
                    shipmentNo = shipmentNo.Trim();
                    string lots = "";
                    string makats = "";
                // string s1=   shipmentNo = "'" + shipmentNo.Trim() + "'";
                    logger.WriteEntry("shipmentNo 2 is:  " + shipmentNo);
                    foreach (string item in barcodes2)
                    {
                        int flager = 0;
                        string s = item;
                        if (item.Contains("/"))
                        {
                            s = item.Substring(0, item.IndexOf("/"));
                            if (s.Length == 10)
                                flager = 1;
                        }

                        long Letter;

                        if (long.TryParse(s, out Letter))
                        {


                            if (((item.Length == 10 && s.Length == 10) || (flager == 1 && s.Length == 10)) && item.Substring(0, 2) == (DateTime.Today.Year % 100).ToString() && mayLot == "")
                                mayLot = item;

                            if (mayLot.Contains("/"))
                                mayLot = mayLot.Substring(0, mayLot.IndexOf("/"));
                            doxLogin();
                            //lots.Add(mayLot);
                            logger.WriteEntry("going to  RunStoredProcLots for " + filename + " maylot: " + mayLot);
                            logger.WriteEntry("tttt");
                            if (lots.Equals(""))
                                lots += mayLot;
                            else
                                lots += "," + mayLot;
                           // string response1 = RunStoredProcLots(shipmentNo, mayLot, filename, type);
                            


                            /*try
                            {
                                using (IDataSupplier dataSupplier = DataManager.GetDataSupplier(DataManager.defaultType, Settings.Default.ConnectionString))
                                {
                                    dataSupplier.OpenQuery();
                                    dataSupplier.AddParameter("shipmentNo", shipmentNo);
                                    dataSupplier.AddParameter("mayLot", mayLot);
                                    DataSet ds = dataSupplier.GetData("select id from tree_tab where cfg_id=21 and title='"+shipmentNo+"'");

                                    DataTable dt = ds.Tables[0];
                                    if (dt != null && dt.Rows.Count > 0)
                                    {


                                        string dr = dt.Rows[0][0].ToString();
                                        if (dr != null)
                                            {
                                                dataSupplier.AddParameter("dr", dr);
                                                dataSupplier.Execute("insert into cfg_27_tab values (@dr,@shipmentNo,@mayLot)");
                                            }
                                   
                                    }

                                }
                            }
                            catch (Exception e)
                            {
                                logger.WriteEntry(e.Message);
                            }*/


                        }
                        else if (item.Length >= 10 && item.Substring(0, 2) == (DateTime.Today.Year % 100).ToString() && mayLot != "")
                            continue;
                        else
                        {
                            try
                            {
                                string item2 = item.ToUpper();
                                Regex rgx = new Regex("[^a-zA-Z0-9 -]");
                                item2 = rgx.Replace(item2, "");
                                mayMakat = item2;
                                doxLogin();
                                if (mayMakat.Equals(""))
                                    makats += mayMakat;
                                else
                                    makats += "," + mayMakat;
                            }
                            catch (Exception ex)
                            {
                                mayMakat = "";
                            }
                        }
                    }
                   // bool flag = true;
                    string exMessage="";
                    
                        try
                        {
                            logger.WriteEntry("going to  RunStoredProcLots for " + filename + " maylot: " + lots);
                            string response = RunStoredProcLots(shipmentNo, lots, filename, type);
                            if (!(response.Equals("")))
                            {
                                //logger.WriteEntry("going to machsanerrors " );
                                 exMessage = "error in running RunStoredProcLots";
                               /* string name = filename.Substring(filename.LastIndexOf('\\'));
                                try
                                {
                                    System.IO.File.Move(filename, doxParams["MachsanErrors"] + type + "_" + name.Substring(1));
                                    filename = doxParams["MachsanErrors"] + type + "_" + name.Substring(1);
                                }

                                catch (Exception we)
                                {
                                    logger.WriteEntry(we.Message);
                                }
                                flag = false;*/
                            }
                            logger.WriteEntry("finish RunStoredProcLots");
                        }
                        catch (Exception e)
                        {
                            logger.WriteEntry( e.Message);
                             exMessage = "error in running RunStoredProcLots";
                            /*string name = filename.Substring(filename.LastIndexOf('\\'));
                            try
                            {
                                System.IO.File.Move(filename, doxParams["MachsanErrors"] + type + "_" + name.Substring(1));
                                filename = doxParams["MachsanErrors"] + type + "_" + name.Substring(1);
                            }

                            catch (Exception we)
                            {
                                logger.WriteEntry(we.Message);
                            }
                            flag = false;*/
                         
                        }

                    

                    
                        try
                        {
                            logger.WriteEntry("going to  RunStoredProcMakats for " + filename + " mayMakat: " + makats);
                            string response = RunStoredProcMakats(shipmentNo, makats, filename, type);
                            if(!(response.Equals("")))
                            {
                                //logger.WriteEntry("going to machsanerrors " );
                                exMessage = "error in running StoredProcMakats";
                               /* string name = filename.Substring(filename.LastIndexOf('\\'));
                                try
                                {
                                    System.IO.File.Move(filename, doxParams["MachsanErrors"] + type + "_" + name.Substring(1));
                                    filename = doxParams["MachsanErrors"] + type + "_" + name.Substring(1);
                                }

                                catch (Exception we)
                                {
                                    logger.WriteEntry(we.Message);
                                }
                                flag = false;*/
                            }
                            logger.WriteEntry("finish RunStoredProcMakats");
                        }
                        catch (Exception e)
                        {
                            logger.WriteEntry(e.Message);
                            exMessage = "error in running StoredProcMakats";
                           /* string name = filename.Substring(filename.LastIndexOf('\\'));
                            try
                            {
                                System.IO.File.Move(filename, doxParams["MachsanErrors"] + type + "_" + name.Substring(1));
                                filename = doxParams["MachsanErrors"] + type + "_" + name.Substring(1);
                            }

                            catch (Exception we)
                            {
                                logger.WriteEntry(we.Message);
                            }*/
                             //flag = false;
                           
                        }

                   
                   // if (flag)
                        return "";
                   // else
                        //return exMessage;


                }
                else
                {
                    string exMessage = "there is no shipmentNo for the lot: " + mayMakat;
                    logger.WriteEntry(exMessage);
                 
                    
                  /*  string name = filename.Substring(filename.LastIndexOf('\\'));
                     try
                    {
                        System.IO.File.Move(filename, doxParams["MachsanErrors"] + type + "_" + name.Substring(1));
                        filename = doxParams["MachsanErrors"] + type + "_" + name.Substring(1);
                    }

                    catch (Exception we)
                    {
                        logger.WriteEntry(we.Message);
                    }*/
                     return exMessage;
                }
            }
            catch (Exception e)
            {
                logger.WriteEntry(e.Message);
                return e.Message;
            }

        


        }
        private string getDataForShipment(MySqlConnection DbConnection, bool moreThan1Sticker, List<string> LotsAndMakats, string typeFile, string fullname)
        {
            logger.WriteEntry("miryam getDataForShipment ");
            File.AppendAllText("c:\\miryam\\noam.txt", "getDataForShipment" + Environment.NewLine);
            bool success = false;

            string mayLot = "";
            string mayMakat = "";

            MySqlDataReader DbReader = null;
            MySqlCommand DbCommand = null;

            logger.WriteEntry("1: - " + DateTime.Now.ToLongTimeString());

            if (!moreThan1Sticker)
            {
                if (LotsAndMakats != null && LotsAndMakats.Count > 0 && LotsAndMakats[0].Length >= 10 && LotsAndMakats[0].Substring(0, 2) == (DateTime.Today.Year % 100).ToString())
                {
                    mayLot = LotsAndMakats[0];
                    if (LotsAndMakats.Count == 2)
                    {
                        try
                        {
                            string item2 = LotsAndMakats[1].ToUpper();
                            Regex rgx = new Regex("[^a-zA-Z0-9 -]");
                            item2 = rgx.Replace(item2, "");
                            mayMakat = item2;
                        }
                        catch
                        {
                            mayMakat = "";
                        }
                    }
                }
                else if (LotsAndMakats != null && LotsAndMakats.Count > 1 && LotsAndMakats[1].Length >= 10 && LotsAndMakats[1].Substring(0, 2) == (DateTime.Today.Year % 100).ToString())
                {
                    mayLot = LotsAndMakats[1];
                    try
                    {
                        string item2 = LotsAndMakats[0].ToUpper();
                        Regex rgx = new Regex("[^a-zA-Z0-9 -]");
                        item2 = rgx.Replace(item2, "");
                        mayMakat = item2;
                    }
                    catch
                    {

                        mayMakat = "";

                    }
                }
            }
            else
            {
                mayLot = "";
                mayMakat = "";
                foreach (string item in LotsAndMakats)
                {
                    int flager = 0;
                    string s = item;
                    if (item.Contains("/"))
                    {
                        s = item.Substring(0, item.IndexOf("/"));
                        if (s.Length == 10)
                            flager = 1;
                    }

                    long Letter;

                    if (long.TryParse(s, out Letter))
                    {
                        if (((item.Length == 10 && s.Length == 10) || (flager == 1 && s.Length == 10)) && item.Substring(0, 2) == (DateTime.Today.Year % 100).ToString() && mayLot == "")
                            mayLot = item;
                    }
                    else if (item.Length >= 10 && item.Substring(0, 2) == (DateTime.Today.Year % 100).ToString() && mayLot != "")
                        continue;
                    else
                    {
                        try
                        {
                            string item2 = item.ToUpper();
                            Regex rgx = new Regex("[^a-zA-Z0-9 -]");
                            item2 = rgx.Replace(item2, "");
                            mayMakat = item2;
                        }
                        catch (Exception ex)
                        {
                            mayMakat = "";
                        }
                    }
                }
            }
            if (mayLot.Contains("/"))
                mayLot = mayLot.Substring(0, mayLot.IndexOf("/"));
            string ErrorReason = "Lot isn''t correct Or Makats aren''t correct Or Lot Or/And Makat isn''t exists";
            logger.WriteEntry("mayLot  - " + mayLot + " , mayMakats-" + mayMakat);
            string cmd = "";

            if (mayLot != "" && mayLot.Length >= 10 && mayMakat != "" && mayMakat != ") and ")
            {
                //Get the Shipment No , Supplier No , And Supplier Data
               // cmd = "select @SupNo:=t_suno,@SupName:=t_nama,@ShipNo:=t_dino,@date:=t_date from dms_dox.receipts where t_clot like '%" + mayLot + "%' and t_item like '%" + mayMakat + "%';";
                //cmd += "select t_date as date,t_suno as supplierNomber,t_nama as supplierName,t_dino as shipmentNo,t_clot as Lot,t_item as Makat,t_namc+' '+t_namd+' '+t_name+' '+t_namf as supplierAdress,t_refs as PurchasingPerson,t_telp as Telephone from dms_dox.receipts where t_suno=@SupNo and t_nama=@SupName and t_dino=@ShipNo group by t_clot";
                //Get the Shipment No 
                cmd = "select top 1 t_dino from dms_dox.receipts where t_clot like '%" + mayLot + "%' and t_item like '%" + mayMakat + "%'";
                logger.WriteEntry("miryam cmd = " + cmd.ToString());
                

                try
                {
                    DbCommand = new MySqlCommand(cmd, DbConnection);
                    DbReader = DbCommand.ExecuteReader();
                    if (DbReader.HasRows)
                    {
                        string shipmentNo = "";
                        while (DbReader.Read())
                        {
                            shipmentNo= DbReader[0].ToString();
                        }
                        DbReader.NextResult();
                        DbReader.Close();
                        return shipmentNo;
                    }
                    else
                    {
                       if (DbReader != null)
                            DbReader.Close();
                     
                    }
                }
                catch (Exception ex)
                {
                    
                    if (DbReader != null)
                        DbReader.Close();
                    logger.WriteEntry("Error " + ex.Message);
                }
            }

            //Without makat
            if (!success && mayLot != "" && mayLot.Length >= 10)
            {    //Check
                logger.WriteEntry("mayLot - " + mayLot);
                //Get the Shipment No , Supplier No , And Supplier Data                
                cmd = "select t_dino from dms_dox.receipts where t_clot like '%" + mayLot + "%';";
                

                try
                {
                    if (DbReader != null)
                        DbReader.Close();
                    DbCommand = new MySqlCommand(cmd, DbConnection);
                    DbReader = DbCommand.ExecuteReader();
                }
                catch (Exception ex)
                {
                
                   return "Error - " + ex.Message;
                }
                try
                {
                    if (DbReader.HasRows)
                    {
                        
                            string shipmentNo = "";
                            while (DbReader.Read())
                            {
                                shipmentNo = DbReader[0].ToString();
                            }
                            DbReader.NextResult();
                            DbReader.Close();
                            return shipmentNo;
                        
                        
                    }
                    else
                    {
                        try
                        {
                            
                            
                            if (DbReader != null)
                                DbReader.Close();
                        }
                        catch (Exception ex2)
                        {
                            if (DbReader != null)
                                DbReader.Close();
                            logger.WriteEntry(ex2.Message);
                        }
                        return "Error in get shipmentNo";
                    }
                }
                catch (Exception ex)
                {
                    logger.WriteEntry("Error in update to shipment Get Data- " + ex.Message);
                    return "Error in get shipmentNo";
                }
            }
            else
            {
                try
                {
                   
                    if (DbReader != null)
                        DbReader.Close();
                }
                catch (Exception ex2)
                {
                    if (DbReader != null)
                        DbReader.Close();
                    logger.WriteEntry("Error in  get shipmentNo- " + ex2.Message);
                }
                return "Error in get shipmentNo";
            }
            
        }

        private string updateDataForShipmentDoc(MySqlCommand DbCommand, MySqlDataReader DbReader, string fullname, string typeFile, string mayLot, ref string ErrorReason, ref bool success, bool isEmail)
        {
            logger.WriteEntry("miryam begin archive");
            List<string> Lots = new List<string>();
            List<string> Makats = new List<string>();

            string shippmentNo = "";
            string supplierID = "";
            string SupplierNo = "";
            string CompanyNo = "";
            string SupplierName = "";
            string SupplierAddress = "";
            string PurchasingPerson = "";
            string SupplierPhone = "";
            string SearchKey = "";
            string SupplierEmail = "";
            string lotNo = "";
            string makat = "";

            if (DbReader.Read())
            {
                shippmentNo = DbReader.GetString(3).Replace(" ", "").Trim();
                supplierID = DbReader.GetString(1).Replace(" ", "").Trim();
                SupplierNo = supplierID;
                CompanyNo = "400";
                SupplierName = DbReader.GetString(2).Trim();
                SupplierAddress = DbReader.GetString(6).Replace(" ", "").Trim();
                PurchasingPerson = DbReader.GetString(7).Replace(" ", "").Trim();
                SupplierPhone = DbReader.GetString(8).Replace(" ", "").Trim();

                if (!Lots.Contains(DbReader.GetString(4).Replace(" ", "").Trim()))
                {
                    lotNo = lotNo + (DbReader.GetString(4).Replace(" ", "").Trim());
                    Lots.Add(DbReader.GetString(4).Replace(" ", "").Trim());
                }
                if (!Makats.Contains(DbReader.GetString(5).Trim(' ')))
                {
                    makat = makat + (DbReader.GetString(5).Trim(' '));
                    Makats.Add(DbReader.GetString(5).Trim(' '));
                }
                logger.WriteEntry("miryam end read db");
            }

            while (DbReader.Read())
            {
                logger.WriteEntry("miryam start read lots and makats");
                if (!Lots.Contains(DbReader.GetString(4).Replace(" ", "").Trim()))
                {
                    lotNo = lotNo + ("," + DbReader.GetString(4).Replace(" ", "").Trim());
                    Lots.Add(DbReader.GetString(4).Replace(" ", "").Trim());
                }
                if (!Makats.Contains(DbReader.GetString(5).Trim(' ')))
                {
                    makat = makat + ("," + DbReader.GetString(5).Trim(' '));
                    Makats.Add(DbReader.GetString(5).Trim(' '));
                }
            }
            logger.WriteEntry("lotNo - " + lotNo + " ,mayLot - " + mayLot);

            DbReader.Close();

            try
            {
                logger.WriteEntry("miryam in try ");
                using (IDataSupplier dataSupplier = DataManager.GetDataSupplier(DataManager.defaultType, Settings.Default.ConnectionString))
                {
                    dataSupplier.OpenQuery();
                    dataSupplier.AddParameter("lotsNew", lotNo);
                    dataSupplier.AddParameter("MakatsNew", makat);
                    dataSupplier.AddParameter("shippmentNo", shippmentNo);
                    dataSupplier.AddParameter("supplierID", supplierID);
                    dataSupplier.AddParameter("CompanyNo", CompanyNo);
                    dataSupplier.AddParameter("SupplierName", SupplierName);
                    dataSupplier.AddParameter("SupplierAddress", SupplierAddress);
                    dataSupplier.AddParameter("PurchasingPerson", PurchasingPerson);
                    dataSupplier.AddParameter("SupplierPhone", SupplierPhone);
                    dataSupplier.AddParameter("SearchKey", SearchKey);
                    dataSupplier.AddParameter("SupplierEmail", SupplierEmail);
                    dataSupplier.AddParameter("FileName", fullname);

                    dataSupplier.AddParameter("status", "1");
                    dataSupplier.Execute("update [shipmentArchive] set lotsNew=@lotsNew,MakatsNew=@MakatsNew,shippmentNo=@shippmentNo ," +
                                        " supplierID=@supplierID,CompanyNo=@CompanyNo,SupplierName=@SupplierName,SupplierAddress=@SupplierAddress,PurchasingPerson=@PurchasingPerson," +
                                        " SupplierPhone=@SupplierPhone,SearchKey=@SearchKey,SupplierEmail=@SupplierEmail,[status]=@status "
                                      + " \r\n\t\t\t\t    where FileName=@FileName and status='0'");
                }
                success = true;
                return "";
            }
            catch (Exception ex)
            {
                logger.WriteEntry("Error in update to shipmentArchive- " + ex.Message);
                ErrorReason = ex.Message;
                try
                {
                    using (IDataSupplier dataSupplier = DataManager.GetDataSupplier(DataManager.defaultType, Settings.Default.ConnectionString))
                    {
                        dataSupplier.OpenQuery();
                        dataSupplier.AddParameter("FileName", fullname);
                        dataSupplier.AddParameter("status", "-1");
                        dataSupplier.AddParameter("ErrorReason", ErrorReason);
                        dataSupplier.Execute("update [shipmentArchive] set status=@status,ErrorReason=@ErrorReason where FileName=@FileName and status='0'");
                    }
                    logger.WriteEntry("Error in Get Data From Baan -5");
                    string renameFile = typeFile.ToLower().Contains("COC".ToLower()) ? "COC_" : typeFile.ToLower().Contains("Invoice".ToLower()) ? "Invoice_" : "DeliveryNote_";
                    LogArchiveShipment(fullname, typeFile, 0, 2, "Error in Get Data From Baan -5");
                    movetoDefQShipment(fullname, 6, ".", renameFile, " Error in Archive the doc");
                }
                catch (Exception ex2)
                {
                    logger.WriteEntry("Error in update to shipment Get Data- " + ex2.Message);
                }
                return ErrorReason;
            }

        }

        public string archiveShipmentDoc(DataRow drWithMetadata)
        {
            SupplierShipmentDoc SuppShipDoc = null;
            if (drWithMetadata["type"].ToString() == "COC")
            {
                SuppShipDoc = new SupplierShipmentDoc(SupplierCOCDocType, SupplierCOCDoc_ShipmentDocNo);
            }
            else if (drWithMetadata["type"].ToString() == "Invoice")
            {
                SuppShipDoc = new SupplierShipmentDoc(SupplierInvoiceDocType, SupplierInvoiceDoc_ShipmentDocNo);
            }
            else
            {
                SuppShipDoc = new SupplierShipmentDoc(SupplierDeliveryNoteDocType, SupplierDeliveryNoteDoc_ShipmentDocNo);
            }
            FlexSupplier supp = new FlexSupplier(flexSupplierBinderType, flexSupplierBinder_SupplierID);
            SuppShipDoc.Filename = (string)drWithMetadata["Filename"];
            SuppShipDoc.Company = (string)drWithMetadata["CompanyNo"];
            SuppShipDoc.LotNo = (string)drWithMetadata["lotsNew"];
            SuppShipDoc.Makat = (string)drWithMetadata["MakatsNew"];
            SuppShipDoc.FilenameWithoutExt = (string)drWithMetadata["type"];
            SuppShipDoc.SupplierID = (string)drWithMetadata["supplierID"];
            SuppShipDoc.ShippmentNo = (string)drWithMetadata["shippmentNo"];
            supp.SupplierNo = (string)drWithMetadata["supplierID"];
            supp.CompanyNo = (string)drWithMetadata["CompanyNo"];
            supp.SupplierName = (string)drWithMetadata["SupplierName"];
            supp.SupplierEmail = (string)drWithMetadata["SupplierEmail"];
            supp.SupplierPhone = (string)drWithMetadata["SupplierPhone"];
            supp.SupplierAddress = (string)drWithMetadata["SupplierAddress"];
            supp.SearchKey = (string)drWithMetadata["SearchKey"];
            supp.PurchasingPerson = (string)drWithMetadata["PurchasingPerson"];
            string response = updateOrCreateSuppBinder(supp);
            DOXAPI.Binder supplierBinder = supp.asIDBinder();
            try
            {
                DOXAPI.TreeItemWithDocType SuppBinder = dox.GetTreeItemWithDocType(token, supp.asFetchItem());
                DOXAPI.Divider divShip = new DOXAPI.Divider();
                SupplierShipmentDoc ShipDoc = new SupplierShipmentDoc(ShipmentDocType, ShipmentDoc_ShipmentDocNo);
                ShipDoc.Filename = (string)drWithMetadata["Filename"];
                ShipDoc.Company = (string)drWithMetadata["CompanyNo"];
                ShipDoc.LotNo = (string)drWithMetadata["lotsNew"];
                ShipDoc.Makat = (string)drWithMetadata["MakatsNew"];
                ShipDoc.FilenameWithoutExt = (string)drWithMetadata["type"];
                ShipDoc.SupplierID = (string)drWithMetadata["supplierID"];
                ShipDoc.ShippmentNo = (string)drWithMetadata["shippmentNo"];
                divShip = ShipDoc.asDivider();
                try
                {

                    //    DOXAPI.TreeItemWithDocType divShipExists = dox.GetTreeItemWithDocType(token, supp.asFetchItem());//
                    try
                    {
                       
                        bool isExists = false;

                        foreach (var item in SuppBinder.Children)
                        {
                            if (item.Title == "Shipments")
                            {
                                using (IDataSupplier dataSupplier = DataManager.GetDataSupplier(DataManager.defaultType, Settings.Default.ConnectionString))
                                {
                                    dataSupplier.OpenQuery();
                                    dataSupplier.AddParameter("parent_id", item.ID);
                                    dataSupplier.AddParameter("title", SuppShipDoc.ShippmentNo);
                                    DataSet ds = dataSupplier.GetData("select id from tree_tab where parent_id=@parent_id and title=@title and cfg_id is not null");
                                    if (ds != null && ds.Tables.Count > 0)
                                    {
                                        DataTable dt = ds.Tables[0];
                                        if (dt != null && dt.Rows.Count > 0)
                                        {
                                            foreach (DataRow dr in dt.Rows)
                                            {
                                                if (dr != null)
                                                {
                                                    isExists = true;
                                                }
                                            }
                                        }

                                    }
                                }
                            }
                        }
                        if (!isExists)
                        {
                            try
                            {
                                logger.WriteEntry("Create divider - " + dox.CreateDivider(token, divShip, (int)SuppBinder.ID, "Shipments", null));
                            }
                            catch (Exception ex)
                            {
                                try
                                {
                                    using (IDataSupplier dataSupplier = DataManager.GetDataSupplier(DataManager.defaultType, Settings.Default.ConnectionString))
                                    {
                                        dataSupplier.OpenQuery();
                                        dataSupplier.AddParameter("FileName", drWithMetadata["Filename"].ToString());
                                        dataSupplier.AddParameter("dateGetFile", Convert.ToDateTime(drWithMetadata["dateGetFile"].ToString()).ToString("MM-dd-yyyy HH:mm:ss"));
                                        dataSupplier.AddParameter("status", "-2");
                                        dataSupplier.AddParameter("ErrorReason", ex.Message);
                                        dataSupplier.Execute("update [shipmentArchive] set status=@status,ErrorReason=@ErrorReason where FileName=@FileName and dateGetFile=@dateGetFile and status='1'");
                                    }
                                    logger.WriteEntry("Error in Archive the doc");
                                    string renameFile = SuppShipDoc.FilenameWithoutExt.ToLower().Contains("COC".ToLower()) ? "COC_" : SuppShipDoc.FilenameWithoutExt.ToLower().Contains("Invoice".ToLower()) ? "Invoice_" : "DeliveryNote_";
                                    LogArchiveShipment(SuppShipDoc.Filename, SuppShipDoc.FilenameWithoutExt, 0, 3, response);
                                    movetoDefQShipment(SuppShipDoc.Filename, 6, ".", renameFile, " Error in Archive the doc");
                                    return "not success";
                                }
                                catch (Exception ex2)
                                {
                                    logger.WriteEntry("Error in  update to shipmentArchive- " + ex2.Message);
                                }
                                return "not success";

                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        logger.WriteEntry("Error in  update to shipmentArchive- " + ex.Message);
                    }

                }
                catch (Exception ex)
                {
                    logger.WriteEntry("fail Create divider - " + ex.Message);
                }
            }
            catch (Exception ex)
            {
                logger.WriteEntry("fail  SuppShipDoc.asDivider(); - " + ex.Message);
            }
            DOXAPI.Document docSuppShip = SuppShipDoc.asDocument();
            //Archive a new shipment in DOX-Pro
            doxLogin();
            response = dox.Archive(token, docSuppShip, supplierBinder, "Shipments\\" + SuppShipDoc.ShippmentNo, false);
            logger.WriteEntry("response - " + response);
            long docID;
            string ErrorReason = "";
            if (long.TryParse(response, out docID))
            {
               
                logger.WriteEntry("Shipment Doc Document no. " + SuppShipDoc.ShippmentNo + " archived with ID=" + docID);
                System.IO.File.Delete(SuppShipDoc.Filename);
                try
                {
                    using (IDataSupplier dataSupplier = DataManager.GetDataSupplier(DataManager.defaultType, Settings.Default.ConnectionString))
                    {
                        dataSupplier.OpenQuery();
                        dataSupplier.AddParameter("FileName", drWithMetadata["Filename"].ToString());
                        dataSupplier.AddParameter("dateGetFile", Convert.ToDateTime(drWithMetadata["dateGetFile"].ToString()).ToString("MM-dd-yyyy HH:mm:ss"));
                        dataSupplier.AddParameter("status", "2");
                        dataSupplier.Execute("update [shipmentArchive] set status=@status where FileName=@FileName and dateGetFile=@dateGetFile and status='1'");
                    }
                }
                catch (Exception ex)
                {
                    logger.WriteEntry("Error in  update to shipmentArchive- " + ex.Message);
                    ErrorReason = ex.Message;
                }
            }
            else
            {
                try
                {
                    using (IDataSupplier dataSupplier = DataManager.GetDataSupplier(DataManager.defaultType, Settings.Default.ConnectionString))
                    {
                        dataSupplier.OpenQuery();
                        dataSupplier.AddParameter("FileName", drWithMetadata["Filename"].ToString());
                        dataSupplier.AddParameter("dateGetFile", Convert.ToDateTime(drWithMetadata["dateGetFile"].ToString()).ToString("MM-dd-yyyy HH:mm:ss"));
                        dataSupplier.AddParameter("status", "-2");
                        dataSupplier.AddParameter("ErrorReason", ErrorReason);
                        dataSupplier.Execute("update [shipmentArchive] set status=@status,ErrorReason=@ErrorReason where FileName=@FileName and dateGetFile=@dateGetFile and status='1'");
                    }
                    logger.WriteEntry("Error in Archive the doc");
                    string renameFile = SuppShipDoc.FilenameWithoutExt.ToLower().Contains("COC".ToLower()) ? "COC_" : SuppShipDoc.FilenameWithoutExt.ToLower().Contains("Invoice".ToLower()) ? "Invoice_" : "DeliveryNote_";
                    LogArchiveShipment(SuppShipDoc.Filename, SuppShipDoc.FilenameWithoutExt, 0, 3, response);
                    movetoDefQShipment(SuppShipDoc.Filename, 6, ".", renameFile, " Error in Archive the doc");
                    return "not success";
                }
                catch (Exception ex)
                {
                    logger.WriteEntry("Error in  update to shipmentArchive- " + ex.Message);
                }
                return "not success";
            }
            LogArchiveShipment(SuppShipDoc.Filename, SuppShipDoc.FilenameWithoutExt, 1, 0);
            return "";
        }
    }
}