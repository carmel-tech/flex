using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.Odbc;


namespace BussinessLogicLayer
{
    class SupplierInvoice
    {
        private string company = string.Empty;
        private string invoiceNo = "0";
        private string supplierID;
        private string filename;

        public string InvoiceNo
        {
            get { return invoiceNo; }
            set { invoiceNo = value; }
        }
        public string SupplierID
        {
            get { return supplierID; }
            set { supplierID = value; }
        }
        public string Company
        {
            get { return company; }
            set { company = value; }
        }
       
        public string Filename
        {
            get { return filename; }
            set { filename = value; }
        }

        private DOXAPI.DocType SuppInvoiceDocType;
        private DOXAPI.DocTypeAttribute SuppInvoiceAtt_SuppInvoiceDocNo;

        public SupplierInvoice(DOXAPI.DocType SupplierInvoiceDocType, DOXAPI.DocTypeAttribute SupplierInvoiceDocNoField)
        {
            SuppInvoiceDocType = SupplierInvoiceDocType;
            SuppInvoiceAtt_SuppInvoiceDocNo = SupplierInvoiceDocNoField;
        }


        public DOXAPI.Document asDocument()
        {
            // Convert local class object to DOX-Pro document object
            DOXAPI.Document SupplierInvoiceDoc = new DOXAPI.Document();
            SupplierInvoiceDoc.DocType = SuppInvoiceDocType;
            // Set documents fields according to doc-type
            SupplierInvoiceDoc.Fields = new DOXAPI.Field[SuppInvoiceDocType.Attributes.Length];
            for (int i = 0; i < SupplierInvoiceDoc.Fields.Length; i++)
            {
                DOXAPI.Field f = new DOXAPI.Field();
                f.Attr = SuppInvoiceDocType.Attributes[i];
                System.Diagnostics.EventLog.WriteEntry("Supplier Invoice- as docoument", "f.Attr.Name: " + f.Attr.Name, System.Diagnostics.EventLogEntryType.Information);
                switch (f.Attr.Name)
                {

                    case "Invoice No":
                        f.Value = invoiceNo;
                        break;
                    case "Supplier No":
                        f.Value = SupplierID;
                        break;
                    default:
                        f.Value = null;
                        break;
                }
                SupplierInvoiceDoc.Fields[i] = f;

            }

            SupplierInvoiceDoc.Title = supplierID + "/" + invoiceNo;
            SupplierInvoiceDoc.FileName = filename;
            return SupplierInvoiceDoc;
        }




    }
}
