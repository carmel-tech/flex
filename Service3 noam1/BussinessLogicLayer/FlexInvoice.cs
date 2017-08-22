using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BussinessLogicLayer
{
    class FlexInvoice
    {
        private string company=string.Empty;
        private int invoiceNo = 0;
        private string packingSlips=String.Empty;
        private DateTime issueDate=DateTime.MinValue;
        private DOXAPI.DocType flexInvoiceType;
        private string filename;
        private string clientID;

        public FlexInvoice(DOXAPI.DocType invoiceType)
        {
            flexInvoiceType = invoiceType;
        }

        public bool good
        {
            get {
                if (invoiceNo == 0) return false;
                //if (packingSlips == String.Empty) return false;
                if (issueDate==DateTime.MinValue) return false;
                return true;
            }
        }

        public int InvoiceNo
        {
            get { return invoiceNo; }
            set { invoiceNo = value; }
        }
        public string Company
        {
            get { return company; }
            set { company = value; }
        }
        public string PackingSlips
        {
            get { return packingSlips; }
            set { packingSlips = value; }
        }
        public DateTime IssueDate
        {
            get { return issueDate; }
            set { issueDate = value; }
        }
        public string Filename
        {
            get { return filename; }
            set { filename = value; }
        }
        public string fullInvNo
        {
            get { return company + invoiceNo; }
        }
        public string ClientID
        {
            get { return clientID; }
            set { clientID = value; }
        }
        public DOXAPI.Document asDocument(string customerID, string customerName, string filename)
        {
            // Convert local class object to DOX-Pro document object
            DOXAPI.Document docInv = new DOXAPI.Document();
            docInv.DocType = flexInvoiceType;
            // Set documents fields according to doc-type
            docInv.Fields = new DOXAPI.Field[flexInvoiceType.Attributes.Length];
            for (int i = 0; i < docInv.Fields.Length; i++)
            {
                DOXAPI.Field f = new DOXAPI.Field();
                f.Attr = flexInvoiceType.Attributes[i];
                switch (f.Attr.Name)
                {
                    case "Customer ID":
                        f.Value = customerID;
                        break;
                    case "Invoice No":
                        f.Value = company + InvoiceNo;
                        break;
                    case "Packing Slips":
                        f.Value = removeDuplicates(PackingSlips);
                        break;
                    case "Issue Date":
                        f.Value = IssueDate;
                        break;
                    default:
                        f.Value = null;
                        break;
                }
                docInv.Fields[i] = f;
            }
            docInv.Title = InvoiceNo + "/" + IssueDate.ToShortDateString();
            docInv.FileName = filename;
            return docInv;
        }

        private string removeDuplicates(string psList)
        {
            string shortList = "";
            string[] pss = psList.Split(',');
            foreach (string ps in pss)
            {
                if (shortList.IndexOf(ps) == -1)
                    shortList += ps + ',';
            }
            return shortList.TrimEnd(',');
        }
    }
}
