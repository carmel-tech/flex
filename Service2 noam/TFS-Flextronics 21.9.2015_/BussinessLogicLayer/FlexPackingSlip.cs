using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BussinessLogicLayer
{
    class FlexPackingSlip
    {
        private string company;
        private string psno;
        private string customerID;
        private DateTime issueDate;
        private string itemList;
        private string customerOrderNo;
        private string filename;

        private DOXAPI.DocType flexPackingSlipType;
        private DOXAPI.DocTypeAttribute flexPackSlip_PackSlipNo;

        public FlexPackingSlip(DOXAPI.DocType packingSlipType, DOXAPI.DocTypeAttribute PSNoType)
        {
            flexPackingSlipType = packingSlipType;
            flexPackSlip_PackSlipNo = PSNoType;
        }

        public string PackingSlipNo
        {
            get { return psno; }
            set { psno = value; }
        }
        public string CustomerID
        {
            get { return customerID; }
            set { customerID = value; }
        }
        public DateTime IssueDate
        {
            get { return issueDate; }
            set { issueDate = value; }
        }
        public string CustomerOrderNo
        {
            get { return customerOrderNo; }
            set { customerOrderNo = value; }
        }
        public string ItemList
        {
            get { return itemList; }
            set { itemList = value; }
        }
        public string Company
        {
            get { return company; }
            set { company = value; }
        }
        public string FullID
        {
            get { return company + psno; }
        }
        public string Filename
        {
            get { return filename; }
            set { filename = value; }
        }

        public DOXAPI.TreeItemWithDocType asFetchItem()
        {
            // Create DOX-Pro object with key DocType from local object
            DOXAPI.TreeItemWithDocType t = new DOXAPI.TreeItemWithDocType();
            t.DocType = flexPackingSlipType;
            DOXAPI.Field f = new DOXAPI.Field();
            f.Attr = flexPackSlip_PackSlipNo;
            f.Value = FullID;
            t.Fields = new DOXAPI.Field[1];
            t.Fields[0] = f;

            return t;
        }

        public DOXAPI.Document asDocument()
        {
            // Convert local class object to DOX-Pro document object
            DOXAPI.Document docPS = new DOXAPI.Document();
            docPS.DocType = flexPackingSlipType;
            // Set documents fields according to doc-type
            docPS.Fields = new DOXAPI.Field[flexPackingSlipType.Attributes.Length];
            for (int i = 0; i < docPS.Fields.Length; i++)
            {
                DOXAPI.Field f = new DOXAPI.Field();
                f.Attr = flexPackingSlipType.Attributes[i];
                switch (f.Attr.Name)
                {
                    case "Packing Slip No":
                        f.Value = FullID;
                        break;
                    case "Customer ID":
                        f.Value = Company + CustomerID;
                        break;
                    case "Customer Order No":
                        f.Value = CustomerOrderNo;
                        break;
                    case "Issue Date":
                        f.Value = IssueDate;
                        break;
                    case "Items":
                        if (itemList.Length > 255)
                            f.Value = ItemList.Substring(0, 255);
                        else
                            f.Value = ItemList;
                        break;
                    case "Return Date":
                        f.Value = DateTime.Now;
                        break;
                    case "Invoice No":
                        f.Value = ""; // This item is set later
                        break;
                    default:
                        f.Value = null;
                        break;
                }
                docPS.Fields[i] = f;
            }
            docPS.Title = CustomerID + "/" + PackingSlipNo + "/" + DateTime.Now.ToShortDateString();
            docPS.FileName = filename;
            return docPS;
        }
    }
}
