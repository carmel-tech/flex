using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using System.IO;

namespace BussinessLogicLayer
{
    class FlexOrder99
    {
        private string companyID = String.Empty;
        private string order99No;
        private string supplierID;
        private string filename;

        private DOXAPI.DocType flexOrder99Type;
        private DOXAPI.DocTypeAttribute flexOrder99_Order99No;


        public FlexOrder99(DOXAPI.DocType Order99Type, DOXAPI.DocTypeAttribute Order99_Order99No)
        {
            flexOrder99Type = Order99Type;
            flexOrder99_Order99No = Order99_Order99No;
        }
        public string Company
        {
            get { return companyID; }
            set { companyID = value; }
        }
        public string Order99No
        {
            get { return order99No; }
            set { order99No = value; }
        }
        public string SupplierID
        {
            get { return supplierID; }
            set { supplierID = value; }
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
            t.DocType = flexOrder99Type;
            DOXAPI.Field f = new DOXAPI.Field();
            f.Attr = flexOrder99_Order99No;
            f.Value = Order99No;
            t.Fields = new DOXAPI.Field[1];
            t.Fields[0] = f;

            return t;
        }

        public DOXAPI.Document asDocument()
        {
            // Convert local class object to DOX-Pro document object
            DOXAPI.Document docOrd99 = new DOXAPI.Document();
            docOrd99.DocType = flexOrder99Type;
            // Set documents fields according to doc-type
            docOrd99.Fields = new DOXAPI.Field[flexOrder99Type.Attributes.Length];
            for (int i = 0; i < docOrd99.Fields.Length; i++)
            {
                DOXAPI.Field f = new DOXAPI.Field();
                f.Attr = flexOrder99Type.Attributes[i];
                switch (f.Attr.Name)
                {

                    case "Order99No":
                        f.Value = Order99No;
                        break;
                    case "Supplier No":
                        f.Value = SupplierID;
                        break;

                    default:
                        f.Value = null;
                        break;
                }
                docOrd99.Fields[i] = f;

            }


            docOrd99.Title = supplierID + "/" + Order99No;
            docOrd99.FileName = filename;
            return docOrd99;
        }
        public DOXAPI.Binder asBinder()
        {
            // Convert local class object to DOX-Pro binder object
            DOXAPI.Binder binder = new DOXAPI.Binder();
            binder.DocType = flexOrder99Type;
            binder.Title = supplierID + "/" + Order99No; ;
            // Create and set binder fileds according to its doc-type
            binder.Fields = new DOXAPI.Field[flexOrder99Type.Attributes.Length];

            for (int i = 0; i < flexOrder99Type.Attributes.Length; i++)
            {
                DOXAPI.Field f = new DOXAPI.Field();
                f.Attr = flexOrder99Type.Attributes[i];
                switch (f.Attr.Name)
                {
                    case "Order99No":
                        f.Value = Order99No;
                        break;
                    case "SupplierID":
                        f.Value = supplierID;
                        break;
                    default:
                        f.Value = null;
                        break;
                }
                binder.Fields[i] = f;
            }
            return binder;
        }
    }
}

