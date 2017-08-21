using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.Odbc;

namespace BussinessLogicLayer
{
    class FlexSupplier
    {

        private string supplierNo = String.Empty;
        private string supplierName = String.Empty;
        private string companyNo = String.Empty;
        private string supplierAddress = String.Empty;
        private string supplierRegion = String.Empty;
        private string supplierType = String.Empty;
        private string purchasingPerson = String.Empty;
        private string supplierContact = String.Empty;
        private string supplierPhone = String.Empty;
        private string supplierEmail = String.Empty;
        private string searchKey = String.Empty;
        private string comment = String.Empty;

        private DOXAPI.DocType flexSupplierBinderType;
        private DOXAPI.DocTypeAttribute flexSupplierBinder_supplierID;

        public FlexSupplier(DOXAPI.DocType supplierBinderType, DOXAPI.DocTypeAttribute SupplierIDField)
        {
            flexSupplierBinderType = supplierBinderType;
            flexSupplierBinder_supplierID = SupplierIDField;
        }
        public string SupplierNo
        {
            get { return supplierNo; }
            set { supplierNo = value; }
        }
        public string SupplierName
        {
            get { return supplierName; }
            set { supplierName = value; }
        }
        public string CompanyNo
        {
            get { return companyNo; }
            set { companyNo = value; }
        }
        public string SupplierAddress
        {
            get { return supplierAddress; }
            set { supplierAddress = value; }
        }
        public string SupplierRegion
        {
            get { return supplierRegion; }
            set { supplierRegion = value; }
        }
        public string SupplierType
        {
            get { return supplierType; }
            set { supplierType = value; }
        }
        public string PurchasingPerson
        {
            get { return purchasingPerson; }
            set { purchasingPerson = value; }
        }
        public string SupplierContact
        {
            get { return supplierContact; }
            set { supplierContact = value; }
        }

        public string SupplierEmail
        {
            get { return supplierEmail; }
            set { supplierEmail = value; }
        }
        public string SupplierPhone
        {
            get { return supplierPhone; }
            set { supplierPhone = value; }
        }
        public string SearchKey
        {
            get { return searchKey; }
            set { searchKey = value; }
        }
        public string Comment
        {
            get { return comment; }
            set { comment = value; }
        }
        public bool GetSupplierDetails(OdbcConnection DbConnection, Dictionary<string, string> doxParams)
        {
            string pmr = string.Empty;
            string ctr = string.Empty;


            OdbcCommand DbCommand = DbConnection.CreateCommand();
            DbCommand.CommandText = String.Format(doxParams["SuppQ"], CompanyNo, SupplierNo);
            System.Diagnostics.EventLog.WriteEntry("SuppQ", DbCommand.CommandText, System.Diagnostics.EventLogEntryType.Information);
            OdbcDataReader DbReader = DbCommand.ExecuteReader();
            System.Diagnostics.EventLog.WriteEntry("SuppQ", "1", System.Diagnostics.EventLogEntryType.Information);
            if (DbReader.Read())
            {
                System.Diagnostics.EventLog.WriteEntry("SuppQ", "2", System.Diagnostics.EventLogEntryType.Information);
                SupplierName = DbReader.GetString(0).Trim();//nama
                SupplierAddress = DbReader.GetString(1).Trim();//namc
                SupplierContact = DbReader.GetString(2).Trim();//refs
                SupplierPhone = DbReader.GetString(3).Trim();//telp
                SearchKey = DbReader.GetString(4).Trim();//seak

            }
            
            DbReader.Close();
            DbCommand.Dispose();
            //now get supplier email
            DbCommand = DbConnection.CreateCommand();
            DbCommand.CommandText = String.Format(doxParams["SuppEmailQ"], CompanyNo, SupplierNo);
            System.Diagnostics.EventLog.WriteEntry("SuppEmailQ", doxParams["SuppEmailQ"], System.Diagnostics.EventLogEntryType.Information);
            DbReader = DbCommand.ExecuteReader();
            if (DbReader.Read())
            {
                SupplierEmail = DbReader.GetString(0).Trim();//email
            }
            System.Diagnostics.EventLog.WriteEntry("SupplierEmail", SupplierEmail, System.Diagnostics.EventLogEntryType.Information);
            return true;
        }

        public DOXAPI.Binder asIDBinder()
        {
            DOXAPI.Binder b = new DOXAPI.Binder();
            b.DocType = flexSupplierBinderType;
            b.Fields = new DOXAPI.Field[1];
            b.Fields[0] = DOXFields.newField("Supplier No", flexSupplierBinder_supplierID, SupplierNo);
            return b;
        }

        public DOXAPI.Binder asBinder()
        {
            // Convert local class object to DOX-Pro binder object
            DOXAPI.Binder binder = new DOXAPI.Binder();
            binder.DocType = flexSupplierBinderType;
            binder.Title = SupplierName;
            // Create and set binder fileds according to its doc-type
            binder.Fields = new DOXAPI.Field[flexSupplierBinderType.Attributes.Length];

            for (int i = 0; i < flexSupplierBinderType.Attributes.Length; i++)
            {
                DOXAPI.Field f = new DOXAPI.Field();
                f.Attr = flexSupplierBinderType.Attributes[i];
                switch (f.Attr.Name)
                {

                    case "Supplier No":
                        f.Value = supplierNo;
                        break;
                    case "Name":
                        f.Value = supplierName;
                        break;
                    case "Address":
                        f.Value = supplierAddress;
                        break;
                    case "Purchasing Person":
                        f.Value = purchasingPerson;
                        break;
                    case "Contact":
                        f.Value = supplierContact;
                        break;
                    default:
                        f.Value = null;
                        break;
                }
                binder.Fields[i] = f;
            }
            return binder;
        }

        public DOXAPI.TreeItemWithDocType asFetchItem()
        {
            // Create DOX-Pro object with key DocType
            DOXAPI.TreeItemWithDocType t = new DOXAPI.TreeItemWithDocType();
            t.DocType = flexSupplierBinderType;
            DOXAPI.Field f = new DOXAPI.Field();
            f.Attr = flexSupplierBinder_supplierID;
            f.Value = SupplierNo;
            t.Fields = new DOXAPI.Field[1];
            t.Fields[0] = f;
            return t;
        }

        public String updateSupplierFields(DOXAPI.TreeItemWithDocType SupplierBinder)
        {
            string debugStep = "k10-0";
            try
            {
                SupplierBinder.Title = SupplierName;
                debugStep = "k10-1";
                DOXFields.SetField(SupplierBinder, "Suppliers name", SupplierName);
                debugStep = "k10-2";
                DOXFields.SetField(SupplierBinder, "Address", SupplierAddress);
                debugStep = "k10-3";
                DOXFields.SetField(SupplierBinder, "Purchasing person", SupplierContact);
                debugStep = "k10-4";
                DOXFields.SetField(SupplierBinder, "Phone no", SupplierPhone);
                debugStep = "k10-5";
                DOXFields.SetField(SupplierBinder, "Email address", SupplierEmail);
                debugStep = "k10-6";
                DOXFields.SetField(SupplierBinder, "Search key", SearchKey);
                debugStep = "k10-7";
                return "";
            }
            catch (Exception ex)
            {
                return debugStep + " " + ex.Message;
            }
        }
    }
}
