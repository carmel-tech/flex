using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.Odbc;

namespace BussinessLogicLayer
{
    class FlexCustomer
    {
        private string companyID = String.Empty;
        private string clientID = String.Empty;
        private string clientName = String.Empty;
        private string clientAddress = String.Empty;
        private string flexProjectManager = String.Empty;
        private string flexController = String.Empty;
        private DOXAPI.DocType flexClientBinderType;
        private DOXAPI.DocTypeAttribute flexClientBinder_customerID;

        public FlexCustomer(DOXAPI.DocType customerBinderType, DOXAPI.DocTypeAttribute customerIDField)
        {
            flexClientBinderType = customerBinderType;
            flexClientBinder_customerID = customerIDField;
        }

        public string Company
        {
            get { return companyID; }
            set { companyID = value; }
        }
        public string ClientID
        {
            get { return clientID; }
            set { clientID = value; }
        }
        public string ClientName
        {
            get { return clientName; }
            set { clientName = value; }
        }
        public string ClientAddress
        {
            get { return clientAddress; }
            set { clientAddress = value; }
        }
        public string FullID
        {
            get { return companyID + clientID; }
        }
        public string ProjectManager
        {
            get { return flexProjectManager; }
            set { flexProjectManager = value; }
        }
        public string Controller
        {
            get { return flexController; }
            set { flexController = value; }
        }

        public bool refreshCustomerDetails(OdbcConnection DbConnection, Dictionary<string, string> doxParams)
        {
            string pmr=string.Empty;
            string ctr=string.Empty;
            bool wasChanged = false;

            OdbcCommand DbCommand = DbConnection.CreateCommand();
            DbCommand.CommandText = String.Format(doxParams["CustPeopleQ"], Company, ClientID);
            OdbcDataReader DbReader = DbCommand.ExecuteReader();
            while (DbReader.Read())
            {
                String nama = DbReader.GetString(1).Trim();
                String telp = DbReader.GetString(2).Trim();
                String code = DbReader.GetString(3).Trim();
                if (code == "PMR") pmr += nama + " " + telp + ";";
                if (code == "CTR") ctr += nama + " " + telp + ";";
            }

            DbReader.Close();
            DbCommand.Dispose();

            wasChanged = (pmr != ProjectManager || ctr != Controller);
            Controller = ctr;
            ProjectManager = pmr;
            return wasChanged;
        }

        public DOXAPI.Binder asIDBinder()
        {
            DOXAPI.Binder b = new DOXAPI.Binder();
            b.DocType = flexClientBinderType;
            b.Fields = new DOXAPI.Field[1];
            b.Fields[0] = DOXFields.newField("Customer ID", flexClientBinder_customerID, FullID);
            return b;
        }

        public DOXAPI.Binder asBinder()
        {
            // Convert local class object to DOX-Pro binder object
            DOXAPI.Binder binder = new DOXAPI.Binder();
            binder.DocType = flexClientBinderType;
            binder.Title = clientName;
            // Create and set binder fileds according to its doc-type
            binder.Fields = new DOXAPI.Field[flexClientBinderType.Attributes.Length];

            for (int i = 0; i < flexClientBinderType.Attributes.Length; i++)
            {
                DOXAPI.Field f = new DOXAPI.Field();
                f.Attr = flexClientBinderType.Attributes[i];
                switch (f.Attr.Name)
                {
                    case "Customer ID":
                        f.Value = FullID;
                        break;
                    case "Name":
                        f.Value = ClientName;
                        break;
                    case "Address":
                        f.Value = ClientAddress;
                        break;
                    case "Project Manager":
                        f.Value = ProjectManager;
                        break;
                    case "Controller":
                        f.Value = Controller;
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
            t.DocType = flexClientBinderType;
            DOXAPI.Field f = new DOXAPI.Field();
            f.Attr = flexClientBinder_customerID;
            f.Value = FullID;
            t.Fields = new DOXAPI.Field[1];
            t.Fields[0] = f;
            return t;
        }

        public String updateCustomerFields(DOXAPI.TreeItemWithDocType customerBinder)
        {
            string debugStep = "k10-0";
            try
            {
                customerBinder.Title = ClientName;
                debugStep = "k10-1";
                DOXFields.SetField(customerBinder, "Name", ClientName);
                debugStep = "k10-2";
                DOXFields.SetField(customerBinder, "Address", ClientAddress);
                debugStep = "k10-3";
                DOXFields.SetField(customerBinder, "Project Manager", ProjectManager);
                debugStep = "k10-4";
                DOXFields.SetField(customerBinder, "Controller", Controller);
                debugStep = "k10-5";
                return "";
            }
            catch (Exception ex)
            {
                return debugStep + " " + ex.Message;
            }
        }
    }
}
