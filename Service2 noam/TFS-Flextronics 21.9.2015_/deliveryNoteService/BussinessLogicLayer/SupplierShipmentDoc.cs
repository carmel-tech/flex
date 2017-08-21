using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.Odbc;


namespace BussinessLogicLayer
{
    class SupplierShipmentDoc
    {
        private string company = string.Empty;
        private string shippmentNo;
        private string supplierID;
        private string lotNo = String.Empty;
        private string makat = String.Empty;
        private DateTime issueDate = DateTime.MinValue;
        private string filename;


        public string ShippmentNo
        {
            get { return shippmentNo; }
            set { shippmentNo = value; }
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
        public string LotNo
        {
            get { return lotNo; }
            set { lotNo = value; }
        }
        public string Makat
        {
            get { return makat; }
            set { makat = value; }
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

        private DOXAPI.DocType SuppShipmentDocType;
        private DOXAPI.DocTypeAttribute SuppShipmentAtt_ShipmentDocNo;

        public SupplierShipmentDoc(DOXAPI.DocType ShipmentDocType, DOXAPI.DocTypeAttribute ShipmentDocNoField)
        {
            SuppShipmentDocType = ShipmentDocType;
            SuppShipmentAtt_ShipmentDocNo = ShipmentDocNoField;
        }


        public DOXAPI.Document asDocument()
        {
            // Convert local class object to DOX-Pro document object
            DOXAPI.Document ShipmentDoc = new DOXAPI.Document();
            ShipmentDoc.DocType = SuppShipmentDocType;
            // Set documents fields according to doc-type
            ShipmentDoc.Fields = new DOXAPI.Field[SuppShipmentDocType.Attributes.Length];
            for (int i = 0; i < ShipmentDoc.Fields.Length; i++)
            {
                DOXAPI.Field f = new DOXAPI.Field();
                f.Attr = SuppShipmentDocType.Attributes[i];
                System.Diagnostics.EventLog.WriteEntry("shipment- as docoument", "f.Attr.Name: " + f.Attr.Name, System.Diagnostics.EventLogEntryType.Information);
                switch (f.Attr.Name)
                {

                    case "Shipment Doc No":
                        f.Value = ShippmentNo;
                        break;
                    case "Supplier No":
                        f.Value = SupplierID;
                        break;
                    default:
                        f.Value = null;
                        break;
                }
                ShipmentDoc.Fields[i] = f;

            }

            ShipmentDoc.Title = supplierID + "/" + ShippmentNo;
            ShipmentDoc.FileName = filename;
            return ShipmentDoc;
        }
        public int GetShippmentNo(OdbcConnection DbConnection, string company, string Lot, string Makat)
        {
            //get order->orno line->pono  DateIn-> trdt 
            string OrderNo = String.Empty;
            string OrderLine = String.Empty;
            string DateIn = String.Empty;
            int ShipNo = 0;
            OdbcCommand DbCommand = DbConnection.CreateCommand();

            string cmd = "select orno, pono, trdt  from baandb.tdltc102{0} where tdltc102.tdltc102=1 and tdltc102.clot=  and tdltc102.item =\"\"";
            DbCommand.CommandText = String.Format(cmd, company, Lot, Makat);
            OdbcDataReader reader;

            reader = DbCommand.ExecuteReader();


            if (reader.Read())
            {
                OrderNo = reader.GetString(0);
                OrderLine = reader.GetString(1);
                DateIn = reader.GetString(2);

            }
            else
            {
                return ShipNo;
            }

            reader.Close();

            cmd = "select dino from baandb.tdpur045{0}  where orno={1} and pono={2}  and date  ={3}";
            DbCommand.CommandText = String.Format(cmd, company, OrderNo, OrderLine, DateIn);

            reader = DbCommand.ExecuteReader();

            if (reader.Read())
            {
                ShipNo = System.Convert.ToInt32(reader.GetString(0));


            }
            return ShipNo;


        }



    }
}
