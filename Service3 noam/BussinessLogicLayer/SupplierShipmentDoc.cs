using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.Odbc;
using System.Diagnostics;


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
        //Ayala 13.04.2015
        private string filenameWithoutExt;
        public string FilenameWithoutExt
        {
            get { return filenameWithoutExt; }
            set { filenameWithoutExt = value; }
        }
        //End Ayala 13.04.2015
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
                    //Ayala 13.04.2015
                    case "Lot":
                        f.Value = LotNo;
                        break;
                    case "Item No":
                        f.Value = Makat;
                        break;
                    case "Date":
                        f.Value = DateTime.Now.ToString("MM-dd-yyyy HH:mm:ss");
                        break;
                    //End Ayala 13.04.2015
                    default:
                        f.Value = null;
                        break;
                }
                ShipmentDoc.Fields[i] = f;

            }

           // ShipmentDoc.Title = supplierID + "/" + ShippmentNo;
            //Ayala 13.04.2015
            ShipmentDoc.Title = filenameWithoutExt;
            //End Ayala 13.04.2015
            ShipmentDoc.FileName = filename;
            return ShipmentDoc;
        }
        //public int GetShippmentNo(OdbcConnection DbConnection, string company, string Lot, string Makat)
        //{
        //    //get order->orno line->pono  DateIn-> trdt 
        //    string OrderNo = String.Empty;
        //    string OrderLine = String.Empty;
        //    string DateIn = String.Empty;
        //    int ShipNo = 0;
        //    OdbcCommand DbCommand = DbConnection.CreateCommand();

        //    string cmd = "select orno, pono, trdt  from baandb.tdltc102{0} where tdltc102.tdltc102=1 and tdltc102.clot=  and tdltc102.item =\"\"";
        //    DbCommand.CommandText = String.Format(cmd, company, Lot, Makat);
        //    OdbcDataReader reader;

        //    reader = DbCommand.ExecuteReader();


        //    if (reader.Read())
        //    {
        //        OrderNo = reader.GetString(0);
        //        OrderLine = reader.GetString(1);
        //        DateIn = reader.GetString(2);

        //    }
        //    else
        //    {
        //        return ShipNo;
        //    }

        //    reader.Close();

        //    cmd = "select dino from baandb.tdpur045{0}  where orno={1} and pono={2}  and date  ={3}";
        //    DbCommand.CommandText = String.Format(cmd, company, OrderNo, OrderLine, DateIn);

        //    reader = DbCommand.ExecuteReader();

        //    if (reader.Read())
        //    {
        //        ShipNo = System.Convert.ToInt32(reader.GetString(0));


        //    }
        //    return ShipNo;


        //}
        public DOXAPI.Divider asDivider()
        {
            // Convert local class object to DOX-Pro document object
            DOXAPI.Divider ShipmentDoc = new DOXAPI.Divider();
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
                    //Ayala 13.04.2015
                    case "Lot":
                        f.Value = LotNo;
                        break;
                    case "Item No":
                        f.Value = Makat;
                        break;
                    case "Date":
                        f.Value = DateTime.Now.ToString("MM-dd-yyyy HH:mm:ss");
                        break;
                    //End Ayala 13.04.2015
                    default:
                        f.Value = null;
                        break;
                }
                ShipmentDoc.Fields[i] = f;
            }
            // ShipmentDoc.Title = supplierID + "/" + ShippmentNo;
            //Ayala 13.04.2015
            ShipmentDoc.Title = ShippmentNo;
            //End Ayala 13.04.2015

            return ShipmentDoc;
        }
        //Ayala 13.04.2015
        public SupplierShipmentDoc GetShippmentNo(OdbcConnection DbConnection, SupplierShipmentDoc ship, FlexSupplier supp, EventLog lg)
        {

            supp.SupplierNo = "";
            supp.CompanyNo = "";
            supp.SupplierName = "";
            supp.SupplierAddress = "";
            supp.PurchasingPerson = "";
            supp.SupplierPhone = "";
            supp.SearchKey = "";
            supp.SupplierEmail = "";
            ship.shippmentNo = "";
            ship.supplierID = "";
            string OrderNo = String.Empty;
            string OrderLine = String.Empty;
            string DateIn = String.Empty;

            int ShipNo = 0;
            OdbcCommand DbCommand = DbConnection.CreateCommand();
            DbCommand.CommandTimeout = 3600;
            OdbcDataReader DbReader;

            DbCommand.CommandText = String.Format("select  b405.t_dino,b405.t_suno,b020.t_nama,b020.t_namc, b020.t_refs, b020.t_telp, b020.t_seak,b040.t_email  from baandb.ttdltc102{0} as b102, baandb.ttdpur045{0} as b405, "
                  + " baandb.ttccom020{0} as b020 , baandb.ttccom040{0} as b040 " +
                    " where b102.t_tord=1  and  b102.t_clot like '%{1}%'  and  b102.t_item like '%{2}%' and b405.t_suno=b020.t_suno  " +
                    " and b040.t_suno=b405.t_suno and  b040.t_actv=1 "
                  + " and b405.t_orno= b102.t_orno and b405.t_pono= b102.t_pono and b405.t_date= b102.t_trdt", "400", ship.lotNo.Replace(" ", ""), ship.makat.Replace(" ", ""));

            bool b = false;
            try
            {
                DbReader = DbCommand.ExecuteReader();

            }
            catch (Exception ex)
            {

                return ship;
            }

            if (DbReader.Read())
            {
                ship.shippmentNo = DbReader.GetString(0).Replace(" ", "").Trim();
                ship.supplierID = DbReader.GetString        (1).Replace(" ", "").Trim();
                supp.SupplierNo = ship.SupplierID;
                supp.CompanyNo = ship.Company;
                supp.SupplierName = DbReader.GetString(2).Trim();
                supp.SupplierAddress = DbReader.GetString(3).Trim();
                supp.PurchasingPerson = DbReader.GetString(4).Trim();
                supp.SupplierPhone = DbReader.GetString(5).Trim();
                supp.SearchKey = DbReader.GetString(6).Trim();
                supp.SupplierEmail = DbReader.GetString(7).Trim();

                b = true;
            }
            DbReader.Close();
            if (b)
            {

                DbCommand.CommandText = String.Format("select b102.t_clot  , b102.t_item   from baandb.ttdltc102{0} as b102, baandb.ttdpur045{0} as b405, "
           + " baandb.ttccom020{0} as b020  " +
             " where b102.t_tord=1  and b405.t_orno= b102.t_orno and b405.t_pono= b102.t_pono and b405.t_date= b102.t_trdt and b405.t_suno=b020.t_suno and "
           + " b405.t_dino like '%{1}%'", "400", ship.shippmentNo);


                try
                {
                    DbReader = DbCommand.ExecuteReader();

                }
                catch (Exception ex)
                {
                    return ship;
                }
                string first = ship.lotNo.Replace(" ", "");
                Dictionary<string, string> lotsAndMakat = new Dictionary<string, string>();
                List<string> Lots = new List<string>();
                List<string> Makats = new List<string>();
                Lots.Add(first);

                string firstMakat = ship.makat.Replace(" ", "");
                Makats.Add(firstMakat);
                ship.lotNo = ship.lotNo.Replace(" ", "");
                ship.makat = ship.makat.Replace(" ", "");
                bool c = false;

                lotsAndMakat.Add(first, firstMakat);
                while (DbReader.Read())
                {

                    if (first == DbReader.GetString(0).Replace(" ", "").Trim() && DbReader.GetString(1).Replace(" ", "").Trim() == firstMakat)
                        c = true;

                    else
                    {
                        if (ship.lotNo.Length < 255 && ship.makat.Length < 255)
                        {
                            if (!Lots.Contains(DbReader.GetString(0).Replace(" ", "").Trim()))
                            {
                                ship.lotNo = ship.lotNo + ("," + DbReader.GetString(0).Replace(" ", "").Trim());
                                Lots.Add(DbReader.GetString(0).Replace(" ", "").Trim());
                            }
                            if (!Makats.Contains(DbReader.GetString(1).Replace(" ", "").Trim()))
                            {
                                ship.makat = ship.makat + ("," + DbReader.GetString(1).Replace(" ", "").Trim());
                                Makats.Add(DbReader.GetString(1).Replace(" ", "").Trim());
                            }
                        }
                    }

                }

                DbReader.Close();

            }



            if (!b)
            {
                DbCommand.CommandText = String.Format("select  b405.t_dino,b405.t_suno,b020.t_nama,b020.t_namc, b020.t_refs, b020.t_telp, b020.t_seak  from baandb.ttdltc102{0} as b102, baandb.ttdpur045{0} as b405, "
         + " baandb.ttccom020{0} as b020  " +
           " where b102.t_tord=1  and  b102.t_clot like '%{1}%'  and  b102.t_item like '%{2}%' and b405.t_suno=b020.t_suno  " +
           " "
         + " and b405.t_orno= b102.t_orno and b405.t_pono= b102.t_pono and b405.t_date= b102.t_trdt", "400", ship.lotNo.Replace(" ", ""), ship.makat.Replace(" ", ""));
           

                try
                {
                    DbReader = DbCommand.ExecuteReader();

                }
                catch (Exception ex)
                {
                    return ship;
                }

                if (DbReader.Read())
                {
                    ship.shippmentNo = DbReader.GetString(0).Replace(" ", "").Trim();
                    ship.supplierID = DbReader.GetString(1).Replace(" ", "").Trim();
                    supp.SupplierNo = ship.SupplierID;
                    supp.CompanyNo = ship.Company;
                    supp.SupplierName = DbReader.GetString(2).Trim();
                    supp.SupplierAddress = DbReader.GetString(3).Trim();
                    supp.PurchasingPerson = DbReader.GetString(4).Trim();
                    supp.SupplierPhone = DbReader.GetString(5).Trim();
                    supp.SearchKey = DbReader.GetString(6).Trim();


                }


                DbReader.Close();

            }

            return ship;


        }
        //End Ayala 13.04.2015

    }
}
