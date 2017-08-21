using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BussinessLogicLayer
{
    public class InvoiceLine
    {
        private string flexLineNo;

        public string LineNo
        {
            get { return flexLineNo; }
            set { flexLineNo = value.TrimEnd(); }
        }
        private string flexDelivery;

        public string Delivery
        {
            get { return flexDelivery; }
            set { flexDelivery = value.TrimEnd(); }
        }

        private string flexReferenceType;

        public string ReferenceType
        {
            get { return flexReferenceType; }
            set { flexReferenceType = value.TrimEnd(); }
        }
        private string flexReferenceNumber;

        public string ReferenceNumber
        {
            get { return flexReferenceNumber; }
            set { flexReferenceNumber = value.Trim(); }
        }

        private DateTime flexReferenceDate;

        public DateTime ReferenceDate
        {
            get { return flexReferenceDate; }
            set { flexReferenceDate = value; }
        }

        private string flexPaymentTerms;

        public string PaymentTerms
        {
            get { return flexPaymentTerms; }
            set { flexPaymentTerms = value.TrimEnd(); }
        }
        private string flexUnitQuantity;

        public string UnitQuantity
        {
            get { return flexUnitQuantity; }
            set { flexUnitQuantity = value; }
        }
        private decimal flexItemPriceBruto;

        public decimal ItemPriceBruto
        {
            get { return flexItemPriceBruto; }
            set { flexItemPriceBruto = value; }
        }
        private decimal flexLineSum;

        public decimal LineSum
        {
            get { return flexLineSum; }
            set { flexLineSum = value; }
        }
        private string flexItemDescription;

        public string ItemDescription
        {
            get { return flexItemDescription; }
            set { flexItemDescription = value.TrimEnd(); }
        }
        private string flexPartNumber;

        public string PartNumber
        {
            get { return flexPartNumber; }
            set { flexPartNumber = value.TrimEnd(); }
        }
        private string flexItemBarcode;

        public string ItemBarcode
        {
            get { return flexItemBarcode; }
            set { flexItemBarcode = value.Trim(); }
        }
        private string flexCustomerBarcode;
        public string CustomerBarcode
        {
            get { return flexCustomerBarcode; }
            set { flexCustomerBarcode = value.Trim(); }
        }
        private string flexSalesOrder;
        private string flexDiscount;

        public string Discount
        {
            get { return flexDiscount; }
            set { flexDiscount = value; }
        }

        public string SalesOrder
        {
            get { return flexSalesOrder; }
            set { flexSalesOrder = value.TrimEnd(); }
        }
    }
}
