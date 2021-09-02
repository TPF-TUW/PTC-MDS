using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MDS.Master
{
    class Vendor
    {
        private int _vendorId;
        private string _vendorCode;
        private string _vendorName;
        private string _shortName;
        private string _contacts;
        private string _email;
        private string _address1;
        private string _address2;
        private string _address3;
        private string _city;
        private string _country;
        private string _telephone;
        private string _fax;
        private int _vendorType;
        private int _paymentTerm;
        private int _paymentCurrency;
        private string _vendorEvaluation;
        private int _calendarNo;
        private int _productionLeadTime;
        private int _deliveryLeadTime;
        private int _arrivalLeadTime;
        private int _poCancelPeriod;
        private string _remark1;
        private string _remark2;
        private string _tempCurrecy;
        private string _tempPayTerm;
        private string _tempBunrs;

        public Vendor()
        {
            VendorId = 0;
            VendorCode = "";
            VendorName = "";
            ShortName = "";
            Contacts = "";
            Email = "";
            Address1 = "";
            Address2 = "";
            Address3 = "";
            City = "";
            Country = "";
            Telephone = "";
            Fax = "";
            VendorType = 0;
            PaymentTerm = 0;
            PaymentCurrency = 0;
            VendorEvaluation = "";
            CalendarNo = 0;
            ProductionLeadTime = 0;
            DeliveryLeadTime = 0;
            ArrivalLeadTime = 0;
            PoCancelPeriod = 0;
            Remark1 = "";
            Remark2 = "";
            TempCurrecy = "";
            TempPayTerm = "";
            TempBunrs = "";
        }

        public Vendor(int vendorId, string vendorName, string vendorCode = "", string shortName = "", string contacts = "", string email = "", string address1 = "", string address2 = "", string address3 = "", string city = "", string country = "", string telephone = "", string fax = "", int vendorType = 0, int paymentTerm = 0, int paymentCurrency = 0, string vendorEvaluation = "", int calendarNo = 0, int productionLeadTime = 0, int deliveryLeadTime = 0, int arrivalLeadTime = 0, int poCancelPeriod = 0, string remark1 = "", string remark2 = "", string tempCurrecy = "", string tempPayTerm = "", string tempBunrs = "") : this()
        {
            VendorId = vendorId;
            VendorCode = vendorCode;
            VendorName = vendorName;
            ShortName = shortName;
            Contacts = contacts;
            Email = email;
            Address1 = address1;
            Address2 = address2;
            Address3 = address3;
            City = city;
            Country = country;
            Telephone = telephone;
            Fax = fax;
            VendorType = vendorType;
            PaymentTerm = paymentTerm;
            PaymentCurrency = paymentCurrency;
            VendorEvaluation = vendorEvaluation;
            CalendarNo = calendarNo;
            ProductionLeadTime = productionLeadTime;
            DeliveryLeadTime = deliveryLeadTime;
            ArrivalLeadTime = arrivalLeadTime;
            PoCancelPeriod = poCancelPeriod;
            Remark1 = remark1;
            Remark2 = remark2;
            TempCurrecy = tempCurrecy;
            TempPayTerm = tempPayTerm;
            TempBunrs = tempBunrs;
        }

        public int VendorId { get => _vendorId; set => _vendorId = value; }
        public string VendorCode { get => _vendorCode; set => _vendorCode = value; }
        public string VendorName { get => _vendorName; set => _vendorName = value; }
        public string ShortName { get => _shortName; set => _shortName = value; }
        public string Contacts { get => _contacts; set => _contacts = value; }
        public string Email { get => _email; set => _email = value; }
        public string Address1 { get => _address1; set => _address1 = value; }
        public string Address2 { get => _address2; set => _address2 = value; }
        public string Address3 { get => _address3; set => _address3 = value; }
        public string City { get => _city; set => _city = value; }
        public string Country { get => _country; set => _country = value; }
        public string Telephone { get => _telephone; set => _telephone = value; }
        public string Fax { get => _fax; set => _fax = value; }
        public int VendorType { get => _vendorType; set => _vendorType = value; }
        public int PaymentTerm { get => _paymentTerm; set => _paymentTerm = value; }
        public int PaymentCurrency { get => _paymentCurrency; set => _paymentCurrency = value; }
        public string VendorEvaluation { get => _vendorEvaluation; set => _vendorEvaluation = value; }
        public int CalendarNo { get => _calendarNo; set => _calendarNo = value; }
        public int ProductionLeadTime { get => _productionLeadTime; set => _productionLeadTime = value; }
        public int DeliveryLeadTime { get => _deliveryLeadTime; set => _deliveryLeadTime = value; }
        public int ArrivalLeadTime { get => _arrivalLeadTime; set => _arrivalLeadTime = value; }
        public int PoCancelPeriod { get => _poCancelPeriod; set => _poCancelPeriod = value; }
        public string Remark1 { get => _remark1; set => _remark1 = value; }
        public string Remark2 { get => _remark2; set => _remark2 = value; }
        public string TempCurrecy { get => _tempCurrecy; set => _tempCurrecy = value; }
        public string TempPayTerm { get => _tempPayTerm; set => _tempPayTerm = value; }
        public string TempBunrs { get => _tempBunrs; set => _tempBunrs = value; }
    }
}
