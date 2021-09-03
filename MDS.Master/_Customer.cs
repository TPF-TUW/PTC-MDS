using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MDS.Master
{
    class Customer
    {
        private int _customerId;
        private string _customerName;
        private string _shortName;
        private string _contacts;
        private string _email;
        private string _address1;
        private string _address2;
        private string _address3;
        private string _country;
        private string _postCode;
        private string _telephone;
        private string _fax;
        private int _customerType;
        private string _salesSection;
        private string _paymentTerm;
        private string _paymentCurrency;
        private int _calendorNo;
        private string _evalutionPoint;
        private string _otherContact;
        private string _otherAddress1;
        private string _otherAddress2;
        private string _otherAddress3;

        public Customer()
        {
            CustomerId = 0;
            CustomerName = "";
            ShortName = "";
            Contacts = "";
            Email = "";
            Address1 = "";
            Address2 = "";
            Address3 = "";
            Country = "";
            PostCode = "";
            Telephone = "";
            Fax = "";
            CustomerType = 0;
            SalesSection = "";
            PaymentTerm = "";
            PaymentCurrency = "";
            CalendorNo = 0;
            EvalutionPoint = "";
            OtherContact = "";
            OtherAddress1 = "";
            OtherAddress2 = "";
            OtherAddress3 = "";
        }

        public Customer(int customerId, string customerName, string shortName, string contacts, string email, string address1, string address2, string address3, string country, string postCode, string telephone, string fax, int customerType, string salesSection, string paymentTerm, string paymentCurrency, int calendorNo, string evalutionPoint, string otherContact, string otherAddress1, string otherAddress2, string otherAddress3) : this()
        {
            CustomerId = customerId;
            CustomerName = customerName;
            ShortName = shortName;
            Contacts = contacts;
            Email = email;
            Address1 = address1;
            Address2 = address2;
            Address3 = address3;
            Country = country;
            PostCode = postCode;
            Telephone = telephone;
            Fax = fax;
            CustomerType = customerType;
            SalesSection = salesSection;
            PaymentTerm = paymentTerm;
            PaymentCurrency = paymentCurrency;
            CalendorNo = calendorNo;
            EvalutionPoint = evalutionPoint;
            OtherContact = otherContact;
            OtherAddress1 = otherAddress1;
            OtherAddress2 = otherAddress2;
            OtherAddress3 = otherAddress3;

        }

        public int CustomerId { get => _customerId; set => _customerId = value; }
        public string CustomerName { get => _customerName; set => _customerName = value; }
        public string ShortName { get => _shortName; set => _shortName = value; }
        public string Contacts { get => _contacts; set => _contacts = value; }
        public string Email { get => _email; set => _email = value; }
        public string Address1 { get => _address1; set => _address1 = value; }
        public string Address2 { get => _address2; set => _address2 = value; }
        public string Address3 { get => _address3; set => _address3 = value; }
        public string Country { get => _country; set => _country = value; }
        public string PostCode { get => _postCode; set => _postCode = value; }
        public string Telephone { get => _telephone; set => _telephone = value; }
        public string Fax { get => _fax; set => _fax = value; }
        public int CustomerType { get => _customerType; set => _customerType = value; }
        public string SalesSection { get => _salesSection; set => _salesSection = value; }
        public string PaymentTerm { get => _paymentTerm; set => _paymentTerm = value; }
        public string PaymentCurrency { get => _paymentCurrency; set => _paymentCurrency = value; }
        public int CalendorNo { get => _calendorNo; set => _calendorNo = value; }
        public string EvalutionPoint { get => _evalutionPoint; set => _evalutionPoint = value; }
        public string OtherContact { get => _otherContact; set => _otherContact = value; }
        public string OtherAddress1 { get => _otherAddress1; set => _otherAddress1 = value; }
        public string OtherAddress2 { get => _otherAddress2; set => _otherAddress2 = value; }
        public string OtherAddress3 { get => _otherAddress3; set => _otherAddress3 = value; }
    }
}
