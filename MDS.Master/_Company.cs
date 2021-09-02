using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MDS.Master
{
    class Company
    {
        private int _companyId;
        private string _companyCode;
        private string _engName;
        private string _engAddress1;
        private string _engAddress2;
        private string _engAddress3;
        private string _thName;
        private string _thAddress1;
        private string _thAddress2;
        private string _thAddress3;
        private string _telephone;
        private string _fax;
        private string _taxId;
        private string _branchNo;

        public Company()
        {
            CompanyId = 0;
            CompanyCode = "";
            EngName = "";
            EngAddress1 = "";
            EngAddress2 = "";
            EngAddress3 = "";
            ThName = "";
            ThAddress1 = "";
            ThAddress2 = "";
            ThAddress3 = "";
            Telephone = "";
            Fax = "";
            TaxId = "";
            BranchNo = "";
        }

        public Company(int companyId, string companyCode = "", string engName = "", string engAddress1 = "", string engAddress2 = "", string engAddress3 = "", string thName = "", string thAddress1 = "", string thAddress2 = "", string thAddress3 = "", string telephone = "", string fax = "", string taxId = "", string branchNo = "") : this()
        {
            CompanyId = companyId;
            CompanyCode = companyCode;
            EngName = engName;
            EngAddress1 = engAddress1;
            EngAddress2 = engAddress2;
            EngAddress3 = engAddress3;
            ThName = thName;
            ThAddress1 = thAddress1;
            ThAddress2 = thAddress2;
            ThAddress3 = thAddress3;
            Telephone = telephone;
            Fax = fax;
            TaxId = taxId;
            BranchNo = branchNo;
        }

        public int CompanyId { get => _companyId; set => _companyId = value; }
        public string CompanyCode { get => _companyCode; set => _companyCode = value; }
        public string EngName { get => _engName; set => _engName = value; }
        public string EngAddress1 { get => _engAddress1; set => _engAddress1 = value; }
        public string EngAddress2 { get => _engAddress2; set => _engAddress2 = value; }
        public string EngAddress3 { get => _engAddress3; set => _engAddress3 = value; }
        public string ThName { get => _thName; set => _thName = value; }
        public string ThAddress1 { get => _thAddress1; set => _thAddress1 = value; }
        public string ThAddress2 { get => _thAddress2; set => _thAddress2 = value; }
        public string ThAddress3 { get => _thAddress3; set => _thAddress3 = value; }
        public string Telephone { get => _telephone; set => _telephone = value; }
        public string Fax { get => _fax; set => _fax = value; }
        public string TaxId { get => _taxId; set => _taxId = value; }
        public string BranchNo { get => _branchNo; set => _branchNo = value; }
    }
}
