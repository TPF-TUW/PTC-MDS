using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MDS.Master
{
    class Branch
    {
        private int _branchId;
        private string _branchName;
        private Company _company;
        private int _branchType;

        public Branch()
        {
            BranchId = 0;
            BranchName = "";
            Company.CompanyId = 0;
            BranchType = 0;
        }

        public Branch(Company company, int branchId, string branchName = "", int branchType = 0) : this()
        {
            BranchId = branchId;
            BranchName = branchName;
            Company = company;
            BranchType = branchType;
        }

        public int BranchId { get => _branchId; set => _branchId = value; }
        public string BranchName { get => _branchName; set => _branchName = value; }
        internal Company Company { get => _company; set => _company = value; }
        public int BranchType { get => _branchType; set => _branchType = value; }
    }
}
