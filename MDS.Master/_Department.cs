using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MDS.Master
{
    class Department
    {
        private int _departmentId;
        private string _departmentCode;
        private string _departmentName;
        private int _departmentType;
        private Company _company;
        private Branch _branch;
        private int _companyId;
        private int _branchId;

        public Department()
        {
            DepartmentId = 0;
            DepartmentCode = "";
            DepartmentName = "";
            DepartmentType = 0;
            Company.CompanyId = 0;
            Branch.BranchId = 0;
        }

        private Department(Company company, Branch branch, int departmentId, string departmentCode = "", string departmentName = "", int departmentType = 0) : this()
        {
            DepartmentId = departmentId;
            DepartmentCode = departmentCode;
            DepartmentName = departmentName;
            DepartmentType = departmentType;
            Company = company;
            Branch = branch;
        }

        public int DepartmentId { get => _departmentId; set => _departmentId = value; }
        public string DepartmentCode { get => _departmentCode; set => _departmentCode = value; }
        public string DepartmentName { get => _departmentName; set => _departmentName = value; }
        public int DepartmentType { get => _departmentType; set => _departmentType = value; }
        internal Company Company { get => _company; set => _company = value; }
        internal Branch Branch { get => _branch; set => _branch = value; }
    }
}
