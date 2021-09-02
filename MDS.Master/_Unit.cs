using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MDS.Master
{
    class Unit
    {
        private int _unitId;
        private string _unitName;

        public Unit()
        {
            UnitId = 0;
            UnitName = "";
        }

        public Unit(int unitId, string unitName) : this()
        {
            UnitId = unitId;
            UnitName = unitName;
        }

        public int UnitId { get => _unitId; set => _unitId = value; }
        public string UnitName { get => _unitName; set => _unitName = value; }
    }
}
