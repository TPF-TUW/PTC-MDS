using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MDS.Master
{
    class Size
    {
        private int _sizeId;
        private string _sizeNo;
        private string _sizeName;

        public Size()
        {
            SizeId = 0;
            SizeNo = "";
            SizeName = "";
        }

        public Size(int sizeId, string sizeNo, string sizeName) : this()
        {
            SizeId = sizeId;
            SizeNo = sizeNo;
            SizeName = sizeName;
        }

        public int SizeId { get => _sizeId; set => _sizeId = value; }
        public string SizeNo { get => _sizeNo; set => _sizeNo = value; }
        public string SizeName { get => _sizeName; set => _sizeName = value; }
    }
}
