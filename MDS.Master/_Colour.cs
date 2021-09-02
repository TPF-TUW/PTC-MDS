using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MDS.Master
{
    class Colour
    {
        private int _colorId;
        private string _colorNo;
        private string _colorName;
        private int _colorType;

        public Colour()
        {
            ColorId = 0;
            ColorNo = "";
            ColorName = "";
            ColorType = 0;
        }

        public Colour(int colorId, string colorNo, string colorName, int colorType) : this()
        {
            ColorId = colorId;
            ColorNo = colorNo;
            ColorName = colorName;
            ColorType = colorType;
        }

        public int ColorId { get => _colorId; set => _colorId = value; }
        public string ColorNo { get => _colorNo; set => _colorNo = value; }
        public string ColorName { get => _colorName; set => _colorName = value; }
        public int ColorType { get => _colorType; set => _colorType = value; }
    }
}
