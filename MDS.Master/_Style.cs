using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MDS.Master
{
    class Style
    {
        private int _styleId;
        private string _styleName;
        private Category _category;

        public Style()
        {
            StyleId = 0;
            StyleName = "";
            Category.CategoryId = 0;
            Category.CategoryName = "";
        }

        public Style(Category category, int styleId, string styleName="") : this()
        {
            StyleId = styleId;
            StyleName = styleName;
            Category = category;
        }

        public int StyleId { get => _styleId; set => _styleId = value; }
        public string StyleName { get => _styleName; set => _styleName = value; }
        internal Category Category { get => _category; set => _category = value; }
    }
}
