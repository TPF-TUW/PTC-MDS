using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MDS.Master
{
    class Category
    {
        private int _categoryId;
        private string _categoryName;

        public Category()
        {
            CategoryId = 0;
            CategoryName = "";
        }

        public Category(int categoryId, string categoryName = "") : this()
        {
            CategoryId = categoryId;
            CategoryName = categoryName;
        }

        public int CategoryId { get => _categoryId; set => _categoryId = value; }
        public string CategoryName { get => _categoryName; set => _categoryName = value; }
    }
}
