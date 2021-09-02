using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MDS.Master
{
    class Item
    {
        private int _itemId;
        private int _materialType;
        private string _itemCode;
        private string _description;
        private string _composition;
        private string _weightOrMoreDetail;
        private string _modelNo;
        private string _modelName;
        private Category _category;
        private Style _style;
        private Colour _color;
        private Size _size;
        private Customer _customer;
        private string _businessUnit;
        private string _season;
        private string _classType;
        private Branch _branch;
        private string _costSheetNo;
        private double _standardPrice;
        private Vendor _firstVendor;
        private int _purchaseType;
        private double _purchaseLoss;
        private int _taxBenefits;
        private DateTime _firstReceiptDate;
        private Vendor _defaultVendor;
        private double _minStock;
        private double _maxStock;
        private int _stockShelfLift;
        private double _standardCost;
        private Unit _defaultUnit;
        private Unit _convertUnit;
        private double _convertFactor;
        private string _pathFile;
        private string _labTestNo;
        private DateTime _approvedLabDate;
        private int _qcInspection;
        private Company _company;
        private Branch _branchs;
        private Department _department;
        private string _groupBOI;
        private string _groupSection;

        public Item()
        {
            ItemId = 0;
            MaterialType = 0;
            ItemCode = "";
            Description = "";
            Composition = "";
            WeightOrMoreDetail = "";
            ModelNo = "";
            ModelName = "";
            Category = new Category();
            Style = new Style();
            Color = new Colour();
            Size = new Size();
            Customer = new Customer();
            BusinessUnit = "";
            Season = "";
            ClassType = "";
            Branch = new Branch();
            CostSheetNo = "";
            StandardPrice = 0;
            FirstVendor = new Vendor();
            PurchaseType = 0;
            PurchaseLoss = 0;
            TaxBenefits = 0;
            FirstReceiptDate = default(DateTime);
            DefaultVendor = new Vendor();
            MinStock = 0;
            MaxStock = 0;
            StockShelfLift = 0;
            StandardCost = 0;
            DefaultUnit = new Unit();
            ConvertUnit = new Unit();
            ConvertFactor = 0;
            PathFile = "";
            LabTestNo = "";
            ApprovedLabDate = default(DateTime);
            QcInspection = 0;
            Company = new Company();
            Branchs = new Branch();
            Department = new Department();
            GroupBOI = "";
            GroupSection = "";
        }

        public Item(int itemId, int materialType = 0, string itemCode = "", string description = "", string composition = "", string weightOrMoreDetail = "", string modelNo = "", string modelName = "", Category category = null, Style style = null, Colour color = null, Size size = null, Customer customer = null, string businessUnit = "", string season = "", string classType = "", Branch branch = null, string costSheetNo = "", double standardPrice = 0, Vendor firstVendor = null, int purchaseType = 0, double purchaseLoss = 0, int taxBenefits = 0, DateTime firstReceiptDate = default(DateTime), Vendor defaultVendor = null, double minStock = 0, double maxStock = 0, int stockShelfLift = 0, double standardCost = 0, Unit defaultUnit = null, Unit convertUnit = null, double convertFactor = 0, string pathFile = "", string labTestNo = "", DateTime approvedLabDate = default(DateTime), int qcInspection = 0, Company company = null, Branch branchs = null, Department department = null, string groupBOI = "", string groupSection = "") : this()
        {
            ItemId = itemId;
            MaterialType = materialType;
            ItemCode = itemCode;
            Description = description;
            Composition = composition;
            WeightOrMoreDetail = weightOrMoreDetail;
            ModelNo = modelNo;
            ModelName = modelName;
            Category = category ?? new Category();
            Style = style ?? new Style();
            Color = color ?? new Colour();
            Size = size ?? new Size();
            Customer = customer ?? new Customer();
            BusinessUnit = businessUnit;
            Season = season;
            ClassType = classType;
            Branch = branch ?? new Branch();
            CostSheetNo = costSheetNo;
            StandardPrice = standardCost;
            FirstVendor = firstVendor ?? new Vendor();
            PurchaseType = purchaseType;
            PurchaseLoss = purchaseLoss;
            TaxBenefits = taxBenefits;
            FirstReceiptDate = firstReceiptDate;
            DefaultVendor = defaultVendor ?? new Vendor();
            MinStock = minStock;
            MaxStock = maxStock;
            StockShelfLift = stockShelfLift;
            StandardCost = standardCost;
            DefaultUnit = defaultUnit ?? new Unit();
            ConvertUnit = convertUnit ?? new Unit();
            ConvertFactor = convertFactor;
            PathFile = pathFile;
            LabTestNo = labTestNo;
            ApprovedLabDate = approvedLabDate;
            QcInspection = qcInspection;
            Company = company ?? new Company();
            Branchs = branch ?? new Branch();
            Department = department ?? new Department();
            GroupBOI = groupBOI;
            GroupSection = groupSection;
        }

        public int ItemId { get => _itemId; set => _itemId = value; }
        public int MaterialType { get => _materialType; set => _materialType = value; }
        public string ItemCode { get => _itemCode; set => _itemCode = value; }
        public string Description { get => _description; set => _description = value; }
        public string Composition { get => _composition; set => _composition = value; }
        public string WeightOrMoreDetail { get => _weightOrMoreDetail; set => _weightOrMoreDetail = value; }
        public string ModelNo { get => _modelNo; set => _modelNo = value; }
        public string ModelName { get => _modelName; set => _modelName = value; }
        public string BusinessUnit { get => _businessUnit; set => _businessUnit = value; }
        public string Season { get => _season; set => _season = value; }
        public string ClassType { get => _classType; set => _classType = value; }
        public string CostSheetNo { get => _costSheetNo; set => _costSheetNo = value; }
        public double StandardPrice { get => _standardPrice; set => _standardPrice = value; }
        internal Vendor FirstVendor { get => _firstVendor; set => _firstVendor = value; }
        public int PurchaseType { get => _purchaseType; set => _purchaseType = value; }
        public double PurchaseLoss { get => _purchaseLoss; set => _purchaseLoss = value; }
        public int TaxBenefits { get => _taxBenefits; set => _taxBenefits = value; }
        public DateTime FirstReceiptDate { get => _firstReceiptDate; set => _firstReceiptDate = value; }
        internal Vendor DefaultVendor { get => _defaultVendor; set => _defaultVendor = value; }
        public double MinStock { get => _minStock; set => _minStock = value; }
        public double MaxStock { get => _maxStock; set => _maxStock = value; }
        public int StockShelfLift { get => _stockShelfLift; set => _stockShelfLift = value; }
        public double StandardCost { get => _standardCost; set => _standardCost = value; }
        internal Unit DefaultUnit { get => _defaultUnit; set => _defaultUnit = value; }
        internal Unit ConvertUnit { get => _convertUnit; set => _convertUnit = value; }
        public double ConvertFactor { get => _convertFactor; set => _convertFactor = value; }
        public string PathFile { get => _pathFile; set => _pathFile = value; }
        public string LabTestNo { get => _labTestNo; set => _labTestNo = value; }
        public DateTime ApprovedLabDate { get => _approvedLabDate; set => _approvedLabDate = value; }
        public int QcInspection { get => _qcInspection; set => _qcInspection = value; }
        public string GroupBOI { get => _groupBOI; set => _groupBOI = value; }
        public string GroupSection { get => _groupSection; set => _groupSection = value; }
        internal Category Category { get => _category; set => _category = value; }
        internal Style Style { get => _style; set => _style = value; }
        internal Colour Color { get => _color; set => _color = value; }
        internal Size Size { get => _size; set => _size = value; }
        internal Customer Customer { get => _customer; set => _customer = value; }
        internal Branch Branch { get => _branch; set => _branch = value; }
        internal Company Company { get => _company; set => _company = value; }
        internal Branch Branchs { get => _branchs; set => _branchs = value; }
        internal Department Department { get => _department; set => _department = value; }
    }
}
