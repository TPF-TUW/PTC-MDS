using System;
using System.Text;
using System.Windows.Forms;
using DevExpress.LookAndFeel;
using DevExpress.Utils.Extensions;
using DBConnect;
using System.Data;
using DevExpress.XtraEditors.Controls;
using System.CodeDom;
using DevExpress.XtraGrid.Views.Grid;
using System.Drawing;
using DevExpress.XtraPrinting;
using DevExpress.XtraEditors;
using TheepClass;
using DevExpress.Utils;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using System.Text.RegularExpressions;
using DevExpress.Spreadsheet;
using System.IO;

namespace MDS.Master
{
    public partial class M07 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private Functionality.Function FUNC = new Functionality.Function();
        private const string imgPathFile = Configuration.CONFIG.PATH_FILE + @"Pictures\";
        private DataTable dtVendor = new DataTable();
        private string selCode = "";
        StringBuilder sbMeterial = new StringBuilder();
        public LogIn UserLogin { get; set; }
        public int Company { get; set; }

        public string ConnectionString { get; set; }
        string CONNECT_STRING = "";
        DatabaseConnect DBC;

        public M07()
        {
            InitializeComponent();
            UserLookAndFeel.Default.StyleChanged += MyStyleChanged;
        }

        private void MyStyleChanged(object sender, EventArgs e)
        {
            UserLookAndFeel userLookAndFeel = (UserLookAndFeel)sender;
            cUtility.SaveRegistry(@"Software\MDS", "SkinName", userLookAndFeel.SkinName);
            cUtility.SaveRegistry(@"Software\MDS", "SkinPalette", userLookAndFeel.ActiveSvgPaletteName);
        }

        private void XtraForm1_Load(object sender, EventArgs e)
        {
           // MessageBox.Show(this.UserLogin.OIDUser.ToString() + ", Company-" + this.UserLogin.OIDCompany.ToString() + ", Dept-" + this.UserLogin.OIDDept.ToString() + ", Branch-" + this.UserLogin.OIDBranch.ToString());
           // MessageBox.Show(Configuration.CONFIG.DATABASE_FILE);
            //***** SET CONNECT DB ********
            if (this.ConnectionString != null)
            {
                if (this.ConnectionString != "")
                {
                    CONNECT_STRING = this.ConnectionString;
                }
            }

            this.DBC = new DatabaseConnect(CONNECT_STRING);

            if (this.DBC.chkCONNECTION_STING() == false)
            {
                this.DBC.setCONNECTION_STRING_INIFILE();
                if (this.DBC.chkCONNECTION_STING() == false)
                {
                    return;
                }
            }
            new ObjDE.setDatabase(this.DBC);
            //*****************************

            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT TOP (1) ReadWriteStatus FROM FunctionAccess WHERE (OIDUser = '" + UserLogin.OIDUser + "') AND(FunctionNo = 'M07') ");
            int chkReadWrite = this.DBC.DBQuery(sbSQL).getInt();
            if (chkReadWrite == 0)
            {
                bbiNew.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
                bbiEdit.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
                bbiSave.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
                bbiDelete.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
                bbiRefresh.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
            }

            sbSQL.Clear();
            sbSQL.Append("SELECT FullName, OIDUSER FROM Users ORDER BY OIDUSER ");
            new ObjDE.setGridLookUpEdit(glueCREATE, sbSQL, "FullName", "OIDUSER").getData();
            new ObjDE.setGridLookUpEdit(glueUPDATE, sbSQL, "FullName", "OIDUSER").getData();

            glueCREATE.EditValue = UserLogin.OIDUser;
            glueUPDATE.EditValue = UserLogin.OIDUser;

            tabbedControlGroup1.SelectedTabPage = layoutControlGroup13; //เลือกแท็บ Search
            rgMaterial.EditValue = 0;

            //glueCode.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard;
            //glueCode.Properties.AcceptEditorTextAsNewValue = DevExpress.Utils.DefaultBoolean.True;

            sbMeterial.Append(" SELECT '0' AS ID, 'Finished Goods' AS MaterialType ");
            sbMeterial.Append("UNION ALL ");
            sbMeterial.Append(" SELECT '1' AS ID, 'Fabric' AS MaterialType ");
            sbMeterial.Append("UNION ALL ");
            sbMeterial.Append(" SELECT '2' AS ID, 'Accessory' AS MaterialType ");
            sbMeterial.Append("UNION ALL ");
            sbMeterial.Append(" SELECT '3' AS ID, 'Packaging' AS MaterialType ");
            sbMeterial.Append("UNION ALL ");
            sbMeterial.Append(" SELECT '4' AS ID, 'Sample' AS MaterialType ");
            sbMeterial.Append("UNION ALL ");
            sbMeterial.Append(" SELECT '8' AS ID, 'Temporary' AS MaterialType ");
            sbMeterial.Append("UNION ALL ");
            sbMeterial.Append(" SELECT '9' AS ID, 'Finished Goods' AS MaterialType ");

            NewData();
            LoadData();
            if (rgMaterial.SelectedIndex > -1)
            {
                LoadDataMeterial();
            }
        }

        private void NewData(bool clearCode = true)
        {
            //rgMaterial.SelectedIndex = 0;
            txeID.Text = "";

            lblStatus.Text = "NEW";
            lblStatus.ForeColor = Color.White;
            lblStatus.BackColor = Color.Green;
            layoutControlItem78.AppearanceItemCaption.BackColor = Color.Green;

            //if (rgMaterial.SelectedIndex > -1)
            //{
            //    string Material = rgMaterial.EditValue.ToString();
            //    txeID.Text = this.DBC.DBQuery("SELECT CASE WHEN ISNULL(MAX(OIDITEM), '') = '' THEN 1 ELSE MAX(OIDITEM) + 1 END AS NewNo FROM Items WHERE (MaterialType = '" + Material + "') ").getString();
            //}

            layoutControlGroup4.Text = ".";

            txeOldCode.Text = "";
            txeID.Text = "";
            if(clearCode == true)
                glueCode.EditValue = "";
            txeDescription.Text = "";
            txeComposition.Text = "";
            txeWeight.Text = "";
            txeModelNo.Text = "";
            txeModelName.Text = "";
            slueCategory.EditValue = "";
            slueStyle.EditValue = "";
            slueColor.EditValue = "";
            slueSize.EditValue = "";
            slueCustomer.EditValue = "";
            txeBusinessUnit.Text = "";
            cbeSeason.EditValue = "";
            cbeClass.EditValue = "";
            glueBranch.EditValue = "";
            txeCostSheet.Text = "";
            txeStdPrice.Text = "";

            slueFirstVendor.EditValue = "";
            txeMatDetails.Text = "";
            txeMatCode.Text = "";
            txeSMPLLotNo.Text = "";
            txePrice.Text = "";
            slueCurrency.EditValue = "";
            rgPurchase.EditValue = 0;
            rgTax.EditValue = 1;
            txePurchaseLoss.Text = "";
            dteFirstReceiptDate.EditValue = DateTime.Now;
            slueDefaultVendor.EditValue = "";

            txeSMPLNo.Text = "";
            dteRequestDate.EditValue = DateTime.Now;
            txeSMPLItem.Text = "";
            txeSMPLPatternNo.Text = "";
            rgZone.EditValue = 0;

            txeMinStock.Text = "";
            txeMaxStock.Text = "";
            txeStockSheifLife.Text = "";
            txeStdCost.Text = "";
            slueDefaultUnit.EditValue = "";
            slueUnit.EditValue = "";
            txeFactor.Text = "";

            picImg.Image = null;
            txePath.Text = "";

            txeLabTestNo.Text = "";
            dteApprovedLabDate.EditValue = DateTime.Now;
            txeQCInspection.Text = "";
            clbQC.Items.Clear();

            slueVendorCode.EditValue = "";
            txeVendorName.Text = "";
            txeLotSize.Text = "";
            txeProductionLead.Text = "";
            txeDeliveryLead.Text = "";
            txeArrivalLead.Text = "";
            txePOCancelPeriod.Text = "";

            txeLots1.Text = "";
            txeLots2.Text = "";
            txeLots3.Text = "";

            txeRemark.Text = "";
            lblIDVENDItem.Text = "";

            glueCREATE.EditValue = UserLogin.OIDUser;
            txeCDATE.Text = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
            glueUPDATE.EditValue = UserLogin.OIDUser;
            txeUDATE.Text = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");

            //**************************************
            slueFirstVendor.ReadOnly = false;
            txeMatDetails.ReadOnly = false;
            txeMatCode.ReadOnly = false;
            txeSMPLLotNo.ReadOnly = false;
            txePrice.ReadOnly = false;
            slueCurrency.ReadOnly = false;
            rgPurchase.ReadOnly = false;
            rgTax.ReadOnly = false;
            txePurchaseLoss.ReadOnly = false;
            dteFirstReceiptDate.ReadOnly = false;
            //**************************************
            selCode = "";

            string Materialx = rgMaterial.EditValue.ToString();
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT OIDCONDQC, ConditionName ");
            sbSQL.Append("FROM ConditionQC ");
            sbSQL.Append("WHERE (ItemType = '" + Materialx + "') ");
            sbSQL.Append("ORDER BY ConditionName ");
            DataTable drMaterial = this.DBC.DBQuery(sbSQL).getDataTable();
            clbQC.ValueMember = "OIDCONDQC";
            clbQC.DisplayMember = "ConditionName";
            clbQC.DataSource = drMaterial;

            dtVendor.Rows.Clear();
            tabbedControlGroup1.SelectedTabPage = layoutControlGroup1; //เลือกแท็บ Main
        }

        private void LoadDataMeterial()
        {
            string Material = rgMaterial.EditValue.ToString();

            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT OIDITEM AS ID, Code, Description ");
            sbSQL.Append("FROM  Items ");
            sbSQL.Append("WHERE (MaterialType = '" + Material + "')");
            sbSQL.Append("ORDER BY Code ");
            new ObjDE.setGridLookUpEdit(glueCode, sbSQL, "Code", "ID").getData(true);
            glueCode.Properties.View.PopulateColumns(glueCode.Properties.DataSource);
            glueCode.Properties.View.Columns["ID"].Visible = false;

            sbSQL.Clear();
            sbSQL.Append("SELECT OIDCONDQC, ConditionName ");
            sbSQL.Append("FROM ConditionQC ");
            sbSQL.Append("WHERE (ItemType = '" + Material + "') ");
            sbSQL.Append("ORDER BY ConditionName ");
            DataTable drMaterial = this.DBC.DBQuery(sbSQL).getDataTable();
            clbQC.ValueMember = "OIDCONDQC";
            clbQC.DisplayMember = "ConditionName";
            clbQC.DataSource = drMaterial;

            LoadVendorData();
        }

        private void LoadData()
        {

            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT CategoryName, OIDGCATEGORY AS ID ");
            sbSQL.Append("FROM  GarmentCategory ");
            sbSQL.Append("ORDER BY CategoryName ");
            new ObjDE.setSearchLookUpEdit(slueCategory, sbSQL, "CategoryName", "ID").getData(true);

            sbSQL.Clear();
            sbSQL.Append("SELECT ColorNo, ColorName, OIDCOLOR AS ID ");
            sbSQL.Append("FROM  ProductColor ");
            sbSQL.Append("ORDER BY ColorNo ");
            new ObjDE.setSearchLookUpEdit(slueColor, sbSQL, "ColorName", "ID").getData(true);

            sbSQL.Clear();
            sbSQL.Append("SELECT SizeNo, SizeName, OIDSIZE AS ID ");
            sbSQL.Append("FROM  ProductSize ");
            sbSQL.Append("ORDER BY SizeNo ");
            new ObjDE.setSearchLookUpEdit(slueSize, sbSQL, "SizeName", "ID").getData(true);

            sbSQL.Clear();
            sbSQL.Append("SELECT Code, Name, ShortName, OIDCUST AS ID ");
            sbSQL.Append("FROM  Customer ");
            sbSQL.Append("ORDER BY Name ");
            new ObjDE.setSearchLookUpEdit(slueCustomer, sbSQL, "ShortName", "ID").getData(true);

            sbSQL.Clear();
            //sbSQL.Append("SELECT DISTINCT ISNULL(Season, '') AS Season ");
            //sbSQL.Append("FROM  Items ");
            //sbSQL.Append("ORDER BY Season ");
            //sbSQL.Append("SELECT SeasonNo AS Season, SeasonName ");
            //sbSQL.Append("FROM  Season ");
            //sbSQL.Append("ORDER BY OIDSEASON ");
            sbSQL.Append("SELECT Season ");
            sbSQL.Append("FROM (SELECT FORMAT(DATEADD(year, -1, GETDATE()), 'yy') + SeasonNo AS Season, FORMAT(DATEADD(year, -1, GETDATE()), 'yy') AS Year, SeasonNo, OIDSEASON ");
            sbSQL.Append("      FROM   Season AS SS1 ");
            sbSQL.Append("      UNION ALL ");
            sbSQL.Append("      SELECT FORMAT(GETDATE(), 'yy') + SeasonNo AS Season, FORMAT(GETDATE(), 'yy') AS Year, SeasonNo, OIDSEASON ");
            sbSQL.Append("      FROM   Season AS SS2 ");
            sbSQL.Append("      UNION ALL ");
            sbSQL.Append("      SELECT FORMAT(DATEADD(year, 1, GETDATE()), 'yy') + SeasonNo AS Season, FORMAT(DATEADD(year, 1, GETDATE()), 'yy') AS Year, SeasonNo, OIDSEASON ");
            sbSQL.Append("      FROM   Season AS SS3) AS SS ");
            sbSQL.Append("ORDER BY Year DESC, OIDSEASON ");
            new ObjDE.setSearchLookUpEdit(cbeSeason, sbSQL, "Season", "Season").getData(true);

            sbSQL.Clear();
            sbSQL.Append("SELECT DISTINCT ISNULL(ClassType, '') AS ClassType ");
            sbSQL.Append("FROM  Items ");
            sbSQL.Append("ORDER BY ClassType ");
            new ObjDE.setComboboxEdit(cbeClass, sbSQL).getDataRange();

            sbSQL.Clear();
            sbSQL.Append("SELECT Name AS Branch, OIDBranch AS ID ");
            sbSQL.Append("FROM  Branchs ");
            sbSQL.Append("WHERE (OIDCOMPANY = '" + this.Company + "') ");
            sbSQL.Append("ORDER BY OIDBranch ");
            new ObjDE.setGridLookUpEdit(glueBranch, sbSQL, "Branch", "ID").getData(true);
            new ObjDE.setSearchLookUpEdit(slueSBranch, sbSQL, "Branch", "ID").getData(true);

            sbSQL.Clear();
            sbSQL.Append("SELECT UnitName, OIDUNIT AS ID ");
            sbSQL.Append("FROM  Unit ");
            sbSQL.Append("ORDER BY UnitName ");
            new ObjDE.setSearchLookUpEdit(slueDefaultUnit, sbSQL, "UnitName", "ID").getData(true);

            new ObjDE.setSearchLookUpEdit(slueUnit, sbSQL, "UnitName", "ID").getData(true);

            sbSQL.Clear();
            sbSQL.Append("SELECT Currency, OIDCURR AS ID ");
            sbSQL.Append("FROM  Currency ");
            sbSQL.Append("ORDER BY OIDCURR ");
            new ObjDE.setSearchLookUpEdit(slueCurrency, sbSQL, "Currency", "ID").getData(true);

            //slueUnit.Properties.DataSource = slueDefaultUnit.Properties.DataSource;
            //slueUnit.Properties.DisplayMember = slueDefaultUnit.Properties.DisplayMember;
            //slueUnit.Properties.ValueMember = slueDefaultUnit.Properties.ValueMember;

            sbSQL.Clear();
            sbSQL.Append("SELECT VD.Code, VD.Name, ENT.Name AS Type, VD.OIDVEND AS ID ");
            sbSQL.Append("FROM   Vendor AS VD INNER JOIN ");
            sbSQL.Append("       ENUMTYPE AS ENT ON VD.VendorType = ENT.No AND ENT.Module = N'Vendor' ");
            sbSQL.Append("ORDER BY VD.Name ");
            //new ObjDE.setSearchLookUpEdit(slueFirstVendor, sbSQL, "Name", "ID").getData(true);
            new ObjDE.setSearchLookUpEdit(slueDefaultVendor, sbSQL, "Name", "ID").getData(true);
            //slueDefaultVendor.Properties.View.Columns["ID"].Visible = false;
            //new ObjDE.setSearchLookUpEdit(slueVendorCode, sbSQL, "Code", "ID").getData(true);

            sbSQL.Clear();
            sbSQL.Append("SELECT VD.Code, VD.Name, ENT.Name AS Type, VD.OIDVEND AS ID ");
            sbSQL.Append("FROM   Vendor AS VD INNER JOIN ");
            sbSQL.Append("       ENUMTYPE AS ENT ON VD.VendorType = ENT.No AND ENT.Module = N'Vendor' ");
            sbSQL.Append("ORDER BY VD.Name ");
            new ObjDE.setSearchLookUpEdit(slueVendorCode, sbSQL, "Code", "ID").getData(true);

            //SEARCH
            new ObjDE.setSearchLookUpEdit(slueSMaterial, sbMeterial, "MaterialType", "ID").getData(true);

            slueSCustomer.Properties.DataSource = slueCustomer.Properties.DataSource;
            slueSCustomer.Properties.DisplayMember = slueCustomer.Properties.DisplayMember;
            slueSCustomer.Properties.ValueMember = slueCustomer.Properties.ValueMember;

            slueFirstVendor.Properties.DataSource = slueDefaultVendor.Properties.DataSource;
            slueFirstVendor.Properties.DisplayMember = slueDefaultVendor.Properties.DisplayMember;
            slueFirstVendor.Properties.ValueMember = slueDefaultVendor.Properties.ValueMember;
            //slueFirstVendor.Properties.View.Columns["ID"].Visible = false;

            slueSVendor.Properties.DataSource = slueDefaultVendor.Properties.DataSource;
            slueSVendor.Properties.DisplayMember = slueDefaultVendor.Properties.DisplayMember;
            slueSVendor.Properties.ValueMember = slueDefaultVendor.Properties.ValueMember;
            //slueSVendor.Properties.View.Columns["ID"].Visible = false;

            slueSCategory.Properties.DataSource = slueCategory.Properties.DataSource;
            slueSCategory.Properties.DisplayMember = slueCategory.Properties.DisplayMember;
            slueSCategory.Properties.ValueMember = slueCategory.Properties.ValueMember;

            slueSStyle.Properties.DataSource = slueStyle.Properties.DataSource;
            slueSStyle.Properties.DisplayMember = slueStyle.Properties.DisplayMember;
            slueSStyle.Properties.ValueMember = slueStyle.Properties.ValueMember;


            sbSQL.Clear();
            sbSQL.Append("SELECT ModelNo FROM Items ORDER BY ModelNo");
            new ObjDE.setSearchLookUpEdit(slueSModel, sbSQL, "ModelNo", "ModelNo").getData(true);

            LoadStyle();

            LoadItemCode();

            LOAD_LIST();
        }

        private void LoadItemCode()
        {
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT Code, Description ");
            sbSQL.Append("FROM  Items ");
            if (slueSMaterial.Text.Trim() != "")
                sbSQL.Append("WHERE (MaterialType = '" + slueSMaterial.EditValue.ToString() + "') ");
            sbSQL.Append("ORDER BY Code ");
            new ObjDE.setSearchLookUpEdit(slueSCode, sbSQL, "Code", "Code").getData(true);
        }

        private void LoadStyle()
        {
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT StyleName, OIDSTYLE AS ID ");
            sbSQL.Append("FROM  ProductStyle ");
            if (slueSCategory.Text.Trim() != "")
                sbSQL.Append("WHERE (OIDGCATEGORY = '" + slueSCategory.EditValue.ToString() + "') ");
            sbSQL.Append("ORDER BY StyleName ");
            new ObjDE.setSearchLookUpEdit(slueSStyle, sbSQL, "StyleName", "ID").getData(true);
        }

        private void LOAD_LIST()
        {
            gcListItem.DataSource = null;
            if (slueSMaterial.Text.Trim() != "" || 
                slueSBranch.Text.Trim() != "" || 
                slueSCustomer.Text.Trim() != "" || 
                slueSVendor.Text.Trim() != "" || 
                slueSCategory.Text.Trim() != "" ||
                slueSStyle.Text.Trim() != "" ||
                slueSCode.Text.Trim() != "" ||
                slueSModel.Text.Trim() != "")
            {
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT IT.OIDITEM, IT.MaterialType AS MaterialTypeID, MTYPE.MaterialType, IT.Code, IT.Description, IT.Composition, IT.WeightOrMoreDetail, IT.ModelNo, IT.ModelName, IT.OIDCATEGORY, GC.CategoryName, IT.OIDSTYLE, ");
                sbSQL.Append("       PS.StyleName, IT.OIDCOLOR, PC.ColorNo, PC.ColorName, IT.OIDSIZE, PSZ.SizeNo, PSZ.SizeName, IT.OIDCUST, CUS.ShortName, CUS.Name, IT.BusinessUnit, IT.Season, IT.ClassType, IT.Branch AS BranchID, ");
                sbSQL.Append("       BN.Name AS Branch, IT.CostSheetNo, IT.StdPrice, IT.FirstVendor AS FirstVendID, VD2.Name AS FirstVendor, IT.PurchaseType AS PurchaseTypeID, ");
                sbSQL.Append("       CASE WHEN IT.PurchaseType = 0 THEN 'Local' ELSE CASE WHEN IT.PurchaseType = 1 THEN 'Import' ELSE CASE WHEN IT.PurchaseType = 9 THEN 'Other' ELSE '' END END END AS PurchaseType, IT.PurchaseLoss, ");
                sbSQL.Append("       IT.TaxBenefits AS TaxBenefitsID, CASE WHEN IT.TaxBenefits = 1 THEN 'BOI' ELSE CASE WHEN IT.TaxBenefits = 2 THEN '19 BIS' ELSE CASE WHEN IT.TaxBenefits = 9 THEN 'Other' ELSE '' END END END AS TaxBenefits, ");
                sbSQL.Append("       CASE WHEN IT.FirstReceiptDate IS NULL THEN '' ELSE CONVERT(VARCHAR(10), IT.FirstReceiptDate, 103) END FirstReceiptDate, IT.DefaultVendor AS DefaultVendID, VD.Name AS DefaultVendor, IT.MinStock, IT.MaxStock, IT.StockShelfLife, IT.StdCost, IT.DefaultUnit AS DefaultUnitID, UN.UnitName AS DefaultUnit, UN2.UnitName AS ConvertUnit, CASE WHEN IT.ConvertFactor = 0 THEN NULL ELSE IT.ConvertFactor END AS ConversionFactor, IT.PathFile, IT.LabTestNo, ");
                sbSQL.Append("       CASE WHEN IT.ApprovedLabDate IS NULL THEN '' ELSE CONVERT(VARCHAR(10), IT.ApprovedLabDate, 103) END AS ApprovedLabDate, IT.QCInspection, IT.CreatedBy, IT.CreatedDate, IT.UpdatedBy, IT.UpdatedDate ");
                sbSQL.Append("FROM   Items AS IT LEFT OUTER JOIN ");
                sbSQL.Append("       ProductStyle AS PS ON IT.OIDSTYLE = PS.OIDSTYLE LEFT OUTER JOIN ");
                sbSQL.Append("       ProductColor AS PC ON IT.OIDCOLOR = PC.OIDCOLOR LEFT OUTER JOIN ");
                sbSQL.Append("       ProductSize AS PSZ ON IT.OIDSIZE = PSZ.OIDSIZE LEFT OUTER JOIN ");
                sbSQL.Append("       Customer AS CUS ON IT.OIDCUST = CUS.OIDCUST LEFT OUTER JOIN ");
                sbSQL.Append("       GarmentCategory AS GC ON IT.OIDCATEGORY = GC.OIDGCATEGORY LEFT OUTER JOIN ");
                sbSQL.Append("       Vendor AS VD ON IT.DefaultVendor = VD.OIDVEND LEFT OUTER JOIN ");
                sbSQL.Append("       Vendor AS VD2 ON IT.FirstVendor = VD2.OIDVEND LEFT OUTER JOIN ");
                sbSQL.Append("       Unit AS UN ON IT.DefaultUnit = UN.OIDUNIT LEFT OUTER JOIN ");
                sbSQL.Append("       Unit AS UN2 ON IT.ConvertUnit = UN2.OIDUNIT LEFT OUTER JOIN ");
                sbSQL.Append("       Branchs AS BN ON IT.Branch = BN.OIDBranch LEFT OUTER JOIN ");
                sbSQL.Append("       (" + sbMeterial.ToString() + ") AS MTYPE ON IT.MaterialType = MTYPE.ID ");
                sbSQL.Append("WHERE (IT.OIDCOMPANY = '" + this.Company + "') ");
                if(slueSMaterial.Text.Trim() != "")
                    sbSQL.Append("AND (IT.MaterialType = '" + slueSMaterial.EditValue.ToString() + "') ");
                if(slueSBranch.Text.Trim() != "")
                    sbSQL.Append("AND (IT.Branch = '" + slueSBranch.EditValue.ToString() + "') ");
                if(slueSCustomer.Text.Trim() != "")
                    sbSQL.Append("AND (IT.OIDCUST = '" + slueSCustomer.EditValue.ToString() + "') ");
                if(slueSVendor.Text.Trim() != "")
                    sbSQL.Append("AND (IT.DefaultVendor = '" + slueSVendor.EditValue.ToString() + "') ");
                if(slueSCategory.Text.Trim() != "")
                    sbSQL.Append("AND (IT.OIDCATEGORY = '" + slueSCategory.EditValue.ToString() + "') ");
                if(slueSStyle.Text.Trim() != "")
                    sbSQL.Append("AND (IT.OIDSTYLE = '" + slueSStyle.EditValue.ToString() + "') ");
                if(slueSCode.Text.Trim() != "")
                    sbSQL.Append("AND (IT.Code = N'" + slueSCode.EditValue.ToString() + "') ");
                if(slueSModel.Text.Trim() != "")
                    sbSQL.Append("AND (IT.ModelNo = N'" + slueSModel.EditValue.ToString() + "') ");
                sbSQL.Append("ORDER BY IT.OIDITEM, IT.Code ");
                //MessageBox.Show(sbSQL.ToString());
                new ObjDE.setGridControl(gcListItem, gvListItem, sbSQL).getData(false, false, false, true);

                gvListItem.Columns["MaterialTypeID"].Visible = false;
                gvListItem.Columns["OIDITEM"].Visible = false;
                gvListItem.Columns["OIDCATEGORY"].Visible = false;
                gvListItem.Columns["OIDSTYLE"].Visible = false;
                gvListItem.Columns["OIDCOLOR"].Visible = false;
                gvListItem.Columns["OIDSIZE"].Visible = false;
                gvListItem.Columns["OIDCUST"].Visible = false;
                gvListItem.Columns["BranchID"].Visible = false;
                gvListItem.Columns["FirstVendID"].Visible = false;
                gvListItem.Columns["PurchaseTypeID"].Visible = false;
                gvListItem.Columns["TaxBenefitsID"].Visible = false;
                gvListItem.Columns["DefaultVendID"].Visible = false;
                gvListItem.Columns["DefaultUnitID"].Visible = false;
                gvListItem.Columns["PathFile"].Visible = false;

                gvListItem.Columns["CreatedBy"].Visible = false;
                gvListItem.Columns["CreatedDate"].Visible = false;
                gvListItem.Columns["UpdatedBy"].Visible = false;
                gvListItem.Columns["UpdatedDate"].Visible = false;
            }
        }

        private void bbiNew_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            NewData();
            LoadData();
            if (rgMaterial.SelectedIndex > -1)
            {
                LoadDataMeterial();
            }
        }

        public static bool IsNumeric(string Expression)
        {
            double retNum;

            bool isNum = Double.TryParse(Expression.ToString(), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
            return isNum;
        }

        private void bbiSave_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (tabbedControlGroup1.SelectedTabPage == layoutControlGroup1 || tabbedControlGroup1.SelectedTabPage == layoutControlGroup2) //Tab Main & Vendor Detail
            {
                bool chkPass = true;
                if (glueCode.Text.Trim() != "")
                {
                    if (glueCode.Text.Length >= 5)
                        if (glueCode.Text.Substring(0, 5) == "TMPFB" || glueCode.Text.Substring(0, 5) == "TMPMT")
                        {
                            FUNC.msgWarning("Cannot set code starting with 'TMPFB' or 'TMPMT'. Please change code.");
                            glueCode.Focus();
                            chkPass = false;
                        }
                }

                if (chkPass == true)
                {
                    if (glueCode.Text.Trim() == "")
                    {
                        FUNC.msgWarning("Please input items code.");
                        glueCode.Focus();
                    }
                    else if (txeDescription.Text.Trim() == "")
                    {
                        FUNC.msgWarning("Please input description.");
                        txeDescription.Focus();
                    }
                    else if (glueBranch.Text.Trim() == "")
                    {
                        FUNC.msgWarning("Please input branch.");
                        glueBranch.Focus();
                    }
                    //else if (txeStdPrice.Text.Trim() == "")
                    //{
                    //    FUNC.msgWarning("Please input standard price.");
                    //    txeStdPrice.Focus();
                    //}
                    else if (slueFirstVendor.Text.Trim() == "")
                    {
                        FUNC.msgWarning("Please select supplier (start raw mat.).");
                        slueFirstVendor.Focus();
                    }
                    else if (slueDefaultVendor.Text.Trim() == "")
                    {
                        FUNC.msgWarning("Please select default supplier.");
                        slueDefaultVendor.Focus();
                    }
                    //else if (txeQCInspection.Text.Trim() == "")
                    //{
                    //    FUNC.msgWarning("Please input quality inspection %.");
                    //    txeQCInspection.Focus();
                    //}
                    else
                    {
                        if (FUNC.msgQuiz("Confirm save ?") == true)
                        {
                            bool chkPASS = true;
                            //Check Vendor
                            if (slueFirstVendor.Text.Trim() != "")
                            {
                                bool chkVendor = false;
                                foreach (DataRow dr in dtVendor.Rows) // search whole table
                                {
                                    if (dr["OIDVEND"].ToString() == slueFirstVendor.EditValue.ToString())
                                    {
                                        chkVendor = true;
                                        break;
                                    }
                                }

                                if (chkVendor == false)
                                {
                                    FUNC.msgWarning("Please add vendor:'" + slueFirstVendor.Text.Trim() + "' to vendor details.");
                                    tabbedControlGroup1.SelectedTabPage = layoutControlGroup2;
                                    slueVendorCode.EditValue = slueFirstVendor.EditValue;
                                    slueVendorCode.Focus();
                                    chkPASS = false;
                                }
                                else
                                {
                                    if (slueFirstVendor.EditValue.ToString() != slueDefaultVendor.EditValue.ToString())
                                    {
                                        chkVendor = false;
                                        foreach (DataRow dr in dtVendor.Rows) // search whole table
                                        {
                                            if (dr["OIDVEND"].ToString() == slueDefaultVendor.EditValue.ToString())
                                            {
                                                chkVendor = true;
                                                break;
                                            }
                                        }

                                        if (chkVendor == false)
                                        {
                                            FUNC.msgWarning("Please add vendor:'" + slueDefaultVendor.Text.Trim() + "' to vendor details.");
                                            tabbedControlGroup1.SelectedTabPage = layoutControlGroup2;
                                            slueVendorCode.EditValue = slueDefaultVendor.EditValue;
                                            slueVendorCode.Focus();
                                            chkPASS = true;
                                        }
                                    }
                                }
                            }

                            if (chkPASS == true)
                            {
                                StringBuilder sbSQL = new StringBuilder();

                                string MaterialType = rgMaterial.EditValue.ToString();
                                string PurchaseType = rgPurchase.EditValue.ToString();
                                string TaxBenefits = "0";
                                if (rgTax.SelectedIndex != -1)
                                {
                                    TaxBenefits = rgTax.EditValue.ToString();
                                }
                                string Zone = rgZone.EditValue.ToString();

                                string newFileName = "";
                                //CopyFile
                                if (txePath.Text.Trim() != "")
                                {
                                    System.IO.FileInfo fi = new System.IO.FileInfo(txePath.Text);
                                    string extn = fi.Extension;
                                    newFileName = glueCode.Text.ToUpper().Trim() + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + extn;
                                    string newPathFileName = imgPathFile + newFileName;
                                    //MessageBox.Show(newFileName);
                                    System.IO.File.Copy(txePath.Text, newPathFileName);
                                }

                                string strCREATE = UserLogin.OIDUser.ToString() != "" ? UserLogin.OIDUser.ToString() : "0";
                                string strUPDATE = UserLogin.OIDUser.ToString() != "" ? UserLogin.OIDUser.ToString() : "0";

                                //******** save Items table ************
                                string OIDCATEGORY = slueCategory.Text == "" ? "NULL" : "'" + slueCategory.EditValue.ToString() + "'";
                                string OIDSTYLE = slueStyle.Text == "" ? "NULL" : "'" + slueStyle.EditValue.ToString() + "'";
                                string OIDCOLOR = slueColor.Text == "" ? "NULL" : "'" + slueColor.EditValue.ToString() + "'";
                                string OIDSIZE = slueSize.Text == "" ? "NULL" : "'" + slueSize.EditValue.ToString() + "'";
                                string OIDCUST = slueCustomer.Text == "" ? "NULL" : "'" + slueCustomer.EditValue.ToString() + "'";
                                string PurchaseLoss = txePurchaseLoss.Text.Trim() == "" ? "NULL" : "'" + txePurchaseLoss.Text.Trim() + "'";
                                string MinStock = txeMinStock.Text.Trim() == "" ? "NULL" : "'" + txeMinStock.Text.Trim() + "'";
                                string MaxStock = txeMaxStock.Text.Trim() == "" ? "NULL" : "'" + txeMaxStock.Text.Trim() + "'";
                                string StockShelfLife = txeStockSheifLife.Text.Trim() == "" ? "NULL" : "'" + txeStockSheifLife.Text.Trim() + "'";
                                string StdCost = txeStdCost.Text.Trim() == "" ? "NULL" : "'" + txeStdCost.Text.Trim() + "'";
                                string DefaultUnit = slueDefaultUnit.Text == "" ? "NULL" : "'" + slueDefaultUnit.EditValue.ToString() + "'";

                                string ConvertUnit = slueUnit.Text == "" ? "NULL" : "'" + slueUnit.ToString() + "'";
                                string Factor = txeFactor.Text.Trim() == "" ? "NULL" : "'" + txeFactor.Text.Trim() + "'";

                                string Branch = glueBranch.Text == "" ? "NULL" : "'" + glueBranch.EditValue.ToString() + "'";

                                string StdPrice = txeStdPrice.Text.Trim() == "" ? "0" : txeStdPrice.Text.Trim();
                                string FirstVendor = slueFirstVendor.Text == "" ? "NULL" : "'" + slueFirstVendor.EditValue.ToString() + "'";
                                string QCInspection = txeQCInspection.Text.Trim() == "" ? "0" : txeQCInspection.Text.Trim();

                                if (lblStatus.Text.ToUpper().Trim() == "NEW" || txeID.Text.Trim() == "" || txeID.Text.Trim() == "0")
                                {
                                    sbSQL.Append("  INSERT INTO Items(MaterialType, Code, Description, Composition, WeightOrMoreDetail, ModelNo, ModelName, OIDCATEGORY, OIDSTYLE, OIDCOLOR, OIDSIZE, OIDCUST, BusinessUnit, Season, ClassType, Branch,  ");
                                    sbSQL.Append("       CostSheetNo, StdPrice, FirstVendor, PurchaseType, PurchaseLoss, TaxBenefits, FirstReceiptDate, DefaultVendor, MinStock, MaxStock, StockShelfLife, StdCost, DefaultUnit, ConvertUnit, ConvertFactor, PathFile, LabTestNo, ApprovedLabDate, QCInspection, ");
                                    sbSQL.Append("       CreatedBy, CreatedDate, UpdatedBy, UpdatedDate, OIDCOMPANY) ");
                                    sbSQL.Append("  VALUES('" + MaterialType + "', N'" + glueCode.Text.Trim().Replace("'", "''") + "', N'" + txeDescription.Text.Trim().Replace("'", "''") + "', N'" + txeComposition.Text.Trim().Replace("'", "''") + "', N'" + txeWeight.Text.Trim().Replace("'", "''") + "', N'" + txeModelNo.Text.Trim().Replace("'", "''") + "', N'" + txeModelName.Text.Trim().Replace("'", "''") + "', " + OIDCATEGORY + ", " + OIDSTYLE + ", " + OIDCOLOR + ", " + OIDSIZE + ", " + OIDCUST + ",  ");
                                    sbSQL.Append("         N'" + txeBusinessUnit.Text.Trim().Replace("'", "''") + "', N'" + cbeSeason.Text.Trim() + "', N'" + cbeClass.Text.Trim().Replace("'", "''") + "', " + Branch + ", N'" + txeCostSheet.Text.Trim().Replace("'", "''") + "', '" + StdPrice + "', " + FirstVendor + ", '" + PurchaseType + "', " + PurchaseLoss + ", '" + TaxBenefits + "', '" + Convert.ToDateTime(dteFirstReceiptDate.Text).ToString("yyyy-MM-dd") + "', '" + slueDefaultVendor.EditValue.ToString() + "', ");
                                    sbSQL.Append("         " + MinStock + ", " + MaxStock + ", " + StockShelfLife + ", " + StdCost + ", " + DefaultUnit + ", " + ConvertUnit + ", " + Factor + ", N'" + newFileName + "', N'" + txeLabTestNo.Text.Trim().Replace("'", "''") + "', '" + Convert.ToDateTime(dteApprovedLabDate.Text).ToString("yyyy-MM-dd") + "', '" + QCInspection + "', ");
                                    sbSQL.Append("         '" + strCREATE + "', GETDATE(), '" + strUPDATE + "', GETDATE(), '" + this.Company + "') ");
                                }
                                else if (lblStatus.Text.ToUpper().Trim() == "EDIT")
                                {
                                    sbSQL.Append("  UPDATE Items SET ");
                                    sbSQL.Append("      MaterialType = '" + MaterialType + "', Code = N'" + glueCode.Text.Trim().Replace("'", "''") + "', Description = N'" + txeDescription.Text.Trim().Replace("'", "''") + "', Composition = N'" + txeComposition.Text.Trim().Replace("'", "''") + "', WeightOrMoreDetail = N'" + txeWeight.Text.Trim().Replace("'", "''") + "',  ");
                                    sbSQL.Append("      ModelNo = N'" + txeModelNo.Text.Trim().Replace("'", "''") + "', ModelName = N'" + txeModelName.Text.Trim().Replace("'", "''") + "', OIDCATEGORY = " + OIDCATEGORY + ", OIDSTYLE = " + OIDSTYLE + ", OIDCOLOR = " + OIDCOLOR + ", ");
                                    sbSQL.Append("      OIDSIZE = " + OIDSIZE + ", OIDCUST = " + OIDCUST + ", BusinessUnit = N'" + txeBusinessUnit.Text.Trim().Replace("'", "''") + "', Season = N'" + cbeSeason.Text.Trim() + "', ClassType = N'" + cbeClass.Text.Trim().Replace("'", "''") + "', ");
                                    sbSQL.Append("      Branch = " + Branch + ", CostSheetNo = N'" + txeCostSheet.Text.Trim().Replace("'", "''") + "', StdPrice = '" + StdPrice + "', FirstVendor = " + FirstVendor + ", PurchaseType = '" + PurchaseType + "', PurchaseLoss = " + PurchaseLoss + ", ");
                                    sbSQL.Append("      TaxBenefits = '" + TaxBenefits + "', FirstReceiptDate = '" + Convert.ToDateTime(dteFirstReceiptDate.Text).ToString("yyyy-MM-dd") + "', DefaultVendor = '" + slueDefaultVendor.EditValue.ToString() + "', MinStock = " + MinStock + ", MaxStock = " + MaxStock + ", ");
                                    sbSQL.Append("      StockShelfLife = " + StockShelfLife + ", StdCost = " + StdCost + ", DefaultUnit = " + DefaultUnit + ", ConvertUnit = " + ConvertUnit + ", ConvertFactor = " + Factor + ", PathFile = N'" + newFileName + "', LabTestNo = N'" + txeLabTestNo.Text.Trim().Replace("'", "''") + "', ");
                                    sbSQL.Append("      ApprovedLabDate = '" + Convert.ToDateTime(dteApprovedLabDate.Text).ToString("yyyy-MM-dd") + "', QCInspection = '" + QCInspection + "', UpdatedBy = '" + strUPDATE + "', UpdatedDate = GETDATE() ");
                                    sbSQL.Append("  WHERE(OIDITEM = '" + txeID.Text.Trim() + "') ");
                                }

                                bool chkSAVE = false;
                                try
                                {
                                    chkSAVE = this.DBC.DBQuery(sbSQL).runSQL();
                                }
                                catch (Exception)
                                { }

                                if (chkSAVE == true)
                                {

                                    sbSQL.Clear();
                                    sbSQL.Append("SELECT OIDITEM FROM Items WHERE MaterialType = '" + MaterialType + "' AND Code = N'" + glueCode.Text.Trim().Replace("'", "''") + "'");
                                    string OIDITEM = this.DBC.DBQuery(sbSQL).getString();
                                    //******** save ItemInspection table ********
                                    sbSQL.Clear();
                                    string strCONDQC = "";
                                    int iCQC = 0;
                                    foreach (DataRowView item in clbQC.CheckedItems)
                                    {
                                        if (iCQC != 0)
                                        {
                                            strCONDQC += ", ";
                                        }
                                        strCONDQC += "'" + item["OIDCONDQC"].ToString() + "'";
                                        sbSQL.Append("IF NOT EXISTS(SELECT OIDITEMINSP FROM ItemInspection WHERE OIDITEM = '" + OIDITEM + "' AND OIDCONDQC = '" + item["OIDCONDQC"].ToString() + "') ");
                                        sbSQL.Append(" BEGIN ");
                                        sbSQL.Append("  INSERT INTO ItemInspection(OIDITEM, OIDCONDQC, CreatedBy, CreatedDate) ");
                                        sbSQL.Append("  VALUES('" + OIDITEM + "', '" + item["OIDCONDQC"].ToString() + "', '" + strCREATE + "', GETDATE()) ");
                                        sbSQL.Append(" END ");
                                        iCQC++;
                                    }

                                    if (strCONDQC == "")
                                    {
                                        sbSQL.Append("DELETE FROM ItemInspection WHERE (OIDITEM = '" + OIDITEM + "')  ");
                                    }
                                    else
                                    {
                                        sbSQL.Append("DELETE FROM ItemInspection WHERE (OIDITEM = '" + OIDITEM + "') AND (OIDCONDQC NOT IN (" + strCONDQC + "))  ");
                                    }

                                    //******** save ItemVendor table ************

                                    if (dtVendor != null)
                                    {
                                        foreach (DataRow dr in dtVendor.Rows) // search whole table
                                        {
                                            string Price = txePrice.Text.Trim() == "" ? "NULL" : "'" + txePrice.Text.Trim() + "'";
                                            string LotSize = dr["LotSize"].ToString() == "" ? "NULL" : "'" + dr["LotSize"].ToString() + "'";
                                            string ProductionLead = dr["ProductionLead"].ToString() == "" ? "NULL" : "'" + dr["ProductionLead"].ToString() + "'";
                                            string DeliveryLead = dr["DeliveryLead"].ToString() == "" ? "NULL" : "'" + dr["DeliveryLead"].ToString() + "'";
                                            string ArrivalLead = dr["ArrivalLead"].ToString() == "" ? "NULL" : "'" + dr["ArrivalLead"].ToString() + "'";
                                            string POCancelPeriod = dr["POCancelPeriod"].ToString() == "" ? "NULL" : "'" + dr["POCancelPeriod"].ToString() + "'";
                                            string PurchaseLots1 = dr["PurchaseLots1"].ToString() == "" ? "NULL" : "'" + dr["PurchaseLots1"].ToString() + "'";
                                            string PurchaseLots2 = dr["PurchaseLots2"].ToString() == "" ? "NULL" : "'" + dr["PurchaseLots2"].ToString() + "'";
                                            string PurchaseLots3 = dr["PurchaseLots3"].ToString() == "" ? "NULL" : "'" + dr["PurchaseLots3"].ToString() + "'";

                                            if (dr["OIDVENDItem"].ToString() == "") //Insert
                                            {
                                                sbSQL.Append("INSERT INTO ItemVendor(OIDVEND, OIDITEM, ");
                                                if (dr["OIDVEND"].ToString() == slueFirstVendor.EditValue.ToString())
                                                {
                                                    sbSQL.Append("MatDetails, MatCode, SMPLLotNo, Price, Currency, ");
                                                }
                                                sbSQL.Append("  LotSize, ProductionLead, DeliveryLead, ArrivalLead, POCancelPeriod, PurchaseLots1, PurchaseLots2, PurchaseLots3, Remark) ");
                                                sbSQL.Append(" VALUES('" + dr["OIDVEND"].ToString() + "', '" + OIDITEM + "',  ");
                                                if (dr["OIDVEND"].ToString() == slueFirstVendor.EditValue.ToString())
                                                {
                                                    sbSQL.Append("N'" + txeMatDetails.Text.Trim().Replace("'", "''") + "', N'" + txeMatCode.Text.Trim().Replace("'", "''") + "', N'" + txeSMPLLotNo.Text.Trim().Replace("'", "''") + "', " + Price + ", N'" + slueCurrency.EditValue.ToString() + "',  ");
                                                }
                                                sbSQL.Append(" " + LotSize + ", " + ProductionLead + ", " + DeliveryLead + ", " + ArrivalLead + ", " + POCancelPeriod + ", " + PurchaseLots1 + ", " + PurchaseLots2 + ", " + PurchaseLots3 + ", N'" + dr["Remark"].ToString().Replace("'", "''") + "')  ");
                                            }
                                            else //Update
                                            {
                                                sbSQL.Append("UPDATE ItemVendor SET ");
                                                sbSQL.Append("  OIDVEND = '" + dr["OIDVEND"].ToString() + "', OIDITEM = '" + OIDITEM + "', ");
                                                if (dr["OIDVEND"].ToString() == slueFirstVendor.EditValue.ToString())
                                                {
                                                    sbSQL.Append("MatDetails = N'" + txeMatDetails.Text.Trim().Replace("'", "''") + "', MatCode = N'" + txeMatCode.Text.Trim().Replace("'", "''") + "', SMPLLotNo = N'" + txeSMPLLotNo.Text.Trim().Replace("'", "''") + "', Price = " + Price + ", Currency = N'" + slueCurrency.EditValue.ToString() + "',  ");
                                                }
                                                sbSQL.Append("LotSize = " + LotSize + ", ProductionLead = " + ProductionLead + ", DeliveryLead = " + DeliveryLead + ", ArrivalLead = " + ArrivalLead + ", POCancelPeriod = " + POCancelPeriod + ", PurchaseLots1 = " + PurchaseLots1 + ", PurchaseLots2 = " + PurchaseLots2 + ", PurchaseLots3 = " + PurchaseLots3 + ", Remark = N'" + dr["Remark"].ToString().Replace("'", "''") + "'  ");
                                                sbSQL.Append("WHERE (OIDVENDItem = '" + dr["OIDVENDItem"].ToString() + "') ");
                                            }
                                        }
                                    }

                                    //MessageBox.Show("2");
                                    //MessageBox.Show(sbSQL.ToString());
                                    if (sbSQL.Length > 0)
                                    {
                                        try
                                        {
                                            chkSAVE = this.DBC.DBQuery(sbSQL).runSQL();
                                            if (chkSAVE == true)
                                            {
                                                FUNC.msgInfo("Save complete.");
                                                bbiNew.PerformClick();
                                            }
                                        }
                                        catch (Exception)
                                        { }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            else if (tabbedControlGroup1.SelectedTabPage == lcgImport) //Import File
            {
                if (txeFilePath.Text.Trim() == "")
                {
                    FUNC.msgWarning("Please select excel file.");
                    txeFilePath.Focus();
                }
                else if (cbeSheet.Text.Trim() == "")
                {
                    FUNC.msgWarning("Please select excel sheet.");
                    cbeSheet.Focus();
                }
                else
                {
                    if (FUNC.msgQuiz("Confirm save excel file import data ?") == true)
                    {
                        StringBuilder sbSQL = new StringBuilder();

                        bool chkSAVE = false;

                        IWorkbook workbook = spsImport.Document;
                        Worksheet WSHEET = workbook.Worksheets[0];

                        lciPregressSave.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                        pbcSave.Properties.Step = 1;
                        pbcSave.Properties.PercentView = true;
                        pbcSave.Properties.Maximum = WSHEET.GetDataRange().RowCount;
                        pbcSave.Properties.Minimum = 0;
                        pbcSave.EditValue = 0;

                        string Customer = "--";
                        string OIDCUST = "0";

                        string Unit = "--";
                        string OIDUNIT = "0";

                        string Vendor = "--";
                        string OIDVEND = "0";

                        string Category = "--";
                        string OIDCATEGORY = "0";

                        string Style = "--";
                        string OIDSTYLE = "0";

                        string Color = "--";
                        string OIDCOLOR = "0";

                        string Size = "--";
                        string OIDSIZE = "0";

                        string ConvertUnit = "--";
                        string OIDCONVERTUNIT = "0";

                        for (int i = 4; i < WSHEET.GetDataRange().RowCount; i++)
                        {
                            string VENDOR_TYPE = "";
                            string COLOR_TYPE = "";
                            string MATERIAL_TYPE = WSHEET.Rows[i][0].DisplayText.ToString().Trim();
                            if (IsNumeric(MATERIAL_TYPE) == false)
                            {
                                string chkTYPE = MATERIAL_TYPE.Trim().ToUpper().Replace(" ", "");
                                if (chkTYPE == "F/G" ||
                                    chkTYPE == "GOODS" ||
                                    chkTYPE == "FINISHGOOD" ||
                                    chkTYPE == "FINISHGOODS" ||
                                    chkTYPE == "FINISHEDGOOD" ||
                                    chkTYPE == "FINISHEDGOODS" ||
                                    chkTYPE == "FINISH" ||
                                    chkTYPE == "FINISHED")
                                {
                                    MATERIAL_TYPE = "0";
                                }
                                else if (chkTYPE == "FABRIC" || chkTYPE == "FABRICS" || chkTYPE == "FB" || chkTYPE == "F/B")
                                {
                                    MATERIAL_TYPE = "1";
                                }
                                else if (chkTYPE == "ACCESSORY" || chkTYPE == "ACCESSORIES" || chkTYPE == "ACC" || chkTYPE == "ACS")
                                {
                                    MATERIAL_TYPE = "2";
                                }
                                else if (chkTYPE == "PACKAGING" || chkTYPE == "PACKAGE" || chkTYPE == "PACK")
                                {
                                    MATERIAL_TYPE = "3";
                                }
                                else if (chkTYPE == "SAMPLE" || chkTYPE == "SP")
                                {
                                    MATERIAL_TYPE = "4";
                                }
                                else if (chkTYPE == "OTHER")
                                {
                                    MATERIAL_TYPE = "9";
                                }
                                else
                                {
                                    MATERIAL_TYPE = "-1";
                                }
                            }
                            VENDOR_TYPE = MATERIAL_TYPE;
                            COLOR_TYPE = MATERIAL_TYPE;
                            string CODE = WSHEET.Rows[i][1].DisplayText.ToString().Trim();

                            if (MATERIAL_TYPE != "" && CODE != "")
                            {
                                string DESCRIPTION = WSHEET.Rows[i][2].DisplayText.ToString().Trim().Replace("'", "''");
                                string GROUP_BOI = WSHEET.Rows[i][3].DisplayText.ToString().Trim().Replace("'", "''");
                                string GROUP_SECTION = WSHEET.Rows[i][4].DisplayText.ToString().Trim().Replace("'", "''");
                                string COMPOSITION = WSHEET.Rows[i][5].DisplayText.ToString().Trim().Replace("'", "''");
                                string MORE_DETAIL = WSHEET.Rows[i][6].DisplayText.ToString().Trim().Replace("'", "''");
                                string MODEL = WSHEET.Rows[i][7].DisplayText.ToString().Trim().Replace("'", "''");
                                string MODEL_NAME = WSHEET.Rows[i][8].DisplayText.ToString().Trim().Replace("'", "''");
                                string CATEGORY = WSHEET.Rows[i][9].DisplayText.ToString().Trim().Replace("'", "''");
                                string STYLE = WSHEET.Rows[i][10].DisplayText.ToString().Trim().Replace("'", "''");
                                string COLOR = WSHEET.Rows[i][11].DisplayText.ToString().Trim().Replace("'", "''");
                                string SIZE = WSHEET.Rows[i][12].DisplayText.ToString().Trim().Replace("'", "''");
                                string SEASON = WSHEET.Rows[i][13].DisplayText.ToString().Trim().Replace("'", "''");
                                string CUSTOMER = WSHEET.Rows[i][14].DisplayText.ToString().Trim().Replace("'", "''");
                                string UNIT = WSHEET.Rows[i][15].DisplayText.ToString().Trim().Replace("'", "''");
                                string CONVERTUNIT = WSHEET.Rows[i][19].DisplayText.ToString().Trim().Replace("'", "''");
                                string PRICE = WSHEET.Rows[i][16].DisplayText.ToString().Trim().Replace("'", "''");
                                PRICE = PRICE == "" ? "0" : PRICE;
                                PRICE = IsNumeric(PRICE) == false ? "0" : PRICE;
                                string SUPPLIER = WSHEET.Rows[i][17].DisplayText.ToString().Trim().Replace("'", "''");
                                string PURCHASE_TYPE = WSHEET.Rows[i][18].DisplayText.ToString().Trim().Replace("'", "''");
                                if (IsNumeric(PURCHASE_TYPE) == false)
                                {
                                    string chkPTYPE = PURCHASE_TYPE.Trim().ToUpper().Replace(" ", "");
                                    if (chkPTYPE == "LOCAL")
                                    {
                                        PURCHASE_TYPE = "0";
                                    }
                                    else if (chkPTYPE == "IMPORT")
                                    {
                                        PURCHASE_TYPE = "1";
                                    }
                                    else if (chkPTYPE == "OTHER")
                                    {
                                        PURCHASE_TYPE = "9";
                                    }
                                    else
                                    {
                                        PURCHASE_TYPE = "-1";
                                    }
                                }

                                string MINIMUM = WSHEET.Rows[i][21].DisplayText.ToString().Trim().Replace("'", "''");
                                if (IsNumeric(MINIMUM) == false)
                                {
                                    MINIMUM = "0";
                                }

                                string MAXIMUM = WSHEET.Rows[i][22].DisplayText.ToString().Trim().Replace("'", "''");
                                if (IsNumeric(MAXIMUM) == false)
                                {
                                    MAXIMUM = "0";
                                }

                                if (SUPPLIER == null) SUPPLIER = "";
                                if (Vendor != SUPPLIER.Replace(" ", "").Replace(".", "").Replace(",", ""))
                                {
                                    Vendor = SUPPLIER.Replace(" ", "").Replace(".", "").Replace(",", "");
                                    StringBuilder sbVENDOR = new StringBuilder();
                                    sbVENDOR.Append("IF NOT EXISTS(SELECT OIDVEND FROM Vendor WHERE (Name=N'" + SUPPLIER + "')) ");
                                    sbVENDOR.Append(" BEGIN ");
                                    sbVENDOR.Append("   INSERT INTO Vendor(Code, Name, VendorType, CreatedBy, CreatedDate, UpdatedBy, UpdatedDate) VALUES(N'', N'" + SUPPLIER + "', '" + VENDOR_TYPE + "', '" + UserLogin.OIDUser + "', GETDATE(), '" + UserLogin.OIDUser + "', GETDATE()) ");
                                    sbVENDOR.Append(" END  ");
                                    sbVENDOR.Append("SELECT OIDVEND FROM Vendor WHERE (Name=N'" + SUPPLIER + "')  ");
                                    OIDVEND = this.DBC.DBQuery(sbVENDOR).getString();
                                }

                                if (CATEGORY == null) CATEGORY = "";
                                if (Category != CATEGORY.Replace(" ", "").Replace(".", "").Replace(",", ""))
                                {
                                    Category = CATEGORY.Replace(" ", "").Replace(".", "").Replace(",", "");
                                    StringBuilder sbCATEGORY = new StringBuilder();
                                    sbCATEGORY.Append("IF NOT EXISTS(SELECT OIDGCATEGORY FROM GarmentCategory WHERE (CategoryName=N'" + CATEGORY + "')) ");
                                    sbCATEGORY.Append(" BEGIN ");
                                    sbCATEGORY.Append("   INSERT INTO GarmentCategory(CategoryName, CreatedBy, CreatedDate) VALUES(N'" + CATEGORY + "', '" + UserLogin.OIDUser + "', GETDATE()) ");
                                    sbCATEGORY.Append(" END  ");
                                    sbCATEGORY.Append("SELECT OIDGCATEGORY FROM GarmentCategory WHERE (CategoryName=N'" + CATEGORY + "')  ");
                                    OIDCATEGORY = this.DBC.DBQuery(sbCATEGORY).getString();
                                }

                                if (STYLE == null) STYLE = "";
                                if (Style != STYLE.Replace(" ", "").Replace(".", "").Replace(",", ""))
                                {
                                    Style = STYLE.Replace(" ", "").Replace(".", "").Replace(",", "");
                                    StringBuilder sbSTYLE = new StringBuilder();
                                    sbSTYLE.Append("IF NOT EXISTS(SELECT OIDSTYLE FROM ProductStyle WHERE (OIDGCATEGORY='" + OIDCATEGORY + "') AND (StyleName=N'" + STYLE + "')) ");
                                    sbSTYLE.Append(" BEGIN ");
                                    sbSTYLE.Append("   INSERT INTO ProductStyle(StyleName, OIDGCATEGORY, CreatedBy, CreatedDate) VALUES(N'" + STYLE + "', '" + OIDCATEGORY + "', '" + UserLogin.OIDUser + "', GETDATE()) ");
                                    sbSTYLE.Append(" END  ");
                                    sbSTYLE.Append("SELECT OIDSTYLE FROM ProductStyle WHERE (OIDGCATEGORY='" + OIDCATEGORY + "') AND (StyleName=N'" + STYLE + "')  ");
                                    OIDSTYLE = this.DBC.DBQuery(sbSTYLE).getString();
                                }

                                if (COLOR == null) COLOR = "";
                                if (Color != COLOR.Replace(" ", "").Replace(".", "").Replace(",", ""))
                                {
                                    Color = COLOR.Replace(" ", "").Replace(".", "").Replace(",", "");
                                    StringBuilder sbCOLOR = new StringBuilder();
                                    sbCOLOR.Append("IF NOT EXISTS(SELECT OIDCOLOR FROM ProductColor WHERE (ColorName=N'" + COLOR + "') AND (ColorType='" + COLOR_TYPE + "')) ");
                                    sbCOLOR.Append(" BEGIN ");
                                    sbCOLOR.Append("   INSERT INTO ProductColor(ColorNo, ColorName, ColorType, CreatedBy, CreatedDate) VALUES(N'" + COLOR + "', N'" + COLOR + "', '" + COLOR_TYPE + "', '" + UserLogin.OIDUser + "', GETDATE()) ");
                                    sbCOLOR.Append(" END  ");
                                    sbCOLOR.Append("SELECT OIDCOLOR FROM ProductColor WHERE (ColorName=N'" + COLOR + "') AND (ColorType='" + COLOR_TYPE + "')  ");
                                    OIDCOLOR = this.DBC.DBQuery(sbCOLOR).getString();
                                }

                                if (SIZE == null) SIZE = "";
                                if (Size != SIZE.Replace(" ", "").Replace(".", "").Replace(",", ""))
                                {
                                    Size = SIZE.Replace(" ", "").Replace(".", "").Replace(",", "");
                                    StringBuilder sbSIZE = new StringBuilder();
                                    sbSIZE.Append("IF NOT EXISTS(SELECT OIDSIZE FROM ProductSize WHERE (SizeName=N'" + SIZE + "')) ");
                                    sbSIZE.Append(" BEGIN ");
                                    sbSIZE.Append("   INSERT INTO ProductSize(SizeNo, SizeName, CreatedBy, CreatedDate) VALUES(N'" + SIZE + "', N'" + SIZE + "', '" + UserLogin.OIDUser + "', GETDATE()) ");
                                    sbSIZE.Append(" END  ");
                                    sbSIZE.Append("SELECT OIDSIZE FROM ProductSize WHERE (SizeName=N'" + SIZE + "')  ");
                                    OIDSIZE = this.DBC.DBQuery(sbSIZE).getString();
                                }

                                if (CUSTOMER == null) CUSTOMER = "";
                                if (Customer != CUSTOMER.Replace(" ", "").Replace(".", "").Replace(",", ""))
                                {
                                    Customer = CUSTOMER.Replace(" ", "").Replace(".", "").Replace(",", "");
                                    string CustomerCode = Customer.Length > 20 ? Customer.Substring(0, 20) : Customer;
                                    string CustomerShort = Customer.Length > 10 ? Customer.Substring(0, 10) : Customer;
                                    StringBuilder sbCUST = new StringBuilder();
                                    sbCUST.Append("IF NOT EXISTS(SELECT OIDCUST FROM Customer WHERE Name = N'" + CUSTOMER + "') ");
                                    sbCUST.Append(" BEGIN ");
                                    sbCUST.Append("   INSERT INTO Customer(Code, Name, ShortName, CustomerType) VALUES(N'" + CustomerCode + "', N'" + CUSTOMER + "', N'" + CustomerShort + "', '0') ");
                                    sbCUST.Append(" END ");
                                    sbCUST.Append("SELECT TOP(1) OIDCUST FROM Customer WHERE Name LIKE N'%" + Customer + "%' ");
                                    OIDCUST = this.DBC.DBQuery(sbCUST).getString();
                                }

                                if (UNIT == null) UNIT = "";
                                if (Unit != UNIT.Replace(" ", "").Replace(".", "").Replace(",", ""))
                                {
                                    Unit = UNIT.Replace(" ", "").Replace(".", "").Replace(",", "");
                                    StringBuilder sbUNIT = new StringBuilder();
                                    sbUNIT.Append("IF NOT EXISTS(SELECT OIDUNIT FROM Unit WHERE (UnitName=N'" + UNIT + "')) ");
                                    sbUNIT.Append(" BEGIN ");
                                    sbUNIT.Append("   INSERT INTO Unit(UnitName, CreatedBy, CreatedDate, UpdatedBy, UpdatedDate) VALUES(N'" + UNIT + "', '" + UserLogin.OIDUser + "', GETDATE(), '" + UserLogin.OIDUser + "', GETDATE()) ");
                                    sbUNIT.Append(" END  ");
                                    sbUNIT.Append("SELECT OIDUNIT FROM Unit WHERE (UnitName=N'" + UNIT + "')  ");
                                    OIDUNIT = this.DBC.DBQuery(sbUNIT).getString();
                                }

                                if (CONVERTUNIT == null) CONVERTUNIT = "";
                                if (ConvertUnit != CONVERTUNIT.Replace(" ", "").Replace(".", "").Replace(",", ""))
                                {
                                    ConvertUnit = CONVERTUNIT.Replace(" ", "").Replace(".", "").Replace(",", "");
                                    StringBuilder sbUNIT = new StringBuilder();
                                    sbUNIT.Append("IF NOT EXISTS(SELECT OIDUNIT FROM Unit WHERE (UnitName=N'" + CONVERTUNIT + "')) ");
                                    sbUNIT.Append(" BEGIN ");
                                    sbUNIT.Append("   INSERT INTO Unit(UnitName, CreatedBy, CreatedDate, UpdatedBy, UpdatedDate) VALUES(N'" + CONVERTUNIT + "', '" + UserLogin.OIDUser + "', GETDATE(), '" + UserLogin.OIDUser + "', GETDATE()) ");
                                    sbUNIT.Append(" END  ");
                                    sbUNIT.Append("SELECT OIDUNIT FROM Unit WHERE (UnitName=N'" + CONVERTUNIT + "')  ");
                                    OIDCONVERTUNIT = this.DBC.DBQuery(sbUNIT).getString();
                                }

                                string FACTOR = WSHEET.Rows[i][20].DisplayText.ToString().Trim().Replace("'", "''");
                                if (IsNumeric(FACTOR) == false)
                                {
                                    FACTOR = "0";
                                }

                                sbSQL.Clear();
                                sbSQL.Append("IF NOT EXISTS(SELECT Code FROM Items WHERE Code = N'" + CODE + "') ");
                                sbSQL.Append(" BEGIN ");
                                sbSQL.Append("   INSERT INTO Items(MaterialType, Code, Description, GroupBOI, GroupSection, Composition, WeightOrMoreDetail, ModelNo, ModelName, OIDCATEGORY, OIDSTYLE, OIDCOLOR, OIDSIZE, Season, OIDCUST, BusinessUnit, DefaultUnit, ConvertUnit, ConvertFactor, StdPrice, FirstVendor, DefaultVendor, PurchaseType, MinStock, MaxStock, Branch, OIDCOMPANY, OIDBranch, OIDDEPT, CreatedBy, CreatedDate, UpdatedBy, UpdatedDate) ");
                                sbSQL.Append("   SELECT '" + MATERIAL_TYPE + "' AS MaterialType,  ");
                                sbSQL.Append("          N'" + CODE + "' AS Code,  ");
                                sbSQL.Append("          N'" + DESCRIPTION + "' AS Description,  ");
                                sbSQL.Append("          '" + GROUP_BOI + "' AS GroupBOI,  ");
                                sbSQL.Append("          '" + GROUP_SECTION + "' AS GroupSection,  ");
                                sbSQL.Append("          N'" + COMPOSITION + "' AS Composition,  ");
                                sbSQL.Append("          N'" + MORE_DETAIL + "' AS WeightOrMoreDetail,  ");
                                sbSQL.Append("          N'" + MODEL + "' AS ModelNo,  ");
                                sbSQL.Append("          N'" + MODEL_NAME + "' AS ModelName,  ");
                                sbSQL.Append("          '" + OIDCATEGORY + "' AS OIDCATEGORY,  ");
                                sbSQL.Append("          '" + OIDSTYLE + "' AS OIDSTYLE,  ");
                                sbSQL.Append("          '" + OIDCOLOR + "' AS OIDCOLOR,  ");
                                sbSQL.Append("          '" + OIDSIZE + "' AS OIDSIZE,  ");
                                sbSQL.Append("          N'" + SEASON + "' AS Season,  ");
                                sbSQL.Append("          '" + OIDCUST + "' AS OIDCUST,  ");
                                sbSQL.Append("          N'" + UNIT + "' AS BusinessUnit,  ");
                                sbSQL.Append("          '" + OIDUNIT + "' AS DefaultUnit,  ");
                                sbSQL.Append("          '" + OIDCONVERTUNIT + "' AS ConvertUnit,  ");
                                sbSQL.Append("          '" + FACTOR + "' AS ConvertFactor,  ");
                                sbSQL.Append("          '" + PRICE + "' AS StdPrice,  ");
                                sbSQL.Append("          '" + OIDVEND + "' AS FirstVendor,  ");
                                sbSQL.Append("          '" + OIDVEND + "' AS DefaultVendor,  ");
                                sbSQL.Append("          '" + PURCHASE_TYPE + "' AS PurchaseType,  ");
                                sbSQL.Append("          '" + MINIMUM + "' AS MinStock, ");
                                sbSQL.Append("          '" + MAXIMUM + "' AS MaxStock,  ");
                                sbSQL.Append("          '" + UserLogin.OIDBranch + "' AS Branch,  ");
                                sbSQL.Append("          '" + UserLogin.OIDCompany + "' AS OIDCOMPANY,  ");
                                sbSQL.Append("          '" + UserLogin.OIDBranch + "' AS OIDBranch,  ");
                                sbSQL.Append("          '" + UserLogin.OIDDept + "' AS OIDDEPT,  ");
                                sbSQL.Append("          '" + UserLogin.OIDUser + "' AS CreatedBy,  ");
                                sbSQL.Append("          GETDATE() AS CreatedDate,  ");
                                sbSQL.Append("          '" + UserLogin.OIDUser + "' AS UpdatedBy,  ");
                                sbSQL.Append("          GETDATE() AS UpdatedDate   ");
                                sbSQL.Append(" END ");
                                sbSQL.Append("ELSE ");
                                sbSQL.Append(" BEGIN ");
                                sbSQL.Append("   UPDATE Items SET ");
                                sbSQL.Append("      MaterialType='" + MATERIAL_TYPE + "',  ");
                                sbSQL.Append("      Description=N'" + DESCRIPTION + "',  ");
                                sbSQL.Append("      GroupBOI='" + GROUP_BOI + "',  ");
                                sbSQL.Append("      GroupSection='" + GROUP_SECTION + "',  ");
                                sbSQL.Append("      Composition=N'" + COMPOSITION + "',  ");
                                sbSQL.Append("      WeightOrMoreDetail=N'" + MORE_DETAIL + "',  ");
                                sbSQL.Append("      ModelNo=N'" + MODEL + "',  ");
                                sbSQL.Append("      ModelName=N'" + MODEL_NAME + "',  ");
                                sbSQL.Append("      OIDCATEGORY='" + OIDCATEGORY + "',  ");
                                sbSQL.Append("      OIDSTYLE='" + OIDSTYLE + "',  ");
                                sbSQL.Append("      OIDCOLOR='" + OIDCOLOR + "',  ");
                                sbSQL.Append("      OIDSIZE='" + OIDSIZE + "',  ");
                                sbSQL.Append("      Season=N'" + SEASON + "',  ");
                                sbSQL.Append("      OIDCUST='" + OIDCUST + "',  ");
                                sbSQL.Append("      BusinessUnit=N'" + UNIT + "',  ");
                                sbSQL.Append("      DefaultUnit='" + OIDUNIT + "',  ");
                                sbSQL.Append("      ConvertUnit='" + OIDCONVERTUNIT + "',  ");
                                sbSQL.Append("      ConvertFactor='" + FACTOR + "',  ");
                                sbSQL.Append("      StdPrice='" + PRICE + "',  ");
                                sbSQL.Append("      FirstVendor='" + OIDVEND + "',  ");
                                sbSQL.Append("      DefaultVendor='" + OIDVEND + "',  ");
                                sbSQL.Append("      PurchaseType='" + PURCHASE_TYPE + "',  ");
                                sbSQL.Append("      MinStock='" + MINIMUM + "', ");
                                sbSQL.Append("      MaxStock='" + MAXIMUM + "',  ");
                                sbSQL.Append("      Branch='" + UserLogin.OIDBranch + "',  ");
                                sbSQL.Append("      OIDCOMPANY='" + UserLogin.OIDCompany + "',  ");
                                sbSQL.Append("      OIDBranch='" + UserLogin.OIDBranch + "',  ");
                                sbSQL.Append("      OIDDEPT='" + UserLogin.OIDDept + "',  ");
                                sbSQL.Append("      UpdatedBy='" + UserLogin.OIDUser + "',  ");
                                sbSQL.Append("      UpdatedDate = GETDATE()  ");
                                sbSQL.Append("    WHERE(Code = N'" + CODE + "') ");
                                sbSQL.Append(" END   ");

                                //memoEdit1.EditValue = sbSQL.ToString();
                                //break;
                                //MessageBox.Show(sbSQL.ToString());
                                try
                                {
                                    chkSAVE = this.DBC.DBQuery(sbSQL).runSQL();
                                    if (chkSAVE == false)
                                    {
                                        break;
                                    }
                                    else
                                    {
                                        pbcSave.PerformStep();
                                        pbcSave.Update();
                                    }
                                }
                                catch (Exception)
                                { }

                            }

                        }

                        if (chkSAVE == true)
                        {
                            lciPregressSave.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                            FUNC.msgInfo("Save complete.");
                            bbiNew.PerformClick();
                        }

                        //if (sbSQL.Length > 0)
                        //{
                        //    //MessageBox.Show(sbSQL.ToString());
                        //    try
                        //    {
                        //        chkSAVE = new DBQuery(sbSQL).runSQL();
                        //        if (chkSAVE == true)
                        //        {
                        //            FUNC.msgInfo("Save complete.");
                        //            bbiNew.PerformClick();
                        //        }
                        //    }
                        //    catch (Exception)
                        //    { }
                        //}

                    }
                }
            }
        }

        //private void selectMaterial(int value)
        //{
        //    rgMaterial.EditValue = value;
        //    //switch (value)
        //    //{
        //    //    case 0:
        //    //        rgMaterial.SelectedIndex = 0;
        //    //        break;
        //    //    case 1:
        //    //        rgMaterial.SelectedIndex = 1;
        //    //        break;
        //    //    case 2:
        //    //        rgMaterial.SelectedIndex = 2;
        //    //        break;
        //    //    case 3:
        //    //        rgMaterial.SelectedIndex = 3;
        //    //        break;
        //    //    case 4:
        //    //        rgMaterial.SelectedIndex = 4;
        //    //        break;
        //    //    case 9:
        //    //        rgMaterial.SelectedIndex = 5;
        //    //        break;
        //    //    default:
        //    //        rgMaterial.SelectedIndex = -1;
        //    //        break;
        //    //}
        //}

        //private void selectPurchase(int value)
        //{
        //    rgPurchase.EditValue = value;
        //    //switch (value)
        //    //{
        //    //    case 0:
        //    //        rgPurchase.SelectedIndex = 0;
        //    //        break;
        //    //    case 1:
        //    //        rgPurchase.SelectedIndex = 1;
        //    //        break;
        //    //    case 9:
        //    //        rgPurchase.SelectedIndex = 2;
        //    //        break;
        //    //    default:
        //    //        rgPurchase.SelectedIndex = -1;
        //    //        break;
        //    //}
        //}

        //private void selectTax(int value)
        //{
        //    rgTax.EditValue = value;
        //    //switch (value)
        //    //{
        //    //    case 1:
        //    //        rgTax.SelectedIndex = 0;
        //    //        break;
        //    //    case 2:
        //    //        rgTax.SelectedIndex = 1;
        //    //        break;
        //    //    case 9:
        //    //        rgTax.SelectedIndex = 2;
        //    //        break;
        //    //    default:
        //    //        rgTax.SelectedIndex = -1;
        //    //        break;
        //    //}
        //}

        //private void selectZone(int value)
        //{
        //    rgZone.EditValue = value;
        //    //switch (value)
        //    //{
        //    //    case 0:
        //    //        rgZone.SelectedIndex = 0;
        //    //        break;
        //    //    case 1:
        //    //        rgZone.SelectedIndex = 1;
        //    //        break;
        //    //    case 2:
        //    //        rgZone.SelectedIndex = 2;
        //    //        break;
        //    //    default:
        //    //        rgZone.SelectedIndex = -1;
        //    //        break;
        //    //}
        //}

        private void rgMaterial_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lblStatus.Text == "NEW")
            {
                NewData();
                LoadDataMeterial();
                //glueCode.Focus();
                txeID.Focus();
            }
        }

        private void gvVendor_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            if (sender is GridView)
            {
                GridView gView = (GridView)sender;
                if (!gView.IsValidRowHandle(e.RowHandle)) return;
                int parent = gView.GetParentRowHandle(e.RowHandle);
                if (gView.IsGroupRow(parent))
                {
                    for (int i = 0; i < gView.GetChildRowCount(parent); i++)
                    {
                        if (gView.GetChildRowHandle(parent, i) == e.RowHandle)
                        {
                            e.Appearance.BackColor = i % 2 == 0 ? Color.AliceBlue : Color.White;
                        }
                    }
                }
                else
                {
                    e.Appearance.BackColor = e.RowHandle % 2 == 0 ? Color.AliceBlue : Color.White;
                }
            }
        }

        private void LoadVendorData()
        {
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT OIDVENDItem, OIDVEND, Code AS Vendor, Name AS VendorName, OIDITEM, LotSize, ProductionLead, DeliveryLead, ArrivalLead, POCancelPeriod, PurchaseLots1, PurchaseLots2, PurchaseLots3, Remark ");
            sbSQL.Append("FROM ItemVendor AS IVD ");
            sbSQL.Append("     CROSS APPLY(SELECT Code, Name FROM Vendor WHERE OIDVEND = IVD.OIDVEND) AS VD ");
            sbSQL.Append("WHERE (OIDITEM = '" + txeID.Text.Trim() + "') ");
            sbSQL.Append("ORDER BY OIDVENDItem ");
            new ObjDE.setGridControl(gcVendor, gvVendor, sbSQL).getData(false, false, false, true);
            dtVendor = this.DBC.DBQuery(sbSQL).getDataTable();
            if (gvVendor.Columns.Count > 0)
            {
                //เปลี่ยนชื่อ Column ใน DataTable ให้ตรงกับ DataGridView
                for (int ii = 0; ii < gvVendor.Columns.Count; ii++)
                {
                    try
                    {
                        dtVendor.Columns[ii].ColumnName = gvVendor.Columns[ii].FieldName;
                        dtVendor.Columns[ii].AllowDBNull = true;
                    }
                    catch (Exception) { }
                    ii++;
                }
            }
            gvVendor.Columns[0].Visible = false;
            gvVendor.Columns[1].Visible = false;
            gvVendor.Columns[4].Visible = false;
        }

        private void LoadCode(string strCODE, bool ShowMsg = true)
        {
            //txeID.Text = "";
            //lblStatus.Text = "NEW";
            //lblStatus.ForeColor = Color.White;
            //lblStatus.BackColor = Color.Green;
            //layoutControlItem78.AppearanceItemCaption.BackColor = Color.Green;
            NewData(false);

            glueCode.EditValue = strCODE;
            txeOldCode.Text = glueCode.EditValue.ToString();
            selCode = glueCode.EditValue.ToString();
            txeID.Text = glueCode.EditValue.ToString();

            //strCODE = strCODE.ToUpper().Trim();
            StringBuilder sbSQL = new StringBuilder();

            selCode = strCODE;
            sbSQL.Clear();
            //StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT OIDITEM, MaterialType, Code, Description, Composition, WeightOrMoreDetail, ModelNo, ModelName, OIDCATEGORY, OIDSTYLE, OIDCOLOR, OIDSIZE, OIDCUST, BusinessUnit, Season, ClassType, Branch, CostSheetNo,  ");
            sbSQL.Append("       FORMAT(StdPrice, '###0.####') AS StdPrice, FirstVendor, PurchaseType, TaxBenefits, FORMAT(PurchaseLoss, '###0.##') AS PurchaseLoss, FirstReceiptDate, DefaultVendor, MinStock, MaxStock, StockShelfLife, FORMAT(StdCost, '###0.####') AS StdCost, ");
            sbSQL.Append("       DefaultUnit, PathFile, LabTestNo, ApprovedLabDate, QCInspection, CreatedBy, CreatedDate, ");
            sbSQL.Append("       UpdatedBy, UpdatedDate, ConvertUnit, ConvertFactor ");
            sbSQL.Append("FROM Items ");
            sbSQL.Append("WHERE (OIDITEM = '" + strCODE + "') ");
            string[] arrItem = this.DBC.DBQuery(sbSQL).getMultipleValue();
            if (arrItem.Length > 0)
            {
                lblStatus.Text = "EDIT";
                lblStatus.ForeColor = Color.White;
                lblStatus.BackColor = Color.Navy;
                layoutControlItem78.AppearanceItemCaption.BackColor = Color.Navy;
                //return;
                if (arrItem[1] != "" && arrItem[1] != "8")
                    rgMaterial.EditValue = Convert.ToInt32(arrItem[1]);
                else if (arrItem[1] == "8")
                    if (arrItem[2].ToUpper().Trim().Length >= 5)
                        if (arrItem[2].ToUpper().Trim().Substring(0, 5) == "TMPFB" && rgMaterial.EditValue.ToString() != "1")
                            rgMaterial.EditValue = 1;
                        else if (arrItem[2].ToUpper().Trim().Substring(0, 5) == "TMPMT" && rgMaterial.EditValue.ToString() != "2" && rgMaterial.EditValue.ToString() != "3")
                            rgMaterial.EditValue = 2;


                //txeOldCode.Text = arrItem[0].ToUpper().Trim();
                //glueCode.EditValue = arrItem[0].ToUpper().Trim();
                //selCode = glueCode.EditValue.ToString();
                //txeID.Text = arrItem[0];
                txeDescription.Text = arrItem[3];
                txeComposition.Text = arrItem[4];
                txeWeight.Text = arrItem[5];
                txeModelNo.Text = arrItem[6];
                txeModelName.Text = arrItem[7];
                slueCategory.EditValue = arrItem[8];
                slueStyle.EditValue = arrItem[9];
                slueColor.EditValue = arrItem[10];
                slueSize.EditValue = arrItem[11];
                slueCustomer.EditValue = arrItem[12];
                txeBusinessUnit.Text = arrItem[13];
                cbeSeason.Text = arrItem[14];
                cbeClass.Text = arrItem[15];
                glueBranch.EditValue = arrItem[16];
                txeCostSheet.Text = arrItem[17];
                txeStdPrice.Text = arrItem[18];

                slueFirstVendor.EditValue = arrItem[19];
                //txeMatDetails.Text = "";
                //txeMatCode.Text = "";
                //txeSMPLLotNo.Text = "";
                //txePrice.Text = "";
                //txeCurrency.Text = "";
                if (arrItem[20] == "")
                    rgPurchase.EditValue = 9;
                else
                    rgPurchase.EditValue = Convert.ToInt32(arrItem[20]);

                if (arrItem[21] == "")
                    rgTax.EditValue = 9;
                else
                    rgTax.EditValue = Convert.ToInt32(arrItem[21]);


                txePurchaseLoss.Text = arrItem[22];
                if (arrItem[23] == "")
                    dteFirstReceiptDate.EditValue = null;
                else
                    dteFirstReceiptDate.EditValue = Convert.ToDateTime(arrItem[23]);

                slueDefaultVendor.EditValue = arrItem[24];

                txeSMPLNo.Text = "";
                dteRequestDate.EditValue = DateTime.Now;
                txeSMPLItem.Text = "";
                txeSMPLPatternNo.Text = "";
                rgZone.EditValue = 0;

                txeMinStock.Text = arrItem[25];
                txeMaxStock.Text = arrItem[26];
                txeStockSheifLife.Text = arrItem[27];
                txeStdCost.Text = arrItem[28];
                slueDefaultUnit.EditValue = arrItem[29];
                slueUnit.EditValue = "";

                if (arrItem[30] != "")
                {
                    txePath.Text = imgPathFile + arrItem[30];
                    picImg.Image = Image.FromFile(txePath.Text);
                }
                else
                {
                    txePath.Text = "";
                    picImg.Image = null;
                }

                txeLabTestNo.Text = arrItem[31];

                if (arrItem[32] == "")
                    dteApprovedLabDate.EditValue = null;
                else
                    dteApprovedLabDate.EditValue = Convert.ToDateTime(arrItem[32]);

                txeQCInspection.Text = arrItem[33];
                //clbQC.Items.Clear();

                slueVendorCode.EditValue = "";
                txeVendorName.Text = "";
                txeLotSize.Text = "";
                txeProductionLead.Text = "";
                txeDeliveryLead.Text = "";
                txeArrivalLead.Text = "";
                txePOCancelPeriod.Text = "";

                txeLots1.Text = "";
                txeLots2.Text = "";
                txeLots3.Text = "";

                txeRemark.Text = "";
                lblIDVENDItem.Text = "";

                glueCREATE.EditValue = arrItem[34];
                txeCDATE.Text = arrItem[35];
                glueUPDATE.EditValue = arrItem[36];
                txeUDATE.Text = arrItem[37];

                slueUnit.EditValue = arrItem[38];
                txeFactor.Text = arrItem[39];

                //**************************************
                if (slueFirstVendor.Text != "")
                {
                    slueFirstVendor.ReadOnly = true;
                    txeMatDetails.ReadOnly = true;
                    txeMatCode.ReadOnly = true;
                    txeSMPLLotNo.ReadOnly = true;
                    txePrice.ReadOnly = true;
                    slueCurrency.ReadOnly = true;
                    rgPurchase.ReadOnly = true;
                    rgTax.ReadOnly = true;
                    txePurchaseLoss.ReadOnly = true;
                    dteFirstReceiptDate.ReadOnly = true;
                }
                //**************************************
            }
            else
            {
                //txeID.Text = this.DBC.DBQuery("SELECT CASE WHEN ISNULL(MAX(OIDCUST), '') = '' THEN 1 ELSE MAX(OIDCUST) + 1 END AS NewNo FROM Customer").getString();

                //bool chkNameDup = chkDuplicateName();
                //if (chkNameDup == false)
                //{
                //    txeDescription.Text = "";
                //}

                txeDescription.Focus();
            }

            //selCode = "";
        

            sbSQL.Clear();
            sbSQL.Append("SELECT OIDCONDQC ");
            sbSQL.Append("FROM ItemInspection ");
            sbSQL.Append("WHERE (OIDITEM = '" + txeID.Text.Trim() + "') ");
            sbSQL.Append("ORDER BY OIDCONDQC ");
            DataTable dtQC = this.DBC.DBQuery(sbSQL).getDataTable();

            foreach (DataRow row in dtQC.Rows)
            {
                for (int i = 0; i < clbQC.ItemCount; i++)
                {
                    if (row["OIDCONDQC"].ToString() == clbQC.GetItemValue(i).ToString())
                    {
                        clbQC.SetItemCheckState(i, CheckState.Checked);
                        break;
                    }
                }
            }

            LoadVendorData();

            txeDescription.Focus();
        }

        private void glueCode_EditValueChanged(object sender, EventArgs e)
        {
            if (glueCode.Text.Trim() != "")
            {
                LoadCode(glueCode.EditValue.ToString());
                txeDescription.Focus();
            }
        }

        private void glueCode_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeDescription.Focus();
            }
        }

        private void glueCode_LostFocus(object sender, EventArgs e)
        {
           

        }

        private void slueCategory_EditValueChanged(object sender, EventArgs e)
        {
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT StyleName, OIDSTYLE AS ID ");
            sbSQL.Append("FROM  ProductStyle ");
            sbSQL.Append("WHERE (OIDGCATEGORY = '" + slueCategory.EditValue.ToString() + "') ");
            sbSQL.Append("ORDER BY StyleName ");
            new ObjDE.setSearchLookUpEdit(slueStyle, sbSQL, "StyleName", "ID").getData(true);
            slueStyle.Focus();
        }

        private void slueFirstVendor_EditValueChanged(object sender, EventArgs e)
        {
            slueDefaultVendor.EditValue = slueFirstVendor.EditValue.ToString();

            txeMatDetails.Text = "";
            txeMatCode.Text = "";
            txeSMPLLotNo.Text = "";
            txePrice.Text = "";
            slueCurrency.EditValue = "";

            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT TOP (1) MatDetails, MatCode, SMPLLotNo, FORMAT(Price, '###0.####') AS Price, Currency ");
            sbSQL.Append("FROM   ItemVendor ");
            sbSQL.Append("WHERE(OIDITEM = '" + txeID.Text.Trim() + "') AND(OIDVEND = '" + slueFirstVendor.EditValue.ToString() + "') ");
            string[] arrVEND = this.DBC.DBQuery(sbSQL).getMultipleValue();
            if (arrVEND.Length > 0)
            {
                txeMatDetails.Text = arrVEND[0];
                txeMatCode.Text = arrVEND[1];
                txeSMPLLotNo.Text = arrVEND[2];
                txePrice.Text = arrVEND[3];
                slueCurrency.EditValue = arrVEND[4];
            }

            txeMatDetails.Focus();
        }

        private void slueDefaultVendor_EditValueChanged(object sender, EventArgs e)
        {
            slueVendorCode.EditValue = slueDefaultVendor.EditValue.ToString();
            txeMinStock.Focus();
        }

        private void btnSelect_Click(object sender, EventArgs e)
        {
            xtraOpenFileDialog1.Filter = "Image Files|*.jpg;*.jpeg;*.gif,*.png";
            xtraOpenFileDialog1.FileName = "";
            xtraOpenFileDialog1.Title = "Select Image File";

            if (xtraOpenFileDialog1.ShowDialog() == DialogResult.OK)
            {
                txePath.Text = xtraOpenFileDialog1.FileName;
                picImg.Image = Image.FromFile(txePath.Text);
                picImg.Properties.SizeMode = DevExpress.XtraEditors.Controls.PictureSizeMode.Zoom;
            }

            txeLabTestNo.Focus();
        }

        private void btnNEW_Click(object sender, EventArgs e)
        {
            slueVendorCode.EditValue = "";
            txeVendorName.Text = "";
            txeLotSize.Text = "";
            txeProductionLead.Text = "";
            txeDeliveryLead.Text = "";
            txeArrivalLead.Text = "";
            txePOCancelPeriod.Text = "";

            txeLots1.Text = "";
            txeLots2.Text = "";
            txeLots3.Text = "";

            txeRemark.Text = "";
            lblIDVENDItem.Text = "";
            selCode = "";

        }

        private void slueVendorCode_EditValueChanged(object sender, EventArgs e)
        {
            txeVendorName.Text = "";
            txeLotSize.Text = "";
            txeProductionLead.Text = "";
            txeDeliveryLead.Text = "";
            txeArrivalLead.Text = "";
            txePOCancelPeriod.Text = "";

            txeLots1.Text = "";
            txeLots2.Text = "";
            txeLots3.Text = "";

            txeRemark.Text = "";
            lblIDVENDItem.Text = "";
            //selCode = "";

            if (slueVendorCode.Text.Trim() != "")
            {
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT IVD.OIDVENDItem, VD.OIDVEND, VD.Code AS Vendor, VD.Name AS VendorName, IVD.OIDITEM, IVD.LotSize, IVD.ProductionLead, IVD.DeliveryLead,  ");
                sbSQL.Append("       IVD.ArrivalLead, IVD.POCancelPeriod, IVD.PurchaseLots1, IVD.PurchaseLots2, IVD.PurchaseLots3, IVD.Remark ");
                sbSQL.Append("FROM   Vendor AS VD LEFT OUTER JOIN ");
                sbSQL.Append("       ItemVendor AS IVD ON IVD.OIDVEND = VD.OIDVEND ");
                sbSQL.Append("WHERE(VD.OIDVEND = '" + slueVendorCode.EditValue.ToString() + "') ");
                string[] arrVendor = this.DBC.DBQuery(sbSQL).getMultipleValue();
                if (arrVendor.Length > 0)
                {
                    txeVendorName.Text = arrVendor[3];
                    txeLotSize.Text = arrVendor[5];
                    txeProductionLead.Text = arrVendor[6];
                    txeDeliveryLead.Text = arrVendor[7];
                    txeArrivalLead.Text = arrVendor[8];
                    txePOCancelPeriod.Text = arrVendor[9];

                    txeLots1.Text = arrVendor[10];
                    txeLots2.Text = arrVendor[11];
                    txeLots3.Text = arrVendor[12];

                    txeRemark.Text = arrVendor[13];
                    lblIDVENDItem.Text = arrVendor[0];
                }
            }
            //MessageBox.Show(slueVendorCode.EditValue.ToString());
            txeVendorName.Focus();
        }

        private void btnADD_Click(object sender, EventArgs e)
        {
            if (slueVendorCode.Text.Trim() == "")
            {
                FUNC.msgWarning("Please select vendor code.");
                slueVendorCode.Focus();
            }
            else
            {
                bool chkVendor = true;
                foreach (DataRow dr in dtVendor.Rows) // search whole table
                {
                    if (dr["OIDVEND"].ToString() == slueVendorCode.EditValue.ToString())
                    {
                        dr["LotSize"] = txeLotSize.Text.Trim();

                        dr["ProductionLead"] = txeProductionLead.Text.Trim();
                        dr["DeliveryLead"] = txeDeliveryLead.Text.Trim();
                        dr["ArrivalLead"] = txeArrivalLead.Text.Trim();
                        dr["POCancelPeriod"] = txePOCancelPeriod.Text.Trim();

                        dr["PurchaseLots1"] = txeLots1.Text.Trim();
                        dr["PurchaseLots2"] = txeLots2.Text.Trim();
                        dr["PurchaseLots3"] = txeLots3.Text.Trim();
                        dr["Remark"] = txeRemark.Text.Trim();

                        chkVendor = false;
                        break;
                    }
                }

                if (chkVendor == true)
                {
                    dtVendor.Rows.Add(new Object[]{
                    "",
                    slueVendorCode.EditValue.ToString(),
                    slueVendorCode.Text,
                    txeVendorName.Text,
                    txeID.Text.Trim(),
                    txeLotSize.Text.Trim(),
                    txeProductionLead.Text.Trim(),
                    txeDeliveryLead.Text.Trim(),
                    txeArrivalLead.Text.Trim(),
                    txePOCancelPeriod.Text.Trim(),
                    txeLots1.Text.Trim(),
                    txeLots2.Text.Trim(),
                    txeLots3.Text.Trim(),
                    txeRemark.Text.Trim()
                });
                }

                gcVendor.DataSource = dtVendor;
                gcVendor.EndUpdate();
                gcVendor.ResumeLayout();
                gvVendor.ClearSelection();
                gvVendor.OptionsView.ColumnAutoWidth = false;
                gvVendor.BestFitColumns();
                gvVendor.BestFitColumns();
                gvVendor.ClearSelection();

                btnNEW.PerformClick();
            }
        }

        private void gvVendor_RowCellClick(object sender, RowCellClickEventArgs e)
        {
            slueVendorCode.EditValue = gvVendor.GetFocusedRowCellValue("OIDVEND").ToString();
            txeVendorName.Text = gvVendor.GetFocusedRowCellValue("VendorName").ToString();
            txeLotSize.Text = gvVendor.GetFocusedRowCellValue("LotSize").ToString();
            txeProductionLead.Text = gvVendor.GetFocusedRowCellValue("ProductionLead").ToString();
            txeDeliveryLead.Text = gvVendor.GetFocusedRowCellValue("DeliveryLead").ToString();
            txeArrivalLead.Text = gvVendor.GetFocusedRowCellValue("ArrivalLead").ToString();
            txePOCancelPeriod.Text = gvVendor.GetFocusedRowCellValue("POCancelPeriod").ToString();

            txeLots1.Text = gvVendor.GetFocusedRowCellValue("PurchaseLots1").ToString();
            txeLots2.Text = gvVendor.GetFocusedRowCellValue("PurchaseLots2").ToString();
            txeLots3.Text = gvVendor.GetFocusedRowCellValue("PurchaseLots3").ToString();

            txeRemark.Text = gvVendor.GetFocusedRowCellValue("Remark").ToString();
            lblIDVENDItem.Text = gvVendor.GetFocusedRowCellValue("OIDVENDItem").ToString();
        }

        private void bbiExcel_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            string pathFile = new ObjSet.Folder(@"C:\MDS\Export\").GetPath() + "ItemsList_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
            gvListItem.ExportToXlsx(pathFile);
            System.Diagnostics.Process.Start(pathFile);
        }

        private void picImg_Click(object sender, EventArgs e)
        {
            if (txePath.Text.Trim() != "")
            {
                ShowImage frmIMG = new ShowImage(txePath.Text.Trim());
                frmIMG.Show();
            }
        }

        private void glueCode_Closed(object sender, ClosedEventArgs e)
        {
            //glueCode.Focus();
            //txeDescription.Focus();
        }

        private void glueCode_ProcessNewValue(object sender, ProcessNewValueEventArgs e)
        {
            //GridLookUpEdit gridLookup = sender as GridLookUpEdit;
            //if (e.DisplayValue == null) return;
            //string newValue = e.DisplayValue.ToString();
            //if (newValue == String.Empty) return;
        }

        private void glueCode_Leave(object sender, EventArgs e)
        {
            //MessageBox.Show("Code:'" + glueCode.Text.ToUpper().Trim() + "', Old:'" + txeOldCode.Text.ToUpper().Trim() + "'");
            //if (glueCode.Text.ToUpper().Trim() == "")
            //{
            //    glueCode.Text = txeOldCode.Text;
            //    selCode = glueCode.Text;
            //}
            //else if (glueCode.Text.ToUpper().Trim() != txeOldCode.Text.ToUpper().Trim())
            //{
            //    if (glueCode.Text.ToUpper().Trim() != selCode)
            //    {
            //        if (lblStatus.Text == "NEW")
            //        {
            //            //if (glueCode.Text.Trim() != "" && glueCode.Text.ToUpper().Trim() != selCode)
            //            //{
            //            glueCode.Text = glueCode.Text.ToUpper().Trim();
            //            selCode = glueCode.Text;
            //            txeOldCode.Text = "";
            //            LoadCode(glueCode.Text);
            //            //}
            //        }
            //        else if (lblStatus.Text == "EDIT")
            //        {
            //            MessageBox.Show(glueCode.Text);
            //            //glueCode.Text = glueCode.Text.ToUpper().Trim();
            //            if (FUNC.msgQuiz("Do you want to chang code from '" + txeOldCode.Text + "' to '" + glueCode.Text.ToUpper().Trim() + "' ?\nคุณต้องการเปลี่ยนรหัสจาก '" + txeOldCode.Text.ToUpper().Trim() + "' เป็น '" + glueCode.Text + "' ใช่หรือไม่") == true)
            //            {
            //                bool chkPass = true;
            //                if (glueCode.Text.ToUpper().Trim().Length >= 5)
            //                {
            //                    if (glueCode.Text.ToUpper().Trim().Substring(0, 5) == "TMPFB" || glueCode.Text.Substring(0, 5) == "TMPMT")
            //                    {
            //                        FUNC.msgWarning("Cannot set code starting with 'TMPFB' or 'TMPMT'. Please change code.\nไม่สามารถตั้งรหัสที่ขึ้นต้นด้วย 'TMPFB' หรือ 'TMPMT' ได้");
            //                        glueCode.Text = txeOldCode.Text;
            //                        selCode = glueCode.Text;
            //                        //txeDescription.Focus();
            //                        chkPass = false;
            //                    }
            //                }

            //                if (chkPass == true)
            //                {
            //                    StringBuilder sbSQL = new StringBuilder();
            //                    sbSQL.Append("SELECT  TOP (1) OIDITEM FROM Items WHERE (Code = N'" + glueCode.Text.ToUpper().Trim() + "')");
            //                    string chkCode = this.DBC.DBQuery(sbSQL).getString();
            //                    if (chkCode != "")
            //                    {
            //                        FUNC.msgError("Duplicate Code !! Please change.\nรหัสซ้ำ กรุณาเปลี่ยน");
            //                        glueCode.Text = txeOldCode.Text;
            //                        selCode = glueCode.Text;
            //                        //txeDescription.Focus();
            //                    }
            //                }
            //            }
            //            else
            //            {
            //                glueCode.Text = txeOldCode.Text;
            //                selCode = glueCode.Text;
            //                //txeDescription.Focus();
            //            }
            //        }
            //    }
            //}

        }

        private void txeDescription_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeComposition.Focus();
            }
        }

        private void txeDescription_Leave(object sender, EventArgs e)
        {

        }

        private void txeComposition_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeWeight.Focus();
            }
        }

        private void txeWeight_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeModelNo.Focus();
            }
        }

        private void txeModelNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeModelName.Focus();
            }
        }

        private void txeModelName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                slueCategory.Focus();
            }
        }

        private void slueStyle_EditValueChanged(object sender, EventArgs e)
        {
            slueColor.Focus();
        }

        private void slueColor_EditValueChanged(object sender, EventArgs e)
        {
            slueSize.Focus();
        }

        private void slueSize_EditValueChanged(object sender, EventArgs e)
        {
            slueCustomer.Focus();
        }

        private void slueCustomer_EditValueChanged(object sender, EventArgs e)
        {
            txeBusinessUnit.Focus();
        }

        private void txeBusinessUnit_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                cbeSeason.Focus();
            }
        }

        private void cbeSeason_SelectedIndexChanged(object sender, EventArgs e)
        {
            //cbeClass.Focus();
        }

        private void cbeClass_SelectedIndexChanged(object sender, EventArgs e)
        {
            //glueBranch.Focus();
        }

        private void glueBranch_EditValueChanged(object sender, EventArgs e)
        {
            txeCostSheet.Focus();
        }

        private void txeCostSheet_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeStdPrice.Focus();
            }
        }

        private void txeStdPrice_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                slueFirstVendor.Focus();
            }
        }

        private void txeMatDetails_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeMatCode.Focus();
            }
        }

        private void txeMatCode_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeSMPLLotNo.Focus();
            }
        }

        private void txeSMPLLotNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txePrice.Focus();
            }
        }

        private void txePrice_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                slueCurrency.Focus();
            }
        }

        private void txeCurrency_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                rgPurchase.Focus();
            }
        }

        private void txePurchaseLoss_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                dteFirstReceiptDate.Focus();
            }
        }

        private void dteFirstReceiptDate_EditValueChanged(object sender, EventArgs e)
        {
            slueDefaultVendor.Focus();
        }

        private void txeMinStock_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeMaxStock.Focus();
            }
        }

        private void txeMaxStock_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeStockSheifLife.Focus();
            }
        }

        private void txeStockSheifLife_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode==Keys.Enter)
            {
                txeStdCost.Focus();
            }
        }

        private void txeStdCost_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                slueDefaultUnit.Focus();
            }
        }

        private void slueDefaultUnit_EditValueChanged(object sender, EventArgs e)
        {
            slueUnit.Focus();
        }



        private void txePath_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeLabTestNo.Focus();
            }
        }

        private void txeLabTestNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                dteApprovedLabDate.Focus();
            }
        }

        private void dteApprovedLabDate_EditValueChanged(object sender, EventArgs e)
        {
            txeQCInspection.Focus();
        }

        private void txeQCInspection_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                clbQC.Focus();
            }
        }

        private void txeSMPLNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                dteRequestDate.Focus();
            }
        }

        private void dteRequestDate_EditValueChanged(object sender, EventArgs e)
        {
            txeSMPLItem.Focus();
        }

        private void txeSMPLItem_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeSMPLPatternNo.Focus();
            }
        }

        private void txeSMPLPatternNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                rgZone.Focus();
            }
        }

        private void txeVendorName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeLotSize.Focus();
            }
        }

        private void txeProductionLead_EditValueChanged(object sender, EventArgs e)
        {

        }

        private void txeLotSize_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeProductionLead.Focus();
            }
        }

        private void txeProductionLead_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeDeliveryLead.Focus();
            }
        }

        private void txeDeliveryLead_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeArrivalLead.Focus();
            }
        }

        private void txeArrivalLead_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txePOCancelPeriod.Focus();
            }
        }

        private void txePOCancelPeriod_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeLots1.Focus();
            }
        }

        private void txeLots1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeLots2.Focus();
            }
        }

        private void txeLots2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeLots3.Focus();
            }
        }

        private void txeLots3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeRemark.Focus();
            }
        }

        private void glueSMaterial_EditValueChanged(object sender, EventArgs e)
        {
            LoadItemCode();
            LOAD_LIST();
        }

        private void glueSBranch_EditValueChanged(object sender, EventArgs e)
        {
            LOAD_LIST();
        }

        private void slueSCustomer_EditValueChanged(object sender, EventArgs e)
        {
            LOAD_LIST();
        }

        private void slueSVendor_EditValueChanged(object sender, EventArgs e)
        {
            LOAD_LIST();
        }

        private void glueSCategory_EditValueChanged(object sender, EventArgs e)
        {
            LoadStyle();
            LOAD_LIST();
        }

        private void slueSStyle_EditValueChanged(object sender, EventArgs e)
        {
            LOAD_LIST();
        }

        private void slueSCode_EditValueChanged(object sender, EventArgs e)
        {
            LOAD_LIST();
        }

        private void slueSModel_EditValueChanged(object sender, EventArgs e)
        {
            LOAD_LIST();
        }

        private void M07_Shown(object sender, EventArgs e)
        {
            tabbedControlGroup1.SelectedTabPage = layoutControlGroup13; //เลือกแท็บ Search
            bbiRefresh.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
            bbiPrintPreview.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
            bbiPrint.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
            bbiExcel.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
            // layoutControlGroup4.Text = "";

        }

        private void gvListItem_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator) e.Info.DisplayText = (e.RowHandle + 1).ToString();
        }

        private void bbiRefresh_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            LOAD_LIST();
        }

        private void tabbedControlGroup1_SelectedPageChanged(object sender, DevExpress.XtraLayout.LayoutTabPageChangedEventArgs e)
        {
            if (tabbedControlGroup1.SelectedTabPage == layoutControlGroup13)
            {
                bbiNew.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
                bbiEdit.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
                bbiSave.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
                bbiDelete.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;

                bbiRefresh.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
                bbiPrintPreview.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
                bbiPrint.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
                bbiExcel.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
            }
            else
            {
                bbiNew.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
                bbiEdit.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
                bbiSave.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
                bbiDelete.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;

                bbiRefresh.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
                bbiPrintPreview.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
                bbiPrint.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
                bbiExcel.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
            }
        }

        private void bbiPrintPreview_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            gcListItem.ShowPrintPreview();
        }

        private void bbiPrint_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            gcListItem.Print();
        }

        private void gvListItem_DoubleClick(object sender, EventArgs e)
        {
            DXMouseEventArgs ea = e as DXMouseEventArgs;
            GridView view = sender as GridView;
            GridHitInfo info = view.CalcHitInfo(ea.Location);
            if (info.InRow || info.InRowCell)
            {
                txeID.Focus();
                GridView gv = gvListItem;
                string Code = gv.GetFocusedRowCellValue("OIDITEM").ToString();
                selCode = Code;
                string MaterialType = gv.GetFocusedRowCellValue("MaterialTypeID").ToString();
                //OIDITEM
                if (Code != "")
                {
                    if (Code.Length >= 5)
                        if (Code.Substring(0, 5).ToUpper().Trim() == "TMPFB")
                            rgMaterial.EditValue = 1;
                        else if (Code.Substring(0, 5).ToUpper().Trim() == "TMPMT")
                            rgMaterial.EditValue = 2;
                        else
                            rgMaterial.EditValue = Convert.ToInt32(MaterialType);
                    else
                        rgMaterial.EditValue = Convert.ToInt32(MaterialType);

                    txeOldCode.EditValue = Code;
                    glueCode.EditValue = Code;
                    LoadCode(Code, false);

                    tabbedControlGroup1.SelectedTabPage = layoutControlGroup1;
                    //txeID.Focus();
                }

            }
        }

        private void gvListItem_DragObjectDrop(object sender, DevExpress.XtraGrid.Views.Base.DragObjectDropEventArgs e)
        {

        }

        private void glueCode_Validated(object sender, EventArgs e)
        {
            //MessageBox.Show("Code:'" + glueCode.Text.ToUpper().Trim() + "', sel:'" + selCode + "'");
            glueCode.Text = glueCode.Text.ToUpper().Trim();
            if (glueCode.Text.ToUpper().Trim() == "")
            {
                glueCode.EditValue = txeOldCode.Text;
                selCode = glueCode.EditValue.ToString();
                if (txeOldCode.Text == "")
                    txeID.Text = "";
            }
            else if (glueCode.EditValue.ToString() != txeOldCode.Text)
            {
                if (glueCode.EditValue.ToString() != selCode)
                {
                    if (lblStatus.Text == "NEW")
                    {
                        //if (glueCode.Text.Trim() != "" && glueCode.Text.ToUpper().Trim() != selCode)
                        //{
                        //glueCode.Text = glueCode.Text.ToUpper().Trim();
                        selCode = glueCode.EditValue.ToString();
                        txeOldCode.Text = "";
                        LoadCode(glueCode.EditValue.ToString());
                        //}
                    }
                    else if (lblStatus.Text == "EDIT")
                    {
                        //MessageBox.Show(glueCode.Text);
                        //glueCode.Text = glueCode.Text.ToUpper().Trim();
                        //if (FUNC.msgQuiz("Do you want to chang code ?\nคุณต้องการเปลี่ยนรหัสจาก '" + txeOldCode.Text.ToUpper().Trim() + "' เป็น '" + glueCode.Text + "' ใช่หรือไม่") == true)
                        //{
                            selCode = glueCode.EditValue.ToString();
                            bool chkPass = true;
                            if (glueCode.Text.ToUpper().Trim().Length >= 5)
                            {
                                if (glueCode.Text.ToUpper().Trim().Substring(0, 5) == "TMPFB" || glueCode.Text.Substring(0, 5) == "TMPMT")
                                {
                                    FUNC.msgWarning("Cannot set code starting with 'TMPFB' or 'TMPMT'. Please change code.\nไม่สามารถตั้งรหัสที่ขึ้นต้นด้วย 'TMPFB' หรือ 'TMPMT' ได้");
                                    glueCode.EditValue = txeOldCode.Text;
                                    selCode = glueCode.EditValue.ToString();
                                    //txeDescription.Focus();
                                    chkPass = false;
                                }
                            }

                            if (chkPass == true)
                            {
                                StringBuilder sbSQL = new StringBuilder();
                                sbSQL.Append("SELECT  TOP (1) OIDITEM FROM Items WHERE (OIDITEM = '" + glueCode.EditValue.ToString() + "')");
                                string chkCode = this.DBC.DBQuery(sbSQL).getString();
                                if (chkCode != "")
                                {
                                    FUNC.msgError("Duplicate Code !! Please change.\nรหัสซ้ำ กรุณาเปลี่ยน");
                                    glueCode.EditValue = txeOldCode.Text;
                                    selCode = glueCode.EditValue.ToString();
                                    //txeDescription.Focus();
                                }
                            }
                        //}
                        //else
                        //{
                        //    glueCode.Text = txeOldCode.Text;
                        //    selCode = glueCode.Text;
                        //    if (txeOldCode.Text == "")
                        //        txeID.Text = "";
                        //    //txeDescription.Focus();
                        //}
                    }
                }
            }
        }

        private void btnDELETE_Click(object sender, EventArgs e)
        {

        }

        private void gvVendor_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator) e.Info.DisplayText = (e.RowHandle + 1).ToString();
            gvVendor.IndicatorWidth = 40;
        }

        private void sbNew_Click(object sender, EventArgs e)
        {
            if (lblStatus.Text == "NEW")
            {
                var frm = new M07_01(this.DBC, UserLogin.OIDUser, rgMaterial.EditValue.ToString(), rgMaterial.Properties.Items[rgMaterial.SelectedIndex].Description, "");
                frm.ShowDialog(this);
            }
            else if (lblStatus.Text == "EDIT")
            {
                var frm = new M07_01(this.DBC, UserLogin.OIDUser, rgMaterial.EditValue.ToString(), rgMaterial.Properties.Items[rgMaterial.SelectedIndex].Description, txeID.Text.Trim());
                frm.ShowDialog(this);
            }
        }

        private void ribbonControl2_Click(object sender, EventArgs e)
        {

        }

        private void sbCategory_Click(object sender, EventArgs e)
        {
            var frm = new M07_02(this.DBC, UserLogin.OIDUser);
            frm.ShowDialog(this);
        }

        private void sbStyle_Click(object sender, EventArgs e)
        {
            if (sbCategory.Text == "")
            {
                FUNC.msgWarning("Please select category (division).");
                sbCategory.Focus();
            }
            else
            {
                var frm = new M07_03(this.DBC, slueCategory.EditValue.ToString(), UserLogin.OIDUser);
                frm.ShowDialog(this);
            }
        }

        private void sbVENDOR_Click(object sender, EventArgs e)
        {
            var frm = new M07_M12(this.DBC, "FG", UserLogin.OIDUser);
            frm.ShowDialog(this);
        }

        private void sbBrowse_Click(object sender, EventArgs e)
        {
            cbeSheet.Properties.Items.Clear();
            cbeSheet.Text = "";

            xtraOpenFileDialog1.Filter = "Excel files |*.xlsx;*.xls;*.csv";
            xtraOpenFileDialog1.FileName = "";
            xtraOpenFileDialog1.Title = "Select Excel File";

            if (xtraOpenFileDialog1.ShowDialog() == DialogResult.OK)
            {
                txeFilePath.Text = xtraOpenFileDialog1.FileName;
                DevExpress.XtraSpreadsheet.SpreadsheetControl xss = new DevExpress.XtraSpreadsheet.SpreadsheetControl();
                IWorkbook workbook = xss.Document;
                using (FileStream stream = new FileStream(txeFilePath.Text, FileMode.Open))
                {
                    string ext = Path.GetExtension(txeFilePath.Text);
                    if (ext == ".xlsx")
                        workbook.LoadDocument(stream, DocumentFormat.Xlsx);
                    else if (ext == ".xls")
                        workbook.LoadDocument(stream, DocumentFormat.Xls);
                    else if (ext == ".csv")
                        workbook.LoadDocument(stream, DocumentFormat.Csv);
                }
                WorksheetCollection worksheets = workbook.Worksheets;
                for (int i = 0; i < worksheets.Count; i++)
                    cbeSheet.Properties.Items.Add(worksheets[i].Name);
            }
        }

        private void cbeSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txeFilePath.Text.Trim() != "" && cbeSheet.Text.Trim() != "")
            {
                IWorkbook workbook = spsImport.Document;

                try
                {
                    using (FileStream stream = new FileStream(txeFilePath.Text, FileMode.Open))
                    {
                        // workbook.CalculateFull();
                        string ext = Path.GetExtension(txeFilePath.Text);
                        if (ext == ".xlsx")
                            workbook.LoadDocument(stream, DocumentFormat.Xlsx);
                        else if (ext == ".xls")
                            workbook.LoadDocument(stream, DocumentFormat.Xls);
                        else if (ext == ".csv")
                            workbook.LoadDocument(stream, DocumentFormat.Csv);
                        //workbook.Worksheets.ActiveWorksheet = workbook.Worksheets[0];

                        //***Delete sheet
                        if (workbook.Worksheets.Count > 0)
                        {
                            for (int i = workbook.Worksheets.Count - 1; i >= 0; i--)
                            {
                                if (workbook.Worksheets[i].Name != cbeSheet.Text)
                                    workbook.Worksheets.RemoveAt(i);
                            }
                        }
                    }
                }
                catch (Exception)
                {
                    FUNC.msgWarning("Please close excel file before import.");
                    txeFilePath.Text = "";
                }
            }
        }

        private void slueUnit_EditValueChanged(object sender, EventArgs e)
        {
            txeFactor.Focus();
        }

        private void txeFactor_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btnSelect.Focus();
            }
        }
    }
}