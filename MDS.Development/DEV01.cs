using System;
using System.Windows.Forms;
using DevExpress.LookAndFeel;
using DevExpress.XtraGrid.Views.Grid;
using System.Data.SqlClient;
using DevExpress.Utils;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using System.Drawing;
using System.IO;
using System.Data;
using DevExpress.XtraEditors;
using DevExpress.XtraGrid;
using System.Collections;
using DevExpress.XtraPrinting;
using System.Diagnostics;
using DBConnect;
using System.Text;
using TheepClass;
using DevExpress.Spreadsheet;

namespace MDS.Development
{
    
    public partial class DEV01 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private Functionality.Function FUNCT = new Functionality.Function();

        private const string COMPANY_CODE = "PTC";

        const int TYPE_FG = 0;
        const int TYPE_FABRIC = 1;
        const int TYPE_ACCESSORY = 2;
        const int TYPE_PACKAGING = 3;
        const int TYPE_SAMPLE = 4;
        const int TYPE_OTHER = 9;
        const int TYPE_TEMPORARY = 8;

        string sql          = string.Empty;
        string imgPath      = Configuration.CONFIG.PATH_FILE + @"Pictures\";
        string reportPath = Configuration.CONFIG.PATH_FILE + @"Report\";

        string currenTab    = string.Empty;
        string dosetOIDSMPL = string.Empty;
        bool PageFBVal      = false;

        string status_Mat = string.Empty;

        DataTable dtQtyRequired = new DataTable();
        DataTable dtFBSample = new DataTable();
        DataTable dtMaterial = new DataTable();
        DataTable dtFBSize = new DataTable();
        DataTable dtFabric = new DataTable();
        DataTable dtMTList = new DataTable();
        DataTable dtCSConsumption = new DataTable();

        string cloneSMPLNo = "";

        string BF_Color = "";
        string BF_Size = "";
        string BF_Qty = "";

        string New_Color = "";
        string New_Size = "";
        string New_Qty = "";

        string tmpColor = "";
        string tmpSize = "";
        string tmpQuantity = "";

        int chkReadWrite = 0;

        public LogIn UserLogin { get; set; }
        public int Company { get; set; }
        public string ConnectionString { get; set; }

        string CONNECT_STRING = "";

        DatabaseConnect DBC;
        public DEV01()
        {
            InitializeComponent();
            UserLookAndFeel.Default.StyleChanged += MyStyleChanged;
        }

        public DEV01(string SMPLNo, LogIn _UserLogin)
        {
            InitializeComponent();
            UserLookAndFeel.Default.StyleChanged += MyStyleChanged;
            if (SMPLNo != "")
            {
                this.cloneSMPLNo = SMPLNo.ToUpper().Trim();
                this.UserLogin = _UserLogin;
            }
        }

        private void MyStyleChanged(object sender, EventArgs e)
        {
            UserLookAndFeel userLookAndFeel = (UserLookAndFeel)sender;
            cUtility.SaveRegistry(@"Software\MDS", "SkinName", userLookAndFeel.SkinName);
            cUtility.SaveRegistry(@"Software\MDS", "SkinPalette", userLookAndFeel.ActiveSvgPaletteName);
        }

        private void XtraForm1_Load(object sender, EventArgs ex)
        {
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
            //******************************

            radioGroup4.EditValue = 0;

            lblUser.Text = "Login : " + UserLogin.FullName;
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT TOP (1) ReadWriteStatus FROM FunctionAccess WHERE (OIDUser = '" + UserLogin.OIDUser + "') AND(FunctionNo = 'DEV01') ");
            chkReadWrite = this.DBC.DBQuery(sbSQL).getInt();

            //MessageBox.Show(chkReadWrite.ToString());
            if (chkReadWrite == 0)
            {
                ribbonPageGroup1.Visible = false;
                rpgManage.Visible = false;

                layoutControlItem29.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                simpleButton2.Enabled = false;
                simpleButton3.Enabled = false;
                simpleButton4.Enabled = false;
                sbColor.Enabled = false;
                sbSize.Enabled = false;
                sbFBColor.Enabled = false;
                sbTempCode.Enabled = false;
                sbMTColor.Enabled = false;
                sbTempCodeMat.Enabled = false;
                btnOpenImg_Main.Enabled = false;
                sbDelete_S.Enabled = false;
                sbClear.Enabled = false;
                simpleButton5.Enabled = false;
                sbDelete_F.Enabled = false;
                sbMatClear.Enabled = false;
                btnUploadMat.Enabled = false;
                simpleButton1.Enabled = false;

                sbUseFor.Enabled = false;
                sbUnit.Enabled = false;

                sbPart.Enabled = false;
                sbFBSupplier.Enabled = false;

                sbMTSupplier.Enabled = false;
            }

            sbSQL.Clear();
            sbSQL.Append("SELECT FullName, OIDUSER FROM Users ORDER BY OIDUSER ");
            new ObjDE.setGridLookUpEdit(glueCreateBy_Main, sbSQL, "FullName", "OIDUSER").getData();
            new ObjDE.setGridLookUpEdit(glueUpdateBy_Main, sbSQL, "FullName", "OIDUSER").getData();
            //new ObjDE.setGridLookUpEdit(glueCreateBy_Main, sbSQL, "FullName", "OIDUSER").getData();
            //new ObjDE.setGridLookUpEdit(glueUpdateBy_Main, sbSQL, "FullName", "OIDUSER").getData();

            glueCreateBy_Main.EditValue = UserLogin.OIDUser;
            glueUpdateBy_Main.EditValue = UserLogin.OIDUser;

            dtQtyRequired = new DataTable();
            dtQtyRequired.Columns.Add("Color", typeof(String));
            dtQtyRequired.Columns.Add("Size", typeof(String));
            dtQtyRequired.Columns.Add("Quantity", typeof(String));
            dtQtyRequired.Columns.Add("ID", typeof(String));
            dtQtyRequired.Columns.Add("Delete", typeof(String));

            dtFBSample = new DataTable();
            dtFBSample.Columns.Add("OIDSMPL", typeof(String));
            dtFBSample.Columns.Add("SMPLPatternNo", typeof(String));
            dtFBSample.Columns.Add("ColorName", typeof(String));
            dtFBSample.Columns.Add("OIDSMPLDT", typeof(String));

            dtMaterial = new DataTable();
            dtMaterial.Columns.Add("OIDSMPL", typeof(String));
            dtMaterial.Columns.Add("SMPLPatternNo", typeof(String));
            dtMaterial.Columns.Add("ColorName", typeof(String));
            dtMaterial.Columns.Add("SizeName", typeof(String));
            dtMaterial.Columns.Add("Quantity", typeof(String));
            dtMaterial.Columns.Add("UnitName", typeof(String));
            dtMaterial.Columns.Add("OIDSMPLDT", typeof(String));

            dtFBSize = new DataTable();
            dtFBSize.Columns.Add("SizeName", typeof(String));
            dtFBSize.Columns.Add("Quantity", typeof(String));
            dtFBSize.Columns.Add("OIDSMPLDT", typeof(String));

            dtFabric = new DataTable();
            dtFabric.Columns.Add("ColorName", typeof(string));
            dtFabric.Columns.Add("VendorFBCode", typeof(string));
            dtFabric.Columns.Add("SMPLotNo", typeof(string));
            dtFabric.Columns.Add("Supplier", typeof(string));
            dtFabric.Columns.Add("FabricColor", typeof(string));
            dtFabric.Columns.Add("FabricCode", typeof(string));
            dtFabric.Columns.Add("Description", typeof(string));
            dtFabric.Columns.Add("Composition", typeof(string));
            dtFabric.Columns.Add("FBWeight", typeof(string));
            dtFabric.Columns.Add("WidthCut", typeof(string));
            dtFabric.Columns.Add("Price", typeof(string));
            dtFabric.Columns.Add("Currency", typeof(string));
            dtFabric.Columns.Add("TTWidth", typeof(string));
            dtFabric.Columns.Add("UsableWidth", typeof(string));
            dtFabric.Columns.Add("PicFile", typeof(string));
            dtFabric.Columns.Add("ChkFBCode", typeof(string));
            dtFabric.Columns.Add("FBPartsID", typeof(string));
            dtFabric.Columns.Add("FBPartsName", typeof(string));
            dtFabric.Columns.Add("FBID", typeof(string));
            dtFabric.Columns.Add("Remark", typeof(string));

            dtMTList = new DataTable();
            dtMTList.Columns.Add("MatID", typeof(string));
            dtMTList.Columns.Add("SampleID", typeof(string));
            dtMTList.Columns.Add("WorkStation", typeof(string));
            dtMTList.Columns.Add("VendMTCode", typeof(string));
            dtMTList.Columns.Add("SMPLotNo", typeof(string));
            dtMTList.Columns.Add("Vendor", typeof(string));
            dtMTList.Columns.Add("MatColor", typeof(string));
            dtMTList.Columns.Add("ColorName", typeof(string));
            dtMTList.Columns.Add("MatSize", typeof(string));
            dtMTList.Columns.Add("Consumption", typeof(string));
            dtMTList.Columns.Add("Unit", typeof(string));
            dtMTList.Columns.Add("Composition", typeof(string));
            dtMTList.Columns.Add("Details", typeof(string));
            dtMTList.Columns.Add("Price", typeof(string));
            dtMTList.Columns.Add("Currency", typeof(string));
            dtMTList.Columns.Add("NAVCode", typeof(string));
            dtMTList.Columns.Add("Description", typeof(string));
            dtMTList.Columns.Add("Situation", typeof(string));
            dtMTList.Columns.Add("Comment", typeof(string));
            dtMTList.Columns.Add("Remark", typeof(string));
            dtMTList.Columns.Add("PathFile", typeof(string));
            dtMTList.Columns.Add("ChkMTCode", typeof(string));
            dtMTList.Columns.Add("MID", typeof(string));

            dtCSConsumption = new DataTable();
            dtCSConsumption.Columns.Add("Color", typeof(string));
            dtCSConsumption.Columns.Add("Size", typeof(string));
            dtCSConsumption.Columns.Add("Consumption", typeof(string));

            gridControl1.DataSource = null;

            gcQtyRequired.DataSource = dtQtyRequired;

            gridControl3.DataSource = dtFBSample;
            gcSize_Fabric.DataSource = dtFBSize;
            gcList_Fabric.DataSource = dtFabric;
            gridControl5.DataSource = null;

            gridControl6.DataSource = dtMaterial;
            gridControl7.DataSource = dtCSConsumption;
            gridControl8.DataSource = dtMTList;

            LoadDefaultData();
            LoadNewData();

        }

        private void NewFabric()
        {
            lblFBStatus.Text = "Status : New";
            lblFBStatus.BackColor = Color.Green;
        }

        internal void LoadSizeColor()
        {
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("Select ColorNo, ColorName, OIDCOLOR AS ID From ProductColor WHERE (ColorType = '0') ORDER BY ColorNo ");
            SearchLookUpEdit slueColor = new SearchLookUpEdit();
            new ObjDE.setSearchLookUpEdit(slueColor, sbSQL, "ColorName", "ID").getData();
            slueColor.Properties.View.PopulateColumns(slueColor.Properties.DataSource);
            slueColor.Properties.View.Columns["ID"].Visible = false;

            rep_slueColor.DataSource = slueColor.Properties.DataSource;
            rep_slueColor.DisplayMember = "ColorName";
            rep_slueColor.ValueMember = "ID";
            rep_slueColor.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
            rep_slueColor.View.PopulateColumns(rep_slueColor.DataSource);
            rep_slueColor.View.Columns["ID"].Visible = false;

            rep_slueColor_FB.DataSource = slueColor.Properties.DataSource;
            rep_slueColor_FB.DisplayMember = "ColorName";
            rep_slueColor_FB.ValueMember = "ID";
            rep_slueColor_FB.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
            rep_slueColor_FB.View.PopulateColumns(rep_slueColor_FB.DataSource);
            rep_slueColor_FB.View.Columns["ID"].Visible = false;

            rep_slueColorFB.DataSource = slueColor.Properties.DataSource;
            rep_slueColorFB.DisplayMember = "ColorName";
            rep_slueColorFB.ValueMember = "ID";
            rep_slueColorFB.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
            rep_slueColorFB.View.PopulateColumns(rep_slueColorFB.DataSource);
            rep_slueColorFB.View.Columns["ID"].Visible = false;

            rep_slColor_Mat.DataSource = slueColor.Properties.DataSource;
            rep_slColor_Mat.DisplayMember = "ColorName";
            rep_slColor_Mat.ValueMember = "ID";
            rep_slColor_Mat.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
            rep_slColor_Mat.View.PopulateColumns(rep_slColor_Mat.DataSource);
            rep_slColor_Mat.View.Columns["ID"].Visible = false;

            rep_slColor_Material.DataSource = slueColor.Properties.DataSource;
            rep_slColor_Material.DisplayMember = "ColorName";
            rep_slColor_Material.ValueMember = "ID";
            rep_slColor_Material.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
            rep_slColor_Material.View.PopulateColumns(rep_slColor_Material.DataSource);
            rep_slColor_Material.View.Columns["ID"].Visible = false;

            rep_MtrColorName.DataSource = slueColor.Properties.DataSource;
            rep_MtrColorName.DisplayMember = "ColorName";
            rep_MtrColorName.ValueMember = "ID";
            rep_MtrColorName.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
            rep_MtrColorName.View.PopulateColumns(rep_MtrColorName.DataSource);
            rep_MtrColorName.View.Columns["ID"].Visible = false;

            new ObjDE.setSearchLookUpEdit(slFGColor_FB, sbSQL, "ColorName", "ID").getData();
            slFGColor_FB.Properties.View.PopulateColumns(slFGColor_FB.Properties.DataSource);
            slFGColor_FB.Properties.View.Columns["ID"].Visible = false;


            sbSQL.Clear();
            sbSQL.Append("Select SizeNo, SizeName, OIDSIZE AS ID From ProductSize");
            SearchLookUpEdit slueSize = new SearchLookUpEdit();
            new ObjDE.setSearchLookUpEdit(slueSize, sbSQL, "SizeName", "ID").getData();
            slueSize.Properties.View.PopulateColumns(slueSize.Properties.DataSource);
            slueSize.Properties.View.Columns["ID"].Visible = false;

            rep_slueSize.DataSource = slueSize.Properties.DataSource;
            rep_slueSize.DisplayMember = "SizeName";
            rep_slueSize.ValueMember = "ID";
            rep_slueSize.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
            rep_slueSize.View.PopulateColumns(rep_slueSize.DataSource);
            rep_slueSize.View.Columns["ID"].Visible = false;

            rep_slueSize_FB.DataSource = slueSize.Properties.DataSource;
            rep_slueSize_FB.DisplayMember = "SizeName";
            rep_slueSize_FB.ValueMember = "ID";
            rep_slueSize_FB.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
            rep_slueSize_FB.View.PopulateColumns(rep_slueSize_FB.DataSource);
            rep_slueSize_FB.View.Columns["ID"].Visible = false;

            rep_slSize_Mat.DataSource = slueSize.Properties.DataSource;
            rep_slSize_Mat.DisplayMember = "SizeName";
            rep_slSize_Mat.ValueMember = "ID";
            rep_slSize_Mat.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
            rep_slSize_Mat.View.PopulateColumns(rep_slSize_Mat.DataSource);
            rep_slSize_Mat.View.Columns["ID"].Visible = false;

            rep_slSize_Material.DataSource = slueSize.Properties.DataSource;
            rep_slSize_Material.DisplayMember = "SizeName";
            rep_slSize_Material.ValueMember = "ID";
            rep_slSize_Material.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
            rep_slSize_Material.View.PopulateColumns(rep_slSize_Material.DataSource);
            rep_slSize_Material.View.Columns["ID"].Visible = false;

            rep_slSize.DataSource = slueSize.Properties.DataSource;
            rep_slSize.DisplayMember = "SizeName";
            rep_slSize.ValueMember = "ID";
            rep_slSize.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
            rep_slSize.View.PopulateColumns(rep_slSize.DataSource);
            rep_slSize.View.Columns["ID"].Visible = false;

            rep_MtrSizeName.DataSource = slueSize.Properties.DataSource;
            rep_MtrSizeName.DisplayMember = "SizeName";
            rep_MtrSizeName.ValueMember = "ID";
            rep_MtrSizeName.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
            rep_MtrSizeName.View.PopulateColumns(rep_MtrSizeName.DataSource);
            rep_MtrSizeName.View.Columns["ID"].Visible = false;

            //sbSQL.Clear();
            //sbSQL.Append("Select UnitName, OIDUNIT AS ID From Unit");
            //new ObjDE.setSearchLookUpEdit(slueUnit, sbSQL, "UnitName", "ID").getData();
            //slueConsumpUnit.Properties.DataSource = slueUnit.Properties.DataSource;
            //slueConsumpUnit.Properties.DisplayMember = "UnitName";
            //slueConsumpUnit.Properties.ValueMember = "ID";
            //slueConsumpUnit.Properties.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;


        }

        private void LoadData()
        {
            //StringBuilder sbSQL = new StringBuilder();
            //sbSQL.Append("SELECT OIDPayment AS No, Name, Description, DuedateCalculation, Status, CreatedBy, CreatedDate ");
            //sbSQL.Append("FROM PaymentTerm ");
            //sbSQL.Append("ORDER BY OIDPayment ");
            //new ObjDE.setGridControl(gridControl1, gridView1, sbSQL).getData(false, false, false, true);
        }

        private void NewData()
        {
            //txeName.Text = "";
            //lblStatus.Text = "* Add Payment Term";
            //lblStatus.ForeColor = Color.Green;

            //txeID.Text = DBC.DBQuery("SELECT CASE WHEN ISNULL(MAX(OIDPayment), '') = '' THEN 1 ELSE MAX(OIDPayment) + 1 END AS NewNo FROM PaymentTerm").getString();
            //txeDescription.Text = "";
            //txeDueDate.Text = "";
            //rgStatus.SelectedIndex = -1;

            //txeCREATE.Text = "0";
            //txeDATE.Text = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");

            //txeID.Focus();
        }

        private void newMain()
        {
            //MessageBox.Show("newMain");
            lblID.Text = "";
            dosetOIDSMPL                                = string.Empty;
            bbiSave.Enabled                             = true;
            bbiEdit.Visibility                          = DevExpress.XtraBars.BarItemVisibility.Never;
            tabbedControlGroup1.SelectedTabPageIndex    = 1;
            btnGenSMPLNo.Enabled                        = false;
            //gcQtyRequired.DataSource                     = db.FGRequestDS();
            gvQtyRequired.CloseEditor();
            gvQtyRequired.UpdateCurrentRow();
            gcQtyRequired.DataSource = null;
            dtQtyRequired.Rows.Clear();
            gcQtyRequired.DataSource                    = dtQtyRequired;
            gcQtyRequired.Enabled                       = true;

            /*UnLock Field*/
            glSaleSection_Main.Enabled      = true;
            txtReferenceNo_Main.Enabled     = true;
            glSeason_Main.Enabled           = true;
            slCustomer_Main.Enabled         = true;
            slStyleName_Main.Enabled        = true;
            txtSMPLPatternNo_Main.Enabled   = true;
            radioGroup3.Enabled             = true;

            /*TextEdit*/
            txtSMPLNo.EditValue = "";
            //lblStatus.Text = "New";
            txtReferenceNo_Main.EditValue = "";
            txtContactName_Main.EditValue = "";
            txtSMPLItemNo_Main.EditValue = "";
            txtModelName_Main.EditValue = "";
            txtSMPLPatternNo_Main.EditValue = "";
            txtSituation_Main.EditValue = "";
            txtStateArrangments_Main.EditValue = "";
            txtPictureFile_Main.Text = "";
            picMain.Image = null;

            /*GridLookup*/
            glBranch_Main.EditValue = "";
            glSaleSection_Main.EditValue = "";
            glSeason_Main.EditValue = "";
            glCategoryDivision_Main.EditValue = "";
            glUseFor.EditValue = "";

            /*SearchLookup*/
            slCustomer_Main.EditValue   = "";
            slStyleName_Main.EditValue  = "";

            /*RadioGroup*/
            radioGroup1.EditValue = 0;
            radioGroup3.EditValue = 0;
            radioGroup4.EditValue = 0;
            radioGroup5.EditValue = 1;
            radioGroup6.EditValue = 1;

            /*DateTime*/
            dtDeliveryRequest_Main.EditValue    = DateTime.Now;
            dtCustomerApproved_Main.EditValue   = DateTime.Now;
            dtACPRBy_Main.EditValue             = DateTime.Now;
            dtFBPRBy_Main.EditValue             = DateTime.Now;
        }

        //private void newFabric()
        //{
        //    //MessageBox.Show("newFabric");
        //    if (dosetOIDSMPL != "")
        //    {
        //        bbiSave.Visibility      = DevExpress.XtraBars.BarItemVisibility.Always;
        //        bbiRefresh.Visibility   = DevExpress.XtraBars.BarItemVisibility.Always;
        //        bbiEdit.Visibility      = DevExpress.XtraBars.BarItemVisibility.Never;
        //        bbiSave.Enabled = true;
        //        gridControl3.Enabled = true;
        //        btngetListFB_FB.Enabled = true;

        //        //Set New OIDFB
        //        txtFabricRacordID_FB.EditValue = db.get_newOIDFB();

        //        //Clear Form
        //        txtVendorFBCode_FB.EditValue = null;
        //        txtSampleLotNo_FB.EditValue = null;
        //        slVendor_FB.EditValue = null;
        //        slFBColor_FB.EditValue = null;
        //        slFBCode_FB.EditValue = null;
        //        slFGColor_FB.EditValue = null;
        //        txtComposition_FB.EditValue = null;
        //        txtWeightFB_FB.EditValue = null;
        //        txtWidthCuttable_FB.EditValue = null;
        //        txtPrice_FB.EditValue = null;
        //        glCurrency_FB.EditValue = null;
        //        txtTotalWidth_FB.EditValue = null;
        //        txtUsableWidth_FB.EditValue = null;
        //        txtImgUpload_FB.EditValue = null;
        //        picUpload_FB.Image = null;

        //        gcSize_Fabric.DataSource = null;
        //        gcList_Fabric.DataSource = null;

        //        db.getGrid_FBListSample(gridControl3, " AND smplQR.OIDSMPL = '" + dosetOIDSMPL + "' ");
        //        //db.getDgv("Select OIDGParts,GarmentParts From GarmentParts", gcPart_Fabric, mainConn);
        //    }
        //    else
        //    {
        //        FUNCT.msgWarning("Please Back to Select List of Sample Request!");
        //        tabbedControlGroup1.SelectedTabPageIndex = 0;
        //    }
        //}

        private void SetReadOnly()
        {
            //bbiSave.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;

            //Main Tab
            radioGroup4.ReadOnly = true;
            //radioGroup4.BackColor = Color.White;
            //radioGroup4.ForeColor = Color.Black;

            dtCustomerApproved_Main.ReadOnly = true;
            //dtCustomerApproved_Main.BackColor = Color.White;
            //dtCustomerApproved_Main.ForeColor = Color.Black;

            glBranch_Main.ReadOnly = true;
            //glBranch_Main.BackColor = Color.White;
            //glBranch_Main.ForeColor = Color.Black;

            glSaleSection_Main.ReadOnly = true;
            //glSaleSection_Main.BackColor = Color.White;
            //glSaleSection_Main.ForeColor = Color.Black;

            txtReferenceNo_Main.ReadOnly = true;
            //txtReferenceNo_Main.BackColor = Color.White;
            //txtReferenceNo_Main.ForeColor = Color.Black;

            dtRequestDate_Main.ReadOnly = true;
            //dtRequestDate_Main.BackColor = Color.White;
            //dtRequestDate_Main.ForeColor = Color.Black;

            radioGroup1.ReadOnly = true;
            //radioGroup1.BackColor = Color.White;
            //radioGroup1.ForeColor = Color.Black;

            glSeason_Main.ReadOnly = true;
            //glSeason_Main.BackColor = Color.White;
            //glSeason_Main.ForeColor = Color.Black;

            slCustomer_Main.ReadOnly = true;
            //slCustomer_Main.BackColor = Color.White;
            //slCustomer_Main.ForeColor = Color.Black;

            txtContactName_Main.ReadOnly = true;
            //txtContactName_Main.BackColor = Color.White;
            //txtContactName_Main.ForeColor = Color.Black;

            dtDeliveryRequest_Main.ReadOnly = true;
            //dtDeliveryRequest_Main.BackColor = Color.White;
            //dtDeliveryRequest_Main.ForeColor = Color.Black;

            glUseFor.ReadOnly = true;
            //glUseFor.BackColor = Color.White;
            //glUseFor.ForeColor = Color.Black;

            txtSMPLItemNo_Main.ReadOnly = true;
            //txtSMPLItemNo_Main.BackColor = Color.White;
            //txtSMPLItemNo_Main.ForeColor = Color.Black;

            txtModelName_Main.ReadOnly = true;
            //txtModelName_Main.BackColor = Color.White;
            //txtModelName_Main.ForeColor = Color.Black;

            glCategoryDivision_Main.ReadOnly = true;
            //glCategoryDivision_Main.BackColor = Color.White;
            //glCategoryDivision_Main.ForeColor = Color.Black;

            slStyleName_Main.ReadOnly = true;
            //slStyleName_Main.BackColor = Color.White;
            //slStyleName_Main.ForeColor = Color.Black;

            txtSMPLPatternNo_Main.ReadOnly = true;
            //txtSMPLPatternNo_Main.BackColor = Color.White;
            //txtSMPLPatternNo_Main.ForeColor = Color.Black;

            radioGroup3.ReadOnly = true;
            //radioGroup3.BackColor = Color.White;
            //radioGroup3.ForeColor = Color.Black;

            txtSituation_Main.ReadOnly = true;
            //txtSituation_Main.BackColor = Color.White;
            //txtSituation_Main.ForeColor = Color.Black;

            txtStateArrangments_Main.ReadOnly = true;
            //txtStateArrangments_Main.BackColor = Color.White;
            //txtStateArrangments_Main.ForeColor = Color.Black;

            radioGroup5.ReadOnly = true;
            //radioGroup5.BackColor = Color.White;
            //radioGroup5.ForeColor = Color.Black;

            dtACPRBy_Main.ReadOnly = true;
            //dtACPRBy_Main.BackColor = Color.White;
            //dtACPRBy_Main.ForeColor = Color.Black;

            radioGroup6.ReadOnly = true;
            //radioGroup6.BackColor = Color.White;
            //radioGroup6.ForeColor = Color.Black;

            dtFBPRBy_Main.ReadOnly = true;
            //dtFBPRBy_Main.BackColor = Color.White;
            //dtFBPRBy_Main.ForeColor = Color.Black;

            slueUnit.ReadOnly = true;
            //slueUnit.BackColor = Color.White;
            //slueUnit.ForeColor = Color.Black;

            txeQtyDF.ReadOnly = true;
            //txeQtyDF.BackColor = Color.White;

            sbUnit.Enabled = false;
            sbColor.Enabled = false;
            sbSize.Enabled = false;
            gvQtyRequired.OptionsBehavior.Editable = false;
            btnOpenImg_Main.Enabled = false;
            sbDelete_S.Enabled = false;

            simpleButton2.Enabled = false;
            sbUseFor.Enabled = false;
            simpleButton3.Enabled = false;
            simpleButton4.Enabled = false;
            picMain.Enabled = false;

            //Fabric Tab
            simpleButton5.Enabled = false;
            sbDelete_F.Enabled = false;
            sbClear.Enabled = false;
            sbFBSupplier.Enabled = false;
            sbFBColor.Enabled = false;
            sbTempCode.Enabled = false;
            sbPart.Enabled = false;
            btngetListFB_FB.Enabled = false;
            sbDeleteRow.Enabled = false;

            txtVendorFBCode_FB.ReadOnly = true;
            //txtVendorFBCode_FB.BackColor = Color.White;
            //txtVendorFBCode_FB.ForeColor = Color.Black;

            txtSampleLotNo_FB.ReadOnly = true;
            //txtSampleLotNo_FB.BackColor = Color.White;
            //txtSampleLotNo_FB.ForeColor = Color.Black;

            slVendor_FB.ReadOnly = true;
            //slVendor_FB.BackColor = Color.White;
            //slVendor_FB.ForeColor = Color.Black;

            slFBColor_FB.ReadOnly = true;
            //slFBColor_FB.BackColor = Color.White;
            //slFBColor_FB.ForeColor = Color.Black;

            slFBCode_FB.ReadOnly = true;
            //slFBCode_FB.BackColor = Color.White;
            //slFBCode_FB.ForeColor = Color.Black;

            txtComposition_FB.ReadOnly = true;
            //txtComposition_FB.BackColor = Color.White;
            //txtComposition_FB.ForeColor = Color.Black;

            txtWeightFB_FB.ReadOnly = true;
            //txtWeightFB_FB.BackColor = Color.White;
            //txtWeightFB_FB.ForeColor = Color.Black;
            
            txeRemark_FB.ReadOnly = true;

            txtWidthCuttable_FB.ReadOnly = true;
            //txtWidthCuttable_FB.BackColor = Color.White;
            //txtWidthCuttable_FB.ForeColor = Color.Black;

            txtPrice_FB.ReadOnly = true;
            //txtPrice_FB.BackColor = Color.White;
            //txtPrice_FB.ForeColor = Color.Black;

            glCurrency_FB.ReadOnly = true;
            //glCurrency_FB.BackColor = Color.White;
            //glCurrency_FB.ForeColor = Color.Black;

            txtTotalWidth_FB.ReadOnly = true;
            //txtTotalWidth_FB.BackColor = Color.White;
            //txtTotalWidth_FB.ForeColor = Color.Black;

            txtUsableWidth_FB.ReadOnly = true;
            //txtUsableWidth_FB.BackColor = Color.White;
            //txtUsableWidth_FB.ForeColor = Color.Black;

            picUpload_FB.Enabled = false;

            //Material Tab
            gridView6.OptionsBehavior.Editable = false;
            gridView7.OptionsBehavior.Editable = false;
            sbMatClear.Enabled = false;
            sbMTSupplier.Enabled = false;
            sbMTColor.Enabled = false;
            btnUploadMat.Enabled = false;
            simpleButton1.Enabled = false;
            sbTempCodeMat.Enabled = false;
            btnGettoLlist_Mat.Enabled = false;
            btnMatDelete.Enabled = false;

            glWorkStation_Mat.ReadOnly = true;
            //glWorkStation_Mat.BackColor = Color.White;
            //glWorkStation_Mat.ForeColor = Color.Black;

            slVendor_Mat.ReadOnly = true;
            //slVendor_Mat.BackColor = Color.White;
            //slVendor_Mat.ForeColor = Color.Black;

            txtVendorMatCode_Mat.ReadOnly = true;
            //txtVendorMatCode_Mat.BackColor = Color.White;
            //txtVendorMatCode_Mat.ForeColor = Color.Black;

            txtSampleLotNo_Mat.ReadOnly = true;
            //txtSampleLotNo_Mat.BackColor = Color.White;
            //txtSampleLotNo_Mat.ForeColor = Color.Black;

            txtMatComposition_Mat.ReadOnly = true;
            //txtMatComposition_Mat.BackColor = Color.White;
            //txtMatComposition_Mat.ForeColor = Color.Black;

            slMatColor_Mat.ReadOnly = true;
            //slMatColor_Mat.BackColor = Color.White;
            //slMatColor_Mat.ForeColor = Color.Black;

            slueConsumpUnit.ReadOnly = true;
            //slueConsumpUnit.BackColor = Color.White;
            //slueConsumpUnit.ForeColor = Color.Black;

            slMatCode_Mat.ReadOnly = true;
            //slMatCode_Mat.BackColor = Color.White;
            //slMatCode_Mat.ForeColor = Color.Black;

            txtPrice_Mat.ReadOnly = true;
            //txtPrice_Mat.BackColor = Color.White;
            //txtPrice_Mat.ForeColor = Color.Black;

            glCurrency_Mat.ReadOnly = true;
            //glCurrency_Mat.BackColor = Color.White;
            //glCurrency_Mat.ForeColor = Color.Black;

            txtSituation_Mat.ReadOnly = true;
            //txtSituation_Mat.BackColor = Color.White;
            //txtSituation_Mat.ForeColor = Color.Black;

            txtComment_Mat.ReadOnly = true;
            //txtComment_Mat.BackColor = Color.White;
            //txtComment_Mat.ForeColor = Color.Black;

            txtRemark_Mat.ReadOnly = true;
            //txtRemark_Mat.BackColor = Color.White;
            //txtRemark_Mat.ForeColor = Color.Black;

            picMat.Enabled = false;
        }


        private void SetWrite()
        {
            //bbiSave.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
            //Main Tab
            radioGroup4.ReadOnly = false;
            dtCustomerApproved_Main.ReadOnly = false;
            glBranch_Main.ReadOnly = false;
            glSaleSection_Main.ReadOnly = false;
            txtReferenceNo_Main.ReadOnly = false;
            dtRequestDate_Main.ReadOnly = false;
            radioGroup1.ReadOnly = false;
            glSeason_Main.ReadOnly = false;
            slCustomer_Main.ReadOnly = false;
            txtContactName_Main.ReadOnly = false;
            dtDeliveryRequest_Main.ReadOnly = false;
            glUseFor.ReadOnly = false;
            txtSMPLItemNo_Main.ReadOnly = false;
            txtModelName_Main.ReadOnly = false;
            glCategoryDivision_Main.ReadOnly = false;
            slStyleName_Main.ReadOnly = false;
            txtSMPLPatternNo_Main.ReadOnly = false;
            radioGroup3.ReadOnly = false;
            txtSituation_Main.ReadOnly = false;
            txtStateArrangments_Main.ReadOnly = false;
            radioGroup5.ReadOnly = false;
            dtACPRBy_Main.ReadOnly = false;
            radioGroup6.ReadOnly = false;
            dtFBPRBy_Main.ReadOnly = false;
            slueUnit.ReadOnly = false;

            txeQtyDF.ReadOnly = false;

            sbUnit.Enabled = true;
            sbColor.Enabled = true;
            sbSize.Enabled = true;
            gvQtyRequired.OptionsBehavior.Editable = true;
            btnOpenImg_Main.Enabled = true;
            sbDelete_S.Enabled = true;

            simpleButton2.Enabled = true;
            sbUseFor.Enabled = true;
            simpleButton3.Enabled = true;
            simpleButton4.Enabled = true;
            picMain.Enabled = true;

            //Fabric Tab
            simpleButton5.Enabled = true;
            sbDelete_F.Enabled = true;
            sbClear.Enabled = true;
            sbFBSupplier.Enabled = true;
            sbFBColor.Enabled = true;
            sbTempCode.Enabled = true;
            sbPart.Enabled = true;
            btngetListFB_FB.Enabled = true;
            sbDeleteRow.Enabled = true;

            txtVendorFBCode_FB.ReadOnly = false;
            txtSampleLotNo_FB.ReadOnly = false;
            slVendor_FB.ReadOnly = false;
            slFBColor_FB.ReadOnly = false;
            slFBCode_FB.ReadOnly = false;
            txtComposition_FB.ReadOnly = false;
            txtWeightFB_FB.ReadOnly = false;
            txeRemark_FB.ReadOnly = false;
            txtWidthCuttable_FB.ReadOnly = false;
            txtPrice_FB.ReadOnly = false;
            glCurrency_FB.ReadOnly = false;
            txtTotalWidth_FB.ReadOnly = false;
            txtUsableWidth_FB.ReadOnly = false;
            picUpload_FB.Enabled = true;

            //Material Tab
            gridView6.OptionsBehavior.Editable = true;
            gridView7.OptionsBehavior.Editable = true;
            sbMatClear.Enabled = true;
            sbMTSupplier.Enabled = true;
            sbMTColor.Enabled = true;
            btnUploadMat.Enabled = true;
            simpleButton1.Enabled = true;
            sbTempCodeMat.Enabled = true;
            btnGettoLlist_Mat.Enabled = true;
            btnMatDelete.Enabled = true;

            glWorkStation_Mat.ReadOnly = false;
            slVendor_Mat.ReadOnly = false;
            txtVendorMatCode_Mat.ReadOnly = false;
            txtSampleLotNo_Mat.ReadOnly = false;
            txtMatComposition_Mat.ReadOnly = false;
            slMatColor_Mat.ReadOnly = false;
            slueConsumpUnit.ReadOnly = false;
            slMatCode_Mat.ReadOnly = false;
            txtPrice_Mat.ReadOnly = false;
            glCurrency_Mat.ReadOnly = false;
            txtSituation_Mat.ReadOnly = false;
            txtComment_Mat.ReadOnly = false;
            txtRemark_Mat.ReadOnly = false;
            picMat.Enabled = true;

        }

        //private void newMaterials()
        //{
        //    //MessageBox.Show("newMaterials");
        //    if (dosetOIDSMPL == "")
        //    {
        //        ct.showInfoMessage("ไม่สามารถทำรายการได้ กรุณากลับไปเลือกรายการ SMPLNo ใหม่!"); return;
        //    }
        //    else
        //    {
        //        status_Mat = "new";
        //        bbiSave.Visibility          = DevExpress.XtraBars.BarItemVisibility.Always;
        //        bbiEdit.Visibility          = DevExpress.XtraBars.BarItemVisibility.Never;
        //        gridControl6.Enabled        = true;
        //        btnGettoLlist_Mat.Enabled   = true;
        //        gridControl7.DataSource     = null;

        //        // Set New OIDMat
        //        txtMatRecordID_Mat.EditValue = db.get_newOIDMat();

        //        // Clear Form
        //        glWorkStation_Mat.EditValue = null;
        //        slVendor_Mat.EditValue      = null;
        //        slVendor_Mat.EditValue      = null;
        //        slMatColor_Mat.EditValue    = null;
        //        slMatCode_Mat.EditValue     = null;
        //        glCurrency_Mat.EditValue    = null;

        //        txtVendorMatCode_Mat.Text = "";
        //        txtSampleLotNo_Mat.Text = "";
        //        txtMatComposition_Mat.Text = "";
        //        txtPrice_Mat.Text = "";
        //        txtSituation_Mat.Text = "";
        //        txtComment_Mat.Text = "";
        //        txtRemark_Mat.Text = "";
        //        txtPathFile_Mat.Text = "";
        //        picMat.Image = null;

        //        // Reload ListofMaterial
        //        db.getListofMaterial(gridControl8, dosetOIDSMPL);
        //    }
        //}

        private void bbiNew_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (lblStatus.Text.Trim() == "New SMPL")
            {
                LoadNewData();
                tabbedControlGroup1.SelectedTabPage = layoutControlGroup2;
                txtSMPLNo.Focus();
            }
            else
            {
                if (FUNCT.msgQuiz("Confirm clear data for new sample request ?") == true)
                {
                    LoadNewData();
                    tabbedControlGroup1.SelectedTabPage = layoutControlGroup2;
                    txtSMPLNo.Focus();
                }
            }
        }

        private void saveSampleRequest()
        {
            gvQtyRequired.CloseEditor();
            gvQtyRequired.UpdateCurrentRow();
            //MessageBox.Show("saveMain");
            string Season = glSeason_Main.Text.ToString().Trim().Replace("'", "''");
            if (Season == "") { FUNCT.msgWarning("Please select season!"); glSeason_Main.Focus(); return; }
            string SaleSection = glSaleSection_Main.EditValue.ToString();
            if (SaleSection == "") { FUNCT.msgWarning("Please select sales-section!"); glSaleSection_Main.Focus(); return; }
            string UseFor = glUseFor.EditValue.ToString();
            if (UseFor == "") { FUNCT.msgWarning("Please select use for."); glUseFor.Focus(); return; }
            string Unit = slueUnit.EditValue.ToString();
            if (Unit == "") { FUNCT.msgWarning("Please select unit."); slueUnit.Focus(); return; }


            string ACTION = "";
            string msgACTION = "Save";
            if (lblStatus.Text.Trim() == "New SMPL")
            {
                ACTION = "NEW";
                msgACTION = "Save New";
            }
            else if (lblStatus.Text.Trim() == "Update SMPL")
            {
                ACTION = "UPDATE";
                msgACTION = "Update";
            }
            else if (lblStatus.Text.Trim() == "Revise SMPL")
            {
                ACTION = "REVISE";
                msgACTION = "Save Revise";
            }
            else if (lblStatus.Text.Trim() == "Clone SMPL")
            {
                ACTION = "CLONE";
                msgACTION = "Save Clone";
            }

            layoutControlItem119.Text = msgACTION.Substring(0, 1).ToUpper() + msgACTION.Substring(1, msgACTION.Length - 1).ToLower() + " sample request processing ..";
            layoutControlItem119.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
            if (FUNCT.msgQuiz("Confirm " + msgACTION + " SampleRequest ?") == true)
            {
                //Check Required
                DataTable dtRequired = (DataTable)gcQtyRequired.DataSource;
                if (dtRequired == null)
                    return;
                else
                {
                    int runLoop = 0;
                    foreach (DataRow drRQ in dtRequired.Rows)
                    {
                        string strColor = drRQ["Color"].ToString();
                        string strSize = drRQ["Size"].ToString();
                        string strQuantity = drRQ["Quantity"].ToString();
                        strQuantity = strQuantity == "" ? "0" : strQuantity;

                        if (strColor != "" || strSize != "" || Convert.ToDouble(strQuantity) > 0)
                        {
                            if (strColor == "")
                            {
                                FUNCT.msgWarning("Please select color"); return;
                            }
                            else if (strSize == "")
                            {
                                FUNCT.msgWarning("Please select size"); return;
                            }
                            else if (strQuantity == "0")
                            {
                                FUNCT.msgWarning("Please input quantity required"); return;
                            }
                        }

                        runLoop++;
                    }
                }

                /*TextEdit*/
                string ReferenceNo = txtReferenceNo_Main.EditValue.ToString().Trim().Replace("'", "''");
                string ContactName = txtContactName_Main.EditValue.ToString().Trim().Replace("'", "''");
                string SMPLItem = txtSMPLItemNo_Main.EditValue.ToString().Trim().Replace("'", "''");
                string ModelName = txtModelName_Main.EditValue.ToString().Trim().Replace("'", "''");
                string SMPLPatternNo = txtSMPLPatternNo_Main.EditValue.ToString().Trim().Replace("'", "''");
                string Situation = txtSituation_Main.EditValue.ToString().Trim().Replace("'", "''");
                string StateArrangements = txtStateArrangments_Main.EditValue.ToString().Trim().Replace("'", "''");
                //string PictureFile = txtPictureFile_Main.EditValue.ToString().Trim().Replace("'", "''");
                string CreatedBy = UserLogin.OIDUser.ToString() != "" ? UserLogin.OIDUser.ToString() : "0";
                string CreatedDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                string UpdatedBy = UserLogin.OIDUser.ToString() != "" ? UserLogin.OIDUser.ToString() : "0";
                string UpdatedDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                ///*GridLookup*/
                string OIDBranch = glBranch_Main.Text.Trim() != "" ? "'" + glBranch_Main.EditValue.ToString() + "'" : "NULL";

                string OIDCATEGORY = glCategoryDivision_Main.Text.Trim() != "" ? "'" + glCategoryDivision_Main.EditValue.ToString() + "'" : "NULL";

                ///*SearchLookup*/
                string OIDCUST = slCustomer_Main.Text.Trim() != "" ? "'" + slCustomer_Main.EditValue.ToString() + "'" : "NULL";
                string OIDSTYLE = slStyleName_Main.Text.Trim() != "" ? "'" + slStyleName_Main.EditValue.ToString() + "'" : "NULL";

                ///*RadioGroup*/
                int SpecificationSize = Convert.ToInt32(radioGroup1.EditValue.ToString());
                
                int PatternSizeZone = Convert.ToInt32(radioGroup3.EditValue.ToString());
                int CustApproved = Convert.ToInt32(radioGroup4.EditValue.ToString());
                int ACPurRecBy = Convert.ToInt32(radioGroup5.EditValue.ToString());
                int FBPurRecBy = Convert.ToInt32(radioGroup6.EditValue.ToString());

                ///*DateTime*/
                string RequestDate = dtRequestDate_Main.Text.Trim() != "" ? "'" + Convert.ToDateTime(dtRequestDate_Main.EditValue.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL";
                string DeliveryRequest = dtDeliveryRequest_Main.Text.Trim() != "" ? "'" + Convert.ToDateTime(dtDeliveryRequest_Main.EditValue.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL";
                string CustApprovedDate = "NULL";
                if (CustApproved != 0)
                    CustApprovedDate = dtCustomerApproved_Main.Text.Trim() != "" ? "'" + Convert.ToDateTime(dtCustomerApproved_Main.EditValue.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL";
                string ACPurRecDate = "NULL";
                if (ACPurRecBy != 0)
                    ACPurRecDate = dtACPRBy_Main.Text.Trim() != "" ? "'" + Convert.ToDateTime(dtACPRBy_Main.EditValue.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL";
                string FBPurRecDate = "NULL";
                if (FBPurRecBy != 0)
                    FBPurRecDate = dtFBPRBy_Main.Text.Trim() != "" ? "'" + Convert.ToDateTime(dtFBPRBy_Main.EditValue.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL";

                // Check is null or Empty Data

                if (CustApproved == 1)
                {
                    if (OIDCUST == "") { FUNCT.msgWarning("Please select customer!"); slCustomer_Main.Focus(); return; }
                    if (OIDCATEGORY == "") { FUNCT.msgWarning("Please select category!"); glCategoryDivision_Main.Focus(); return; }
                    if (OIDSTYLE == "") { FUNCT.msgWarning("Please select style!"); slStyleName_Main.Focus(); return; }
                    if (OIDBranch == "") { FUNCT.msgWarning("Please select branch!"); glBranch_Main.Focus(); return; }
                    if (ContactName == "") { FUNCT.msgWarning("Please select contact-name!"); txtContactName_Main.Focus(); return; }
                    if (SMPLItem == "") { FUNCT.msgWarning("Please select SMPL Item No.!"); txtSMPLItemNo_Main.Focus(); return; }
                    if (ModelName == "") { FUNCT.msgWarning("Please select model name!"); txtModelName_Main.Focus(); return; }
                    if (SMPLPatternNo == "") { FUNCT.msgWarning("Please select SMPL Pattern No.!"); txtSMPLPatternNo_Main.Focus(); return; }
                }

                
                pbcSAVE.Properties.Step = 1;
                pbcSAVE.Properties.PercentView = true;
                pbcSAVE.Properties.Maximum = 5;
                pbcSAVE.Properties.Minimum = 0;

                StringBuilder sbSQL = new StringBuilder();
                string SMPLNo = "";
                int SMPLRevise = 0;
                string DEPCode = DBC.DBQuery("SELECT Code FROM Departments WHERE (OIDDEPT = '" + SaleSection + "') ").getString();
                string SMPLID = lblID.Text.Trim();

                if (ACTION == "NEW" || ACTION == "CLONE")
                {
                    StringBuilder sbGEN = new StringBuilder();
                    sbGEN.Append("SELECT TOP (1) FORMAT(CAST(SUBSTRING(SMPLNo, CASE WHEN CHARINDEX('-', SMPLNo) > 0 THEN CHARINDEX('-', SMPLNo) - 4 ELSE LEN(SMPLNo) - 3 END, 4) + 1 AS Int), '0000') AS genD4 ");
                    sbGEN.Append("FROM SMPLRequest ");
                    sbGEN.Append("WHERE (SMPLNo LIKE N'" + Season + DEPCode + "%') AND (LEN(SMPLNo) > 9) ");
                    sbGEN.Append("ORDER BY genD4 DESC ");
                    string strRUN = DBC.DBQuery(sbGEN.ToString()).getString();
                    if (strRUN == "")
                        strRUN = "0001";
                    SMPLNo = Season + DEPCode + strRUN;
                    SMPLRevise = 0;
                }
                else if (ACTION == "UPDATE")
                {
                    StringBuilder sbSMPL = new StringBuilder();
                    sbSMPL.Append("SELECT SMPLNo, SMPLRevise FROM SMPLRequest WHERE (OIDSMPL = '" + SMPLID + "') ");
                    string[] arrSMPL = DBC.DBQuery(sbSMPL.ToString()).getMultipleValue();
                    if (arrSMPL.Length > 0)
                    {
                        SMPLNo = arrSMPL[0];
                        SMPLRevise = Convert.ToInt32(arrSMPL[1]);
                    }
                }
                else if (ACTION == "REVISE")
                {
                    StringBuilder sbSMPL = new StringBuilder();
                    sbSMPL.Append("SELECT REPLACE({ fn CONCAT({ fn CONCAT(SUBSTRING(SMPLNo, 1, CASE WHEN CHARINDEX('-', SMPLNo) > 0 THEN CHARINDEX('-', SMPLNo) ELSE LEN(SMPLNo) END), '-') }, CONVERT(NVARCHAR, SMPLRevise + 1)) }, '--', '-') AS NewSMPLNo, SMPLRevise + 1 AS SMPLRevise ");
                    sbSMPL.Append("FROM   SMPLRequest ");
                    sbSMPL.Append("WHERE (OIDSMPL = ");
                    sbSMPL.Append("           (SELECT MAX(OIDSMPL) AS OIDSMPL ");
                    sbSMPL.Append("            FROM   SMPLRequest AS SR ");
                    sbSMPL.Append("            WHERE  (SUBSTRING(SMPLNo, 1, CASE WHEN CHARINDEX('-', SMPLNo) > 0 THEN CHARINDEX('-', SMPLNo) ELSE LEN(SMPLNo) END) = ");
                    sbSMPL.Append("                             (SELECT SUBSTRING(SMPLNo, 1, CASE WHEN CHARINDEX('-', SMPLNo) > 0 THEN CHARINDEX('-', SMPLNo) ELSE LEN(SMPLNo) END) AS SMPL ");
                    sbSMPL.Append("                              FROM   SMPLRequest AS xSR ");
                    sbSMPL.Append("                              WHERE  (OIDSMPL = '" + SMPLID + "'))))) ");
                    string[] arrSMPL = DBC.DBQuery(sbSMPL.ToString()).getMultipleValue();
                    if (arrSMPL.Length > 0)
                    {
                        SMPLNo = arrSMPL[0];
                        SMPLRevise = Convert.ToInt32(arrSMPL[1]);
                    }

                    sbSQL.Append("UPDATE SMPLRequest ");
                    sbSQL.Append("SET    SMPLStatus = 0 ");
                    sbSQL.Append("WHERE  (SUBSTRING(SMPLNo, 1, CASE WHEN CHARINDEX('-', SMPLNo) > 0 THEN CHARINDEX('-', SMPLNo) ELSE LEN(SMPLNo) END) = ");
                    sbSQL.Append("              (SELECT SUBSTRING(SMPLNo, 1, CASE WHEN CHARINDEX('-', SMPLNo) > 0 THEN CHARINDEX('-', SMPLNo) ELSE LEN(SMPLNo) END) AS SMPL ");
                    sbSQL.Append("               FROM   SMPLRequest AS xSR ");
                    sbSQL.Append("               WHERE (OIDSMPL = '" + SMPLID + "'))) ");

                }

                pbcSAVE.PerformStep();
                pbcSAVE.Update();

                //MessageBox.Show(SMPLNo);
                string PictureFile = "NULL";
                if(txtPictureFile_Main.Text.Trim() != "")
                    PictureFile = "N'" + uploadImg(txtPictureFile_Main, SMPLNo + "-FG" + DateTime.Now.ToString("-yyMMddHHmmss")) + "'";

                //Save SMPLRequest Table
                if (ACTION == "NEW" || ACTION == "CLONE" || ACTION == "REVISE")
                {
                    sbSQL.Append("INSERT INTO SMPLRequest(SMPLNo, SMPLRevise, Status, ReferenceNo, RequestDate, SMPLItem, ModelName, OIDCUST, OIDCATEGORY, OIDSTYLE, OIDBranch, OIDDEPT, SMPLPatternNo, PatternSizeZone, Season, SpecificationSize, ContactName, DeliveryRequest, UseFor, Situation, StateArrangements, PictureFile, CustApproved, CustApprovedDate, ACPurRecBy, ACPurRecDate, FBPurRecBy, FBPurRecDate, SMPLStatus, CreatedBy, CreatedDate, UpdatedBy, UpdatedDate) ");
                    sbSQL.Append(" VALUES(N'" + SMPLNo + "', '" + SMPLRevise + "', 0, N'" + ReferenceNo + "', " + RequestDate + ", N'" + SMPLItem + "', N'" + ModelName + "', " + OIDCUST + ", " + OIDCATEGORY + ", " + OIDSTYLE + ", " + OIDBranch + ", '" + SaleSection + "', N'" + SMPLPatternNo + "', '" + PatternSizeZone + "', N'" + Season + "', '" + SpecificationSize + "', N'" + ContactName + "', " + DeliveryRequest + ", '" + UseFor + "', N'" + Situation + "', N'" + StateArrangements + "', " + PictureFile + ", '" + CustApproved + "', " + CustApprovedDate + ", '" + ACPurRecBy + "', " + ACPurRecDate + ", '" + FBPurRecBy + "', " + FBPurRecDate + ", '1', '" + CreatedBy + "', '" + CreatedDate + "', '" + UpdatedBy + "', '" + UpdatedDate + "')  ");
                }
                else if (ACTION == "UPDATE")
                {
                    sbSQL.Append("UPDATE SMPLRequest SET ");
                    sbSQL.Append("  ReferenceNo=N'" + ReferenceNo + "', RequestDate=" + RequestDate + ", SMPLItem=N'" + SMPLItem + "', ");
                    sbSQL.Append("  ModelName=N'" + ModelName + "', OIDCUST=" + OIDCUST + ", OIDCATEGORY=" + OIDCATEGORY + ", OIDSTYLE=" + OIDSTYLE + ", ");
                    sbSQL.Append("  OIDBranch=" + OIDBranch + ", OIDDEPT='" + SaleSection + "', SMPLPatternNo=N'" + SMPLPatternNo + "', ");
                    sbSQL.Append("  PatternSizeZone='" + PatternSizeZone + "', Season=N'" + Season + "', SpecificationSize='" + SpecificationSize + "', ");
                    sbSQL.Append("  ContactName=N'" + ContactName + "', DeliveryRequest=" + DeliveryRequest + ", UseFor='" + UseFor + "', ");
                    sbSQL.Append("  Situation=N'" + Situation + "', StateArrangements=N'" + StateArrangements + "', PictureFile=" + PictureFile + ", ");
                    sbSQL.Append("  CustApproved='" + CustApproved + "', CustApprovedDate=" + CustApprovedDate + ", ACPurRecBy='" + ACPurRecBy + "', ");
                    sbSQL.Append("  ACPurRecDate=" + ACPurRecDate + ", FBPurRecBy='" + FBPurRecBy + "', FBPurRecDate=" + FBPurRecDate + ", ");
                    sbSQL.Append("  UpdatedBy='" + UpdatedBy + "', UpdatedDate='" + UpdatedDate + "'  ");
                    sbSQL.Append("WHERE (OIDSMPL = '" + SMPLID + "')  ");
                }

                //MessageBox.Show("A");
                bool chkSave = DBC.DBQuery(sbSQL.ToString()).runSQL();
                pbcSAVE.PerformStep();
                pbcSAVE.Update();

                if (chkSave == true)
                {
                    //Save SMPLQuantityRequired Table
                    sbSQL.Clear();
                    string maxOIDSMPL = "";
                    if (ACTION == "NEW" || ACTION == "CLONE" || ACTION == "REVISE")
                        maxOIDSMPL = DBC.DBQuery("SELECT OIDSMPL FROM SMPLRequest WHERE (SMPLNo='" + SMPLNo + "') AND (SMPLRevise='" + SMPLRevise + "')").getString();
                    else if (ACTION == "UPDATE")
                        maxOIDSMPL = SMPLID;

                    if (gvQtyRequired.RowCount > 0)
                    {
                        string COLOR_SIZE = "";
                        for (int j = 0; j < gvQtyRequired.RowCount - 1; j++)
                        {
                            string Color = gvQtyRequired.GetRowCellValue(j, "Color").ToString();
                            string Size = gvQtyRequired.GetRowCellValue(j, "Size").ToString();
                            string Quantity = gvQtyRequired.GetRowCellValue(j, "Quantity").ToString();

                            if (Color != "" && Size != "")
                            {
                                if (COLOR_SIZE != "")
                                    COLOR_SIZE += ", ";
                                COLOR_SIZE += "'" + Color + "-" + Size + "'";

                                if (ACTION == "NEW" || ACTION == "CLONE" || ACTION == "REVISE")
                                {
                                    sbSQL.Append("INSERT INTO SMPLQuantityRequired(OIDSMPL, OIDCOLOR, OIDSIZE, Quantity, OIDUnit) ");
                                    sbSQL.Append("  VALUES('" + maxOIDSMPL + "', '" + Color + "', '" + Size + "', '" + Quantity + "', '" + Unit + "')  ");
                                }
                                else if (ACTION == "UPDATE")
                                {
                                    sbSQL.Append("IF NOT EXISTS(SELECT OIDSMPLDT FROM SMPLQuantityRequired WHERE OIDSMPL='" + maxOIDSMPL + "' AND OIDCOLOR='" + Color + "' AND OIDSIZE='" + Size + "') ");
                                    sbSQL.Append(" BEGIN ");
                                    sbSQL.Append("  INSERT INTO SMPLQuantityRequired(OIDSMPL, OIDCOLOR, OIDSIZE, Quantity, OIDUnit) ");
                                    sbSQL.Append("  VALUES('" + maxOIDSMPL + "', '" + Color + "', '" + Size + "', '" + Quantity + "', '" + Unit + "')  ");
                                    sbSQL.Append(" END ");
                                    sbSQL.Append("ELSE ");
                                    sbSQL.Append(" BEGIN ");
                                    sbSQL.Append("  UPDATE SMPLQuantityRequired SET ");
                                    sbSQL.Append("    Quantity='" + Quantity + "', OIDUnit='" + Unit + "'  ");
                                    sbSQL.Append("  WHERE (OIDSMPL='" + maxOIDSMPL + "') AND (OIDCOLOR='" + Color + "') AND (OIDSIZE='" + Size + "') ");
                                    sbSQL.Append(" END   ");
                                }
                            }
                        }

                        if (ACTION == "UPDATE" && COLOR_SIZE != "")
                        {
                            sbSQL.Append("DELETE FROM SMPLQuantityRequired ");
                            sbSQL.Append("WHERE (OIDSMPL = '" + maxOIDSMPL + "') AND ((CONVERT(VARCHAR, OIDCOLOR) + '-' + CONVERT(VARCHAR, OIDSIZE)) NOT IN (" + COLOR_SIZE + "))  ");
                        }

                        if (sbSQL.Length > 0)
                        {
                            //MessageBox.Show(sbSQL.ToString());
                            chkSave = DBC.DBQuery(sbSQL.ToString()).runSQL();
                            pbcSAVE.PerformStep();
                            pbcSAVE.Update();

                            if (chkSave == true)
                            {
                                sbSQL.Clear();
                                //Fabric Tab
                                if (ACTION == "UPDATE") //Delete Fabric
                                {
                                    string tmpFBID = "";
                                    DataTable dtxFB = (DataTable)gcList_Fabric.DataSource;
                                    if (dtxFB != null)
                                    {
                                        foreach (DataRow rxFB in dtxFB.Rows)
                                        {
                                            string FBID = rxFB["FBID"].ToString();
                                            if (FBID != "")
                                            {
                                                if (tmpFBID != "")
                                                    tmpFBID += ", ";
                                                tmpFBID += "'" + FBID + "'";
                                            }
                                        }


                                        sbSQL.Append("DELETE FROM SMPLRequestFabricParts ");
                                        sbSQL.Append("WHERE (OIDSMPLDT IN (SELECT OIDSMPLDT FROM SMPLQuantityRequired WHERE (OIDSMPL = '" + maxOIDSMPL + "'))) ");
                                        if (tmpFBID != "")
                                        {
                                            sbSQL.Append("AND (OIDSMPLFB NOT IN (" + tmpFBID + "))  ");
                                        }

                                        sbSQL.Append("DELETE FROM SMPLRequestFabric ");
                                        sbSQL.Append("WHERE (OIDSMPLDT IN (SELECT OIDSMPLDT FROM SMPLQuantityRequired WHERE (OIDSMPL = '" + maxOIDSMPL + "'))) ");
                                        if (tmpFBID != "")
                                        {
                                            sbSQL.Append("AND (OIDSMPLFB NOT IN (" + tmpFBID + "))  ");
                                        }
                                    }
                                }


                                DataTable dtFB = (DataTable)gcList_Fabric.DataSource;
                                if (dtFB != null)
                                {
                                    //chk fabric code duplicate
                                    foreach (DataRow rFB in dtFB.Rows)
                                    {
                                        string FBID = rFB["FBID"].ToString().Trim();
                                        string ColorName = rFB["ColorName"].ToString();                 ColorName = ColorName.Trim() == "" ? "0" : ColorName.Trim();
                                        string VendorFBCode = rFB["VendorFBCode"].ToString();
                                        string SMPLotNo = rFB["SMPLotNo"].ToString();
                                        string Supplier = rFB["Supplier"].ToString();                   Supplier = Supplier.Trim() == "" ? "0" : Supplier.Trim();
                                        string FabricColor = rFB["FabricColor"].ToString();             FabricColor = FabricColor.Trim() == "" ? "0" : FabricColor.Trim();
                                        string FabricCode = rFB["FabricCode"].ToString();               FabricCode = FabricCode.Trim() == "" ? "NULL" : "'" + FabricCode.Trim() + "'";
                                        string Description = rFB["Description"].ToString();
                                        string Composition = rFB["Composition"].ToString();
                                        string FBWeight = rFB["FBWeight"].ToString();
                                        string WidthCut = rFB["WidthCut"].ToString();
                                        string Price = rFB["Price"].ToString();
                                        Price = Price == "" ? Price = "0" : Price;
                                        string Currency = rFB["Currency"].ToString();                   Currency = Currency.Trim() == "" ? "NULL" : "'" + Currency.Trim() + "'";
                                        string TTWidth = rFB["TTWidth"].ToString();
                                        TTWidth = TTWidth == "" ? "0" : TTWidth;
                                        string UsableWidth = rFB["UsableWidth"].ToString();
                                        UsableWidth = UsableWidth == "" ? "0" : UsableWidth;
                                        string Remark = rFB["Remark"].ToString();
                                        string PicFile = rFB["PicFile"].ToString();
                                        string PathFile = "NULL";
                                        if (PicFile != "")
                                        {
                                            string FileName = SMPLNo + "-FB-" + FabricColor;
                                            if (FabricCode != "NULL")
                                                FileName += "-" + FabricCode.Replace("'", "");

                                            PathFile = "N'" + uploadImg(PicFile, FileName + DateTime.Now.ToString("-yyMMddHHmmss")) + "'";
                                        }

                                        if (ACTION == "NEW" || ACTION == "CLONE" || ACTION == "REVISE")
                                        {
                                            sbSQL.Append("INSERT INTO SMPLRequestFabric(OIDSMPLDT, OIDCOLOR, OIDVEND, OIDITEM, VendFBCode, SMPLotNo, Composition, FBWeight, WidthCuttable, Price, OIDCURR, TotalWidth, UsableWidth, PathFile, Remark) ");
                                            sbSQL.Append("SELECT OIDSMPLDT, '" + FabricColor + "' AS OIDCOLOR, '" + Supplier + "' AS OIDVEND, " + FabricCode + " AS OIDITEM, N'" + VendorFBCode + "' AS VendFBCode, N'" + SMPLotNo + "' AS SMPLotNo, N'" + Composition + "' AS Composition, N'" + FBWeight + "' AS FBWeight, N'" + WidthCut + "' AS WidthCuttable, '" + Price + "' AS Price, " + Currency + " AS OIDCURR, '" + TTWidth + "' AS TotalWidth, '" + UsableWidth + "' AS UsableWidth, " + PathFile + " AS PathFile, N'" + Remark + "' AS Remark  ");
                                            sbSQL.Append("FROM   SMPLQuantityRequired  ");
                                            sbSQL.Append("WHERE  (OIDSMPL = '" + maxOIDSMPL + "') AND (OIDCOLOR = '" + ColorName + "')  ");
                                        }
                                        else if (ACTION == "UPDATE")
                                        {
                                            if (FBID == "")
                                            {
                                                sbSQL.Append("  INSERT INTO SMPLRequestFabric(OIDSMPLDT, OIDCOLOR, OIDVEND, OIDITEM, VendFBCode, SMPLotNo, Composition, FBWeight, WidthCuttable, Price, OIDCURR, TotalWidth, UsableWidth, PathFile, Remark) ");
                                                sbSQL.Append("  SELECT OIDSMPLDT, '" + FabricColor + "' AS OIDCOLOR, '" + Supplier + "' AS OIDVEND, " + FabricCode + " AS OIDITEM, N'" + VendorFBCode + "' AS VendFBCode, N'" + SMPLotNo + "' AS SMPLotNo, N'" + Composition + "' AS Composition, N'" + FBWeight + "' AS FBWeight, N'" + WidthCut + "' AS WidthCuttable, '" + Price + "' AS Price, " + Currency + " AS OIDCURR, '" + TTWidth + "' AS TotalWidth, '" + UsableWidth + "' AS UsableWidth, " + PathFile + " AS PathFile, N'" + Remark + "' AS Remark  ");
                                                sbSQL.Append("  FROM   SMPLQuantityRequired  ");
                                                sbSQL.Append("  WHERE  (OIDSMPL = '" + maxOIDSMPL + "') AND (OIDCOLOR = '" + ColorName + "')  ");
                                            }
                                            else
                                            {
                                                sbSQL.Append("  UPDATE SMPLRequestFabric SET ");
                                                sbSQL.Append("    OIDCOLOR='" + FabricColor + "', OIDVEND='" + Supplier + "', OIDITEM=" + FabricCode + ", ");
                                                sbSQL.Append("    VendFBCode=N'" + VendorFBCode + "', SMPLotNo=N'" + SMPLotNo + "', Composition=N'" + Composition + "', ");
                                                sbSQL.Append("    FBWeight=N'" + FBWeight + "', WidthCuttable=N'" + WidthCut + "', Price='" + Price + "', ");
                                                sbSQL.Append("    OIDCURR=" + Currency + ", TotalWidth='" + TTWidth + "', UsableWidth='" + UsableWidth + "', ");
                                                sbSQL.Append("    PathFile=" + PathFile + ", Remark=N'" + Remark + "'  ");
                                                sbSQL.Append("  WHERE (OIDSMPLFB = '" + FBID + "') ");
                                            }
                                        }
                                       
                                        //FABRIC PARTS
                                        string FBPartsID = rFB["FBPartsID"].ToString();
                                        if (FBPartsID.Trim() != "")
                                        {
                                            FBPartsID = FBPartsID.Trim().Replace(" ", "");
                                            if (FBPartsID != "")
                                            {
                                                string Condition = " WHERE (CONVERT(VARCHAR, OIDSMPLFB) + CONVERT(VARCHAR, OIDSMPLDT) = (SELECT CONVERT(VARCHAR, OIDSMPLFB) + CONVERT(VARCHAR, OIDSMPLDT) AS FB FROM SMPLRequestFabric WHERE (OIDCOLOR = '" + FabricColor + "') AND (OIDVEND = '" + Supplier + "') AND (OIDSMPLDT IN (SELECT OIDSMPLDT FROM SMPLQuantityRequired WHERE (OIDSMPL = '" + maxOIDSMPL + "') AND (OIDCOLOR = '" + ColorName + "'))))) ";

                                                string FBPart = "";
                                                if (FBPartsID.IndexOf(',') != -1)
                                                {
                                                    string[] ID = FBPartsID.Split(',');
                                                    if (ID.Length > 0)
                                                    {
                                                        foreach (string idPart in ID)
                                                        {
                                                            if (FBPart != "")
                                                                FBPart += ", ";
                                                            FBPart += "'" + idPart + "'";

                                                            if (ACTION == "NEW" || ACTION == "CLONE" || ACTION == "REVISE")
                                                            {
                                                                sbSQL.Append("INSERT INTO SMPLRequestFabricParts(OIDSMPLFB, OIDSMPLDT, OIDGParts) ");
                                                                sbSQL.Append("SELECT OIDSMPLFB, OIDSMPLDT, '" + idPart + "' AS OIDGParts ");
                                                                sbSQL.Append("FROM   SMPLRequestFabric ");
                                                                sbSQL.Append("WHERE  (OIDCOLOR='" + FabricColor + "') AND (OIDVEND='" + Supplier + "') AND (OIDSMPLDT IN ");
                                                                sbSQL.Append("            (SELECT OIDSMPLDT ");
                                                                sbSQL.Append("             FROM   SMPLQuantityRequired ");
                                                                sbSQL.Append("             WHERE  (OIDSMPL = '" + maxOIDSMPL + "') AND (OIDCOLOR = '" + ColorName + "')))  ");
                                                            }
                                                            else if (ACTION == "UPDATE")
                                                            {
                                                                sbSQL.Append("IF NOT EXISTS(SELECT OIDSMFBGParts FROM SMPLRequestFabricParts " + Condition + " AND (OIDGParts = '" + idPart + "')) ");
                                                                sbSQL.Append(" BEGIN ");
                                                                sbSQL.Append("  INSERT INTO SMPLRequestFabricParts(OIDSMPLFB, OIDSMPLDT, OIDGParts) ");
                                                                sbSQL.Append("  SELECT OIDSMPLFB, OIDSMPLDT, '" + idPart + "' AS OIDGParts ");
                                                                sbSQL.Append("  FROM   SMPLRequestFabric ");
                                                                sbSQL.Append("  WHERE  (OIDCOLOR='" + FabricColor + "') AND (OIDVEND='" + Supplier + "') AND (OIDSMPLDT IN ");
                                                                sbSQL.Append("            (SELECT OIDSMPLDT ");
                                                                sbSQL.Append("             FROM   SMPLQuantityRequired ");
                                                                sbSQL.Append("             WHERE  (OIDSMPL = '" + maxOIDSMPL + "') AND (OIDCOLOR = '" + ColorName + "')))  ");
                                                                sbSQL.Append(" END ");

                                                            }
                                                        }
                                                    }

                                                }
                                                else
                                                {
                                                    if (FBPart != "")
                                                        FBPart += ", ";
                                                    FBPart += "'" + FBPartsID + "'";

                                                    if (ACTION == "NEW" || ACTION == "CLONE" || ACTION == "REVISE")
                                                    {
                                                        sbSQL.Append("INSERT INTO SMPLRequestFabricParts(OIDSMPLFB, OIDSMPLDT, OIDGParts) ");
                                                        sbSQL.Append("SELECT OIDSMPLFB, OIDSMPLDT, '" + FBPartsID + "' AS OIDParts ");
                                                        sbSQL.Append("FROM   SMPLRequestFabric ");
                                                        sbSQL.Append("WHERE  (OIDCOLOR='" + FabricColor + "') AND (OIDVEND='" + Supplier + "') AND (OIDSMPLDT IN ");
                                                        sbSQL.Append("            (SELECT OIDSMPLDT ");
                                                        sbSQL.Append("             FROM   SMPLQuantityRequired ");
                                                        sbSQL.Append("             WHERE  (OIDSMPL = '" + maxOIDSMPL + "') AND (OIDCOLOR = '" + ColorName + "')))  ");
                                                    }
                                                    else if (ACTION == "UPDATE")
                                                    {
                                                        sbSQL.Append("IF NOT EXISTS(SELECT OIDSMFBGParts FROM SMPLRequestFabricParts " + Condition + " AND (OIDGParts = '" + FBPartsID + "')) ");
                                                        sbSQL.Append(" BEGIN ");
                                                        sbSQL.Append("  INSERT INTO SMPLRequestFabricParts(OIDSMPLFB, OIDSMPLDT, OIDGParts) ");
                                                        sbSQL.Append("  SELECT OIDSMPLFB, OIDSMPLDT, '" + FBPartsID + "' AS OIDGParts ");
                                                        sbSQL.Append("  FROM   SMPLRequestFabric ");
                                                        sbSQL.Append("  WHERE  (OIDCOLOR='" + FabricColor + "') AND (OIDVEND='" + Supplier + "') AND (OIDSMPLDT IN ");
                                                        sbSQL.Append("            (SELECT OIDSMPLDT ");
                                                        sbSQL.Append("             FROM   SMPLQuantityRequired ");
                                                        sbSQL.Append("             WHERE  (OIDSMPL = '" + maxOIDSMPL + "') AND (OIDCOLOR = '" + ColorName + "')))  ");
                                                        sbSQL.Append(" END ");
                                                    }
                                                }

                                                if (ACTION == "UPDATE")
                                                {
                                                    sbSQL.Append("DELETE FROM SMPLRequestFabricParts " + Condition + "  ");
                                                    if (FBPart != "")
                                                    {
                                                        sbSQL.Append("AND (OIDGParts NOT IN (" + FBPart + "))  ");
                                                    }
                                                }

                                            }
                                        }
                                    }
                                    //MessageBox.Show(sbSQL.ToString());
                                    if (sbSQL.Length > 0)
                                    {
                                        //MessageBox.Show("C");
                                        //textBox1.Text = sbSQL.ToString();
                                        chkSave = this.DBC.DBQuery(sbSQL.ToString()).runSQL();
                                        pbcSAVE.PerformStep();
                                        pbcSAVE.Update();

                                        if (chkSave == false)
                                        {
                                            //MessageBox.Show(sbSQL.ToString());
                                            FUNCT.msgERROR("Found problem on save.\nพบปัญหาในการบันทึกข้อมูล (Tab Fabric)");
                                        }
                                    }
                                }

                                if (chkSave == true)
                                {
                                    sbSQL.Clear();
                                    //Meterial Tab
                                    if (ACTION == "UPDATE") //Delete Fabric
                                    {
                                        string tmpMTID = "";
                                        DataTable dtxMT = (DataTable)gridControl8.DataSource;
                                        if (dtxMT != null)
                                        {
                                            foreach (DataRow rxMT in dtxMT.Rows)
                                            {
                                                string FBID = rxMT["MID"].ToString();
                                                if (FBID != "")
                                                {
                                                    if (tmpMTID != "")
                                                        tmpMTID += ", ";
                                                    tmpMTID += "'" + FBID + "'";
                                                }
                                            }

                                            sbSQL.Append("DELETE FROM SMPLRequestMaterial ");
                                            sbSQL.Append("WHERE (OIDSMPLDT IN (SELECT OIDSMPLDT FROM SMPLQuantityRequired WHERE (OIDSMPL = '" + maxOIDSMPL + "'))) ");
                                            if (tmpMTID != "")
                                            {
                                                sbSQL.Append("AND (OIDSMPLMT NOT IN (" + tmpMTID + "))  ");
                                            }
                                        }
                                    }


                                    DataTable dtMT = (DataTable)gridControl8.DataSource;
                                    if (dtMT != null)
                                    {
                                        string tmpMT = "";
                                        foreach (DataRow drMT in dtMT.Rows)
                                        {
                                            string MID = drMT["MID"].ToString().Trim();
                                            string cMatID = drMT["MatID"].ToString();
                                            string cSampleID = drMT["SampleID"].ToString();
                                            string cWorkStation = drMT["WorkStation"].ToString();                       cWorkStation = cWorkStation.Trim() == "" ? "NULL" : "'" + cWorkStation.Trim() + "'";
                                            string cVendMTCode = drMT["VendMTCode"].ToString();
                                            string cSMPLotNo = drMT["SMPLotNo"].ToString();
                                            string cVendor = drMT["Vendor"].ToString();                                 cVendor = cVendor.Trim() == "" ? "NULL" : "'" + cVendor.Trim() + "'";
                                            string cMatColor = drMT["MatColor"].ToString();                             
                                            string cColorName = drMT["ColorName"].ToString();
                                            string cMatSize = drMT["MatSize"].ToString();
                                            string cConsumption = drMT["Consumption"].ToString();
                                            string cUnit = drMT["Unit"].ToString();                                     cUnit = cUnit.Trim() == "" ? "NULL" : "'" + cUnit.Trim() + "'";
                                            string cComposition = drMT["Composition"].ToString().Replace("'", "''");
                                            string cDetails = drMT["Details"].ToString().Replace("'", "''");
                                            string cPrice = drMT["Price"].ToString();
                                            cPrice = cPrice == "" ? cPrice = "0" : cPrice;
                                            string cCurrency = drMT["Currency"].ToString();                             cCurrency = cCurrency.Trim() == "" ? "NULL" : "'" + cCurrency.Trim() + "'";
                                            string cNAVCode = drMT["NAVCode"].ToString();                               cNAVCode = cNAVCode.Trim() == "" ? "NULL" : "'" + cNAVCode.Trim() + "'";
                                            string cSituation = drMT["Situation"].ToString().Replace("'", "''");
                                            string cComment = drMT["Comment"].ToString().Replace("'", "''");
                                            string cRemark = drMT["Remark"].ToString().Replace("'", "''");
                                            string cPathFile = drMT["PathFile"].ToString().Replace("'", "''");
                                            string cChkMTCode = drMT["ChkMTCode"].ToString();
                                            string PathFile = "NULL";
                                            if (cPathFile != "")
                                            {
                                                string FileName = SMPLNo + "-MT-" + cMatColor + "-" + cMatSize;
                                                if (cNAVCode != "NULL")
                                                    FileName += "-" + cNAVCode.Replace("'", "");

                                                PathFile = "N'" + uploadImg(cPathFile, FileName+DateTime.Now.ToString("-yyMMddHHmmss")) + "'";
                                            }

                                            if (tmpMT == "")
                                                tmpMT += "WHERE ";
                                            else
                                                tmpMT += "OR ";
                                            tmpMT += "(OIDSMPLDT = (SELECT TOP (1) OIDSMPLDT FROM SMPLQuantityRequired WHERE (OIDSMPL = '" + maxOIDSMPL + "') AND (OIDCOLOR = '" + cColorName + "') AND (OIDSIZE = '" + cMatSize + "'))) AND (OIDITEM = " + cNAVCode + ") ";

                                            if (ACTION == "NEW" || ACTION == "CLONE" || ACTION == "REVISE")
                                            {
                                                sbSQL.Append("INSERT INTO SMPLRequestMaterial(OIDSMPLDT, OIDITEM, OIDVEND, OIDDEPT, OIDUNIT, OIDCURR, VendMTCode, SMPLotNo, MTColor, MTSize, Consumption, Composition, Details, Price, Situation, Comment, Remark, PathFile) ");
                                                sbSQL.Append(" SELECT (SELECT TOP (1) OIDSMPLDT FROM SMPLQuantityRequired WHERE (OIDSMPL = '" + maxOIDSMPL + "') AND (OIDCOLOR = '" + cColorName + "') AND (OIDSIZE = '" + cMatSize + "')) AS OIDSMPLDT, " + cNAVCode + " AS OIDITEM,  ");
                                                sbSQL.Append("        " + cVendor + " AS OIDVEND, " + cWorkStation + " AS OIDDEPT, " + cUnit + " AS OIDUNIT, " + cCurrency + " AS OIDCURR, N'" + cVendMTCode + "' AS VendMTCode, N'" + cSMPLotNo + "' AS SMPLotNo, ");
                                                sbSQL.Append("        N'" + cMatColor + "' AS MTColor, N'" + cMatSize + "' AS MTSize, '" + cConsumption + "' AS Consumption, N'" + cComposition + "' AS Composition, N'" + cDetails + "' AS Details, '" + cPrice + "' AS Price, ");
                                                sbSQL.Append("        N'" + cSituation + "' AS Situation, N'" + cComment + "' AS Comment, N'" + cRemark + "' AS Remark, " + PathFile + " AS PathFile   ");
                                            }
                                            else if (ACTION == "UPDATE")
                                            {
                                                if (MID == "")
                                                {
                                                    sbSQL.Append("  INSERT INTO SMPLRequestMaterial(OIDSMPLDT, OIDITEM, OIDVEND, OIDDEPT, OIDUNIT, OIDCURR, VendMTCode, SMPLotNo, MTColor, MTSize, Consumption, Composition, Details, Price, Situation, Comment, Remark, PathFile) ");
                                                    sbSQL.Append("  SELECT (SELECT TOP (1) OIDSMPLDT FROM SMPLQuantityRequired WHERE (OIDSMPL = '" + maxOIDSMPL + "') AND (OIDCOLOR = '" + cColorName + "') AND (OIDSIZE = '" + cMatSize + "')) AS OIDSMPLDT, " + cNAVCode + " AS OIDITEM,  ");
                                                    sbSQL.Append("          " + cVendor + " AS OIDVEND, " + cWorkStation + " AS OIDDEPT, " + cUnit + " AS OIDUNIT, " + cCurrency + " AS OIDCURR, N'" + cVendMTCode + "' AS VendMTCode, N'" + cSMPLotNo + "' AS SMPLotNo, ");
                                                    sbSQL.Append("          N'" + cMatColor + "' AS MTColor, N'" + cMatSize + "' AS MTSize, '" + cConsumption + "' AS Consumption, N'" + cComposition + "' AS Composition, N'" + cDetails + "' AS Details, '" + cPrice + "' AS Price, ");
                                                    sbSQL.Append("          N'" + cSituation + "' AS Situation, N'" + cComment + "' AS Comment, N'" + cRemark + "' AS Remark, " + PathFile + " AS PathFile   ");
                                                }
                                                else
                                                {
                                                    sbSQL.Append("  UPDATE SMPLRequestMaterial SET ");
                                                    sbSQL.Append("          OIDVEND=" + cVendor + ", OIDDEPT=" + cWorkStation + ", OIDUNIT=" + cUnit + ", OIDCURR=" + cCurrency + ", VendMTCode=N'" + cVendMTCode + "', SMPLotNo=N'" + cSMPLotNo + "', ");
                                                    sbSQL.Append("          MTColor=N'" + cMatColor + "', MTSize=N'" + cMatSize + "', Consumption='" + cConsumption + "', Composition=N'" + cComposition + "', Details=N'" + cDetails + "', Price='" + cPrice + "', ");
                                                    sbSQL.Append("          Situation=N'" + cSituation + "', Comment=N'" + cComment + "', Remark=N'" + cRemark + "', PathFile=" + PathFile + "   ");
                                                    sbSQL.Append("  WHERE (OIDSMPLMT = '" + MID + "') ");
                                                }
         
                                            }
                                        }

                                        if (sbSQL.Length > 0)
                                        {
                                            //MessageBox.Show("D");
                                            chkSave = DBC.DBQuery(sbSQL.ToString()).runSQL();
                                            pbcSAVE.PerformStep();
                                            pbcSAVE.Update();

                                            if (chkSave == false)
                                                FUNCT.msgERROR("Found problem on save.\nพบปัญหาในการบันทึกข้อมูล (Tab Material)");
                                        }

                                    }
                                        

                                }
                            }
                            else
                                FUNCT.msgERROR("Found problem on save.\nพบปัญหาในการบันทึกข้อมูล (Tab Main Quantity Required)");
                        }
                    }
                }
                else
                {
                    FUNCT.msgERROR("Found problem on save.\nพบปัญหาในการบันทึกข้อมูล (Tab Main)");
                }

                layoutControlItem119.Text = "Status ..";
                layoutControlItem119.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;

                if (chkSave == true)
                {
                    if (FUNCT.msgQuiz("Save sample request complete. Do you want to load this sample request ?") == true)
                    {
                        //Load Sample Request after save.
                        tabbedControlGroup1.SelectedTabPage = layoutControlGroup2;
                        LoadSampleRequestDocument(SMPLNo, "UPDATE");
                    }
                    else
                    {
                        //Clear All Document
                        tabbedControlGroup1.SelectedTabPage = layoutControlGroup1;
                        LoadNewData();

                    }
                }

            }
            layoutControlItem119.Text = "Status ..";
            layoutControlItem119.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
        }

        private void LoadSampleRequestDocument(string SMPLNo, string SMPLModify)
        {
            //ex
            //SMPLNo = "21FWS10001-0";

            LoadNewData(SMPLModify);
            SMPLNo = SMPLNo.ToUpper().Trim();
            if (SMPLNo != "")
            {
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT OIDSMPL, SMPLNo, SMPLRevise, Status, (CASE Status WHEN 0 THEN 'New' WHEN 1 THEN 'Wait Approved' WHEN 2 THEN 'Customer Approved' END) AS [Status Name], ReferenceNo, RequestDate, SMPLItem, ModelName, OIDCUST, ");
                sbSQL.Append("       OIDCATEGORY, OIDSTYLE, OIDBranch, OIDDEPT, SMPLPatternNo, PatternSizeZone, Season, SpecificationSize, ContactName, DeliveryRequest, UseFor, Situation, StateArrangements, PictureFile, CustApproved, CustApprovedDate, ");
                sbSQL.Append("       ACPurRecBy, ACPurRecDate, FBPurRecBy, FBPurRecDate, SMPLStatus, CreatedBy, CreatedDate, UpdatedBy, UpdatedDate ");
                sbSQL.Append("FROM   SMPLRequest ");
                sbSQL.Append("WHERE  (SMPLNo = N'" + SMPLNo + "') ");
                string[] arrSMPL = DBC.DBQuery(sbSQL.ToString()).getMultipleValue();
                if (arrSMPL.Length > 0)
                {
                    btnGenSMPLNo.Enabled = true;
                    //Load Main Tab
                    txtSMPLNo.Text = SMPLNo;
                    lblID.Text = arrSMPL[0];
                    //lblStatus.Text = arrSMPL[4];
                    txtReferenceNo_Main.Text = arrSMPL[5];

                    dtRequestDate_Main.EditValue = null;
                    if (arrSMPL[6] != "")
                        dtRequestDate_Main.EditValue = Convert.ToDateTime(arrSMPL[6]);

                    txtSMPLItemNo_Main.Text = arrSMPL[7];
                    txtModelName_Main.Text = arrSMPL[8];
                    slCustomer_Main.EditValue = arrSMPL[9];
                    glCategoryDivision_Main.EditValue = arrSMPL[10];
                    slStyleName_Main.EditValue = arrSMPL[11];
                    glBranch_Main.EditValue = arrSMPL[12];
                    glSaleSection_Main.EditValue = arrSMPL[13];
                    txtSMPLPatternNo_Main.Text = arrSMPL[14];
                    radioGroup3.EditValue = Convert.ToInt32(arrSMPL[15]);
                    glSeason_Main.EditValue = arrSMPL[16];
                    radioGroup1.EditValue = Convert.ToInt32(arrSMPL[17]);
                    txtContactName_Main.Text = arrSMPL[18];

                    dtDeliveryRequest_Main.EditValue = null;
                    if (arrSMPL[19] != "")
                        dtDeliveryRequest_Main.EditValue = Convert.ToDateTime(arrSMPL[19]);

                    glUseFor.EditValue = Convert.ToInt32(arrSMPL[20]);
                    txtSituation_Main.Text = arrSMPL[21];
                    txtStateArrangments_Main.Text = arrSMPL[22];

                    string picPath = imgPath + arrSMPL[23];
                    if (arrSMPL[23] != "")
                    {
                        txtPictureFile_Main.Text = picPath;
                        picMain.Image = null;
                        try
                        {
                            picMain.Image = Image.FromFile(picPath);
                        }
                        catch (Exception) { }
                    }
                    else
                    {
                        txtPictureFile_Main.Text = "";
                        picMain.Image = null;
                    }

                    radioGroup4.EditValue = Convert.ToInt32(arrSMPL[24]);
                    dtCustomerApproved_Main.EditValue = null;
                    if (arrSMPL[25] != "")
                        dtCustomerApproved_Main.EditValue = Convert.ToDateTime(arrSMPL[25]);

                    radioGroup5.EditValue = Convert.ToInt32(arrSMPL[26]);
                    dtACPRBy_Main.EditValue = null;
                    if (arrSMPL[27] != "")
                        dtACPRBy_Main.EditValue = Convert.ToDateTime(arrSMPL[27]);

                    radioGroup6.EditValue = Convert.ToInt32(arrSMPL[28]);
                    dtFBPRBy_Main.EditValue = null;
                    if (arrSMPL[29] != "")
                        dtFBPRBy_Main.EditValue = Convert.ToDateTime(arrSMPL[29]);

                    glueCreateBy_Main.EditValue = arrSMPL[31];
                    txtCreateDate_Main.Text = arrSMPL[32];
                    glueUpdateBy_Main.EditValue = arrSMPL[33];
                    txtUpdateDate_Main.Text = arrSMPL[34];

                    sbSQL.Clear();
                    sbSQL.Append("SELECT TOP (1) OIDUnit ");
                    sbSQL.Append("FROM   SMPLQuantityRequired ");
                    sbSQL.Append("WHERE  (OIDSMPL = ");
                    sbSQL.Append("           (SELECT OIDSMPL ");
                    sbSQL.Append("            FROM   SMPLRequest ");
                    sbSQL.Append("            WHERE  (SMPLNo = N'" + SMPLNo + "'))) ");
                    slueUnit.EditValue = DBC.DBQuery(sbSQL.ToString()).getString();

                    sbSQL.Clear();
                    sbSQL.Append("SELECT OIDCOLOR AS Color, OIDSIZE AS Size, Quantity, OIDSMPLDT AS ID ");
                    sbSQL.Append("FROM   SMPLQuantityRequired ");
                    sbSQL.Append("WHERE  (OIDSMPL = ");
                    sbSQL.Append("           (SELECT OIDSMPL ");
                    sbSQL.Append("            FROM   SMPLRequest ");
                    sbSQL.Append("            WHERE  (SMPLNo = N'" + SMPLNo + "'))) ");
                    sbSQL.Append("ORDER BY ID ");
                    int QCount = this.DBC.DBQuery(sbSQL).getCount();
                    if (QCount > 0)
                    {
                        new ObjDE.setGridControl(gcQtyRequired, gvQtyRequired, sbSQL).getData(false, false, true, true);
                    }
                    else
                    {
                        gcQtyRequired.DataSource = dtQtyRequired;
                    }

                    //Load Fabric Tab
                    txtSampleID_FB.Text = lblID.Text;
                    sbSQL.Clear();
                    sbSQL.Append("SELECT DISTINCT SQR.OIDSMPL, SRQ.SMPLPatternNo, SQR.OIDCOLOR AS ColorName, '' AS OIDSMPLDT ");
                    sbSQL.Append("FROM   SMPLQuantityRequired AS SQR INNER JOIN ");
                    sbSQL.Append("       SMPLRequest AS SRQ ON SQR.OIDSMPL = SRQ.OIDSMPL ");
                    sbSQL.Append("WHERE (SRQ.SMPLNo = N'" + SMPLNo + "') ");
                    sbSQL.Append("ORDER BY SQR.OIDSMPL, SRQ.SMPLPatternNo ");
                    new ObjDE.setGridControl(gridControl3, gridView3, sbSQL).getData(false, false, true, true);

                    sbSQL.Clear();
                    sbSQL.Append("SELECT DISTINCT SQR.OIDCOLOR AS ColorName, SFB.VendFBCode AS VendorFBCode, SFB.SMPLotNo, SFB.OIDVEND AS Supplier, SFB.OIDCOLOR AS FabricColor, SFB.OIDITEM AS FabricCode, ITM.Description, SFB.Composition, SFB.FBWeight, ");
                    sbSQL.Append("       SFB.WidthCuttable AS WidthCut, FORMAT(SFB.Price, '###.####') AS Price, SFB.OIDCURR AS Currency, FORMAT(SFB.TotalWidth, '###.####') AS TTWidth, FORMAT(SFB.UsableWidth, '###.####') AS UsableWidth, CASE WHEN ISNULL(SFB.PathFile, '') <> '' THEN '" + imgPath + "' + SFB.PathFile ELSE '' END AS PicFile, ITM.Code AS ChkFBCode, ");
                    sbSQL.Append("       REPLACE(SUBSTRING((SELECT CONVERT(varchar, OIDGParts) + ',' AS 'data()' FROM SMPLRequestFabricParts WHERE OIDSMPLFB = SFB.OIDSMPLFB FOR XML PATH('')), 1, LEN((SELECT CONVERT(varchar, OIDGParts) + ',' AS 'data()' FROM SMPLRequestFabricParts WHERE OIDSMPLFB = SFB.OIDSMPLFB FOR XML PATH(''))) -1), ' ', '')  AS FBPartsID, ");
                    sbSQL.Append("       SUBSTRING((SELECT GarmentParts + ',' AS 'data()' FROM(SELECT GP.OIDGParts, GP.GarmentParts FROM SMPLRequestFabricParts AS SFBP INNER JOIN GarmentParts AS GP ON SFBP.OIDGParts = GP.OIDGParts AND SFBP.OIDSMPLFB = SFB.OIDSMPLFB) AS FParts FOR XML PATH('')), 1, LEN((SELECT GarmentParts + ',' AS 'data()' FROM(SELECT GP.OIDGParts, GP.GarmentParts FROM SMPLRequestFabricParts AS SFBP INNER JOIN GarmentParts AS GP ON SFBP.OIDGParts = GP.OIDGParts AND SFBP.OIDSMPLFB = SFB.OIDSMPLFB) AS FParts FOR XML PATH(''))) -1) AS FBPartsName, SFB.OIDSMPLFB AS FBID, SFB.Remark  ");
                    sbSQL.Append("FROM   SMPLRequestFabric AS SFB INNER JOIN ");
                    sbSQL.Append("       SMPLQuantityRequired AS SQR ON SFB.OIDSMPLDT = SQR.OIDSMPLDT INNER JOIN ");
                    sbSQL.Append("       SMPLRequest AS SRQ ON SQR.OIDSMPL = SRQ.OIDSMPL INNER JOIN ");
                    sbSQL.Append("       Items AS ITM ON SFB.OIDITEM = ITM.OIDITEM ");
                    sbSQL.Append("WHERE  (SRQ.SMPLNo = N'" + SMPLNo + "') ");
                    sbSQL.Append("ORDER BY ColorName, FBPartsID ");
                    new ObjDE.setGridControl(gcList_Fabric, gridView4, sbSQL).getData(false, false, false, true);


                    sbSQL.Clear();
                    sbSQL.Append("SELECT DISTINCT ");
                    sbSQL.Append("        MKD.OIDMARKDT AS ID, (CASE MKD.OIDSIZEZONE WHEN 0 THEN 'Japan' WHEN 1 THEN 'Europe' WHEN 2 THEN 'US' END) AS Zone, SRQ.SMPLPatternNo AS PatternNo, MKD.DetailsType, ");
                    sbSQL.Append("        SUBSTRING((SELECT GarmentParts + ', ' AS 'data()' FROM(SELECT GP.OIDGParts, GP.GarmentParts FROM SMPLRequestFabricParts AS SFBP INNER JOIN GarmentParts AS GP ON SFBP.OIDGParts = GP.OIDGParts AND SFBP.OIDSMPLFB IN (SELECT value FROM STRING_SPLIT(MKD.OIDSMPLDTStuff, ','))) AS FParts FOR XML PATH('')), 1, LEN((SELECT GarmentParts + ', ' AS 'data()' FROM(SELECT GP.OIDGParts, GP.GarmentParts FROM SMPLRequestFabricParts AS SFBP INNER JOIN GarmentParts AS GP ON SFBP.OIDGParts = GP.OIDGParts AND SFBP.OIDSMPLFB IN (SELECT value FROM STRING_SPLIT(MKD.OIDSMPLDTStuff, ','))) AS FParts FOR XML PATH(''))) -2) AS FabricParts, ");
                    sbSQL.Append("        PS.SizeName AS Size, FORMAT(MKD.TotalWidthSTD, '###.####') AS StdWidth, FORMAT(MKD.UsableWidth, '###.####') AS UsableWidth, FORMAT(MKD.GM2, '###.####') AS Weight, FORMAT(MKD.PracticalLengthCM, '###.####') AS ActualLength, FORMAT(MKD.QuantityPCS, '###.####') AS Quantity, FORMAT(MKD.LengthPer1CM, '###.####') AS LengthBody ");
                    sbSQL.Append("FROM    Marking AS MK INNER JOIN ");
                    sbSQL.Append("        MarkingDetails AS MKD ON MK.OIDMARK = MKD.OIDMARK INNER JOIN ");
                    sbSQL.Append("        SMPLRequest AS SRQ ON MK.OIDSMPL = SRQ.OIDSMPL INNER JOIN ");
                    sbSQL.Append("        ProductSize AS PS ON MKD.OIDSIZE = PS.OIDSIZE ");
                    sbSQL.Append("WHERE (SRQ.SMPLNo = N'" + SMPLNo + "') ");
                    sbSQL.Append("ORDER BY MKD.OIDMARKDT ");
                    new ObjDE.setGridControl(gridControl5, gridView5, sbSQL).getData(false, false, false, true);


                    //Load Meterial Tab
                    txtSampleID_Mat.Text = lblID.Text;
                    sbSQL.Clear();
                    sbSQL.Append("SELECT DISTINCT SQR.OIDSMPL, SRQ.SMPLPatternNo, SQR.OIDCOLOR AS ColorName, SQR.OIDSIZE AS SizeName, SQR.Quantity, SQR.OIDUnit AS UnitName, '' AS OIDSMPLDT ");
                    sbSQL.Append("FROM   SMPLQuantityRequired AS SQR INNER JOIN ");
                    sbSQL.Append("       SMPLRequest AS SRQ ON SQR.OIDSMPL = SRQ.OIDSMPL ");
                    sbSQL.Append("WHERE (SRQ.SMPLNo = N'" + SMPLNo + "') ");
                    sbSQL.Append("ORDER BY SQR.OIDSMPL, SRQ.SMPLPatternNo ");
                    new ObjDE.setGridControl(gridControl6, gridView6, sbSQL).getData(false, false, true, true);

                    sbSQL.Clear();
                    sbSQL.Append("SELECT SRM.OIDSMPLMT AS MatID, SRM.OIDSMPLDT AS SampleID, SRM.OIDDEPT AS WorkStation, SRM.VendMTCode, SRM.SMPLotNo, SRM.OIDVEND AS Vendor, SRM.MTColor AS MatColor, SQR.OIDCOLOR AS ColorName,  ");
                    sbSQL.Append("       SRM.MTSize AS MatSize, FORMAT(SRM.Consumption, '###.####') AS Consumption, SRM.OIDUNIT AS Unit, SRM.Composition, SRM.Details, FORMAT(SRM.Price, '###.####') AS Price, SRM.OIDCURR AS Currency, SRM.OIDITEM AS NAVCode, ITM.Description, SRM.Situation, SRM.Comment, SRM.Remark, ");
                    sbSQL.Append("       CASE WHEN ISNULL(SRM.PathFile, '') <> '' THEN '" + imgPath + "' + SRM.PathFile ELSE '' END AS PathFile, ITM.Code AS ChkMTCode, SRM.OIDSMPLMT AS MID ");
                    sbSQL.Append("FROM   SMPLRequestMaterial AS SRM INNER JOIN ");
                    sbSQL.Append("       SMPLQuantityRequired AS SQR ON SRM.OIDSMPLDT = SQR.OIDSMPLDT INNER JOIN ");
                    sbSQL.Append("       SMPLRequest AS SRQ ON SQR.OIDSMPL = SRQ.OIDSMPL INNER JOIN ");
                    sbSQL.Append("       Items AS ITM ON SRM.OIDITEM = ITM.OIDITEM ");
                    sbSQL.Append("WHERE (SRQ.SMPLNo = N'" + SMPLNo + "') ");
                    sbSQL.Append("ORDER BY WorkStation, ColorName, MatSize, MatID ");
                    new ObjDE.setGridControl(gridControl8, gridView8, sbSQL).getData(false, false, false, true);


                }
            }

            if (SMPLModify == "READ-ONLY")
            {
                SetReadOnly();
            }
            else
            {
                SetWrite();
            }
        }

        private void LoadDefaultData()
        {
            //Branch
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT Name AS Branch, OIDBranch AS ID FROM Branchs WHERE OIDCOMPANY = (SELECT TOP(1) OIDCOMPANY FROM Company WHERE Code=N'" + COMPANY_CODE + "') ORDER BY OIDBranch");
            new ObjDE.setSearchLookUpEdit(glBranch_Main, sbSQL, "Branch", "ID").getData();
            glBranch_Main.Properties.View.PopulateColumns(glBranch_Main.Properties.DataSource);
            glBranch_Main.Properties.View.Columns["ID"].Visible = false;

            //Season
            sbSQL.Clear();
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
            new ObjDE.setSearchLookUpEdit(glSeason_Main, sbSQL, "Season", "Season").getData();

            //UseFor
            sbSQL.Clear();
            sbSQL.Append("SELECT UseFor, OIDUF AS ID FROM SMPLUseFor ORDER BY OIDUF");
            new ObjDE.setGridLookUpEdit(glUseFor, sbSQL, "UseFor", "ID").getData();
            glUseFor.Properties.View.PopulateColumns(glUseFor.Properties.DataSource);
            glUseFor.Properties.View.Columns["ID"].Visible = false;


            //Customer
            sbSQL.Clear();
            sbSQL.Append("SELECT ShortName, Name, OIDCUST AS ID FROM Customer ORDER BY ShortName");
            new ObjDE.setSearchLookUpEdit(slCustomer_Main, sbSQL, "Name", "ID").getData();
            slCustomer_Main.Properties.View.PopulateColumns(slCustomer_Main.Properties.DataSource);
            slCustomer_Main.Properties.View.Columns["ID"].Visible = false;

            //GarmentCategory
            sbSQL.Clear();
            sbSQL.Append("SELECT CategoryName, OIDGCATEGORY AS ID FROM GarmentCategory ORDER BY CategoryName");
            new ObjDE.setSearchLookUpEdit(glCategoryDivision_Main, sbSQL, "CategoryName", "ID").getData();
            glCategoryDivision_Main.Properties.View.PopulateColumns(glCategoryDivision_Main.Properties.DataSource);
            glCategoryDivision_Main.Properties.View.Columns["ID"].Visible = false;

            //ProductStyle
            sbSQL.Clear();
            sbSQL.Append("SELECT StyleName, OIDSTYLE AS ID FROM ProductStyle ORDER BY StyleName");
            new ObjDE.setSearchLookUpEdit(slStyleName_Main, sbSQL, "StyleName", "ID").getData();
            slStyleName_Main.Properties.View.PopulateColumns(slStyleName_Main.Properties.DataSource);
            slStyleName_Main.Properties.View.Columns["ID"].Visible = false;

            /*Set GridAdd Bind Color and Size*/

            LoadSizeColor();

            sbSQL.Clear();
            sbSQL.Append("SELECT UnitName, OIDUNIT AS ID FROM Unit ORDER BY UnitName");
            new ObjDE.setSearchLookUpEdit(slueUnit, sbSQL, "UnitName", "ID").getData();
            slueUnit.Properties.View.PopulateColumns(slueUnit.Properties.DataSource);
            slueUnit.Properties.View.Columns["ID"].Visible = false;

            slueConsumpUnit.Properties.DataSource = slueUnit.Properties.DataSource;
            slueConsumpUnit.Properties.DisplayMember = "UnitName";
            slueConsumpUnit.Properties.ValueMember = "ID";
            slueConsumpUnit.Properties.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
            slueConsumpUnit.Properties.View.PopulateColumns(slueConsumpUnit.Properties.DataSource);
            slueConsumpUnit.Properties.View.Columns["ID"].Visible = false;

            rep_slueUnit_FB.DataSource = slueUnit.Properties.DataSource;
            rep_slueUnit_FB.DisplayMember = "UnitName";
            rep_slueUnit_FB.ValueMember = "ID";
            rep_slueUnit_FB.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
            rep_slueUnit_FB.View.PopulateColumns(rep_slueUnit_FB.DataSource);
            rep_slueUnit_FB.View.Columns["ID"].Visible = false;

            rep_slUnit_Material.DataSource = slueUnit.Properties.DataSource;
            rep_slUnit_Material.DisplayMember = "UnitName";
            rep_slUnit_Material.ValueMember = "ID";
            rep_slUnit_Material.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
            rep_slUnit_Material.View.PopulateColumns(rep_slUnit_Material.DataSource);
            rep_slUnit_Material.View.Columns["ID"].Visible = false;

            rep_MtrUnit.DataSource = slueUnit.Properties.DataSource;
            rep_MtrUnit.DisplayMember = "UnitName";
            rep_MtrUnit.ValueMember = "ID";
            rep_MtrUnit.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
            rep_MtrUnit.View.PopulateColumns(rep_MtrUnit.DataSource);
            rep_MtrUnit.View.Columns["ID"].Visible = false;

            // Tab : Fabric Load
            sbSQL.Clear();
            sbSQL.Append("SELECT Code AS VendorCode, Name AS VendorName, OIDVEND AS ID FROM Vendor ORDER BY VendorType");
            new ObjDE.setSearchLookUpEdit(slVendor_FB, sbSQL, "VendorName", "ID").getData();
            slVendor_FB.Properties.View.PopulateColumns(slVendor_FB.Properties.DataSource);
            slVendor_FB.Properties.View.Columns["ID"].Visible = false;


            rep_slueSupplier.DataSource = slVendor_FB.Properties.DataSource;
            rep_slueSupplier.DisplayMember = "VendorName";
            rep_slueSupplier.ValueMember = "ID";
            rep_slueSupplier.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
            rep_slueSupplier.View.PopulateColumns(rep_slueSupplier.DataSource);
            rep_slueSupplier.View.Columns["ID"].Visible = false;

            sbSQL.Clear();
            sbSQL.Append("SELECT ColorName, OIDCOLOR AS ID FROM ProductColor WHERE (ColorType = 1) ORDER BY ColorName");
            new ObjDE.setSearchLookUpEdit(slFBColor_FB, sbSQL, "ColorName", "ID").getData();
            slFBColor_FB.Properties.View.PopulateColumns(slFBColor_FB.Properties.DataSource);
            slFBColor_FB.Properties.View.Columns["ID"].Visible = false;

            rep_slueFBColor.DataSource = slFBColor_FB.Properties.DataSource;
            rep_slueFBColor.DisplayMember = "ColorName";
            rep_slueFBColor.ValueMember = "ID";
            rep_slueFBColor.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
            rep_slueFBColor.View.PopulateColumns(rep_slueFBColor.DataSource);
            rep_slueFBColor.View.Columns["ID"].Visible = false;

            sbSQL.Clear();
            sbSQL.Append("SELECT Code, Description, Type, ID ");
            sbSQL.Append("FROM (");
            sbSQL.Append("  SELECT Code, Description, 'Fabric' AS Type, MaterialType, OIDITEM AS ID FROM Items WHERE (MaterialType = '" + TYPE_FABRIC + "') ");
            sbSQL.Append("  UNION ALL ");
            sbSQL.Append("  SELECT Code, Description, 'Temporary' AS Type, MaterialType, OIDITEM AS ID FROM Items WHERE (MaterialType = '" + TYPE_TEMPORARY + "') AND (Code LIKE 'TMPFB%') ");
            sbSQL.Append(") AS FBCode ");
            sbSQL.Append("ORDER BY MaterialType, Code ");
            new ObjDE.setSearchLookUpEdit(slFBCode_FB, sbSQL, "Code", "ID").getData();
            slFBCode_FB.Properties.View.PopulateColumns(slFBCode_FB.Properties.DataSource);
            slFBCode_FB.Properties.View.Columns["ID"].Visible = false;

            rep_slueFBCode.DataSource = slFBCode_FB.Properties.DataSource;
            rep_slueFBCode.DisplayMember = "Code";
            rep_slueFBCode.ValueMember = "ID";
            rep_slueFBCode.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
            rep_slueFBCode.View.PopulateColumns(rep_slueFBCode.DataSource);
            rep_slueFBCode.View.Columns["ID"].Visible = false;

            sbSQL.Clear();
            sbSQL.Append("SELECT Currency, OIDCURR AS ID FROM Currency ORDER BY OIDCURR");
            new ObjDE.setSearchLookUpEdit(glCurrency_FB, sbSQL, "Currency", "ID").getData();
            glCurrency_FB.Properties.View.PopulateColumns(glCurrency_FB.Properties.DataSource);
            glCurrency_FB.Properties.View.Columns["ID"].Visible = false;

            rep_slueCurrency.DataSource = glCurrency_FB.Properties.DataSource;
            rep_slueCurrency.DisplayMember = "Currency";
            rep_slueCurrency.ValueMember = "ID";
            rep_slueCurrency.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
            rep_slueCurrency.View.PopulateColumns(rep_slueCurrency.DataSource);
            rep_slueCurrency.View.Columns["ID"].Visible = false;


            sbSQL.Clear();
            sbSQL.Append("Select OIDGParts AS ID, GarmentParts From GarmentParts ORDER BY GarmentParts");
            new ObjDE.setGridControl(gcPart_Fabric, gridView11, sbSQL).getData(false, false, true);

            // Tab : Material Load
            sbSQL.Clear();
            sbSQL.Append("SELECT Code, Description, Type, ID ");
            sbSQL.Append("FROM (");
            sbSQL.Append("  SELECT Code, Description, 'Meterial' AS Type, MaterialType, OIDITEM AS ID FROM Items WHERE (MaterialType IN ('" + TYPE_ACCESSORY + "', '" + TYPE_PACKAGING + "')) ");
            sbSQL.Append("  UNION ALL ");
            sbSQL.Append("  SELECT Code, Description, 'Temporary' AS Type, MaterialType, OIDITEM AS ID FROM Items WHERE (MaterialType = '" + TYPE_TEMPORARY + "') AND (Code LIKE 'TMPMT%') ");
            sbSQL.Append(") AS FBCode ");
            sbSQL.Append("ORDER BY MaterialType, Code ");
            new ObjDE.setSearchLookUpEdit(slMatCode_Mat, sbSQL, "Code", "ID").getData();
            slMatCode_Mat.Properties.View.PopulateColumns(slMatCode_Mat.Properties.DataSource);
            slMatCode_Mat.Properties.View.Columns["ID"].Visible = false;

            rep_MtrItem.DataSource = slMatCode_Mat.Properties.DataSource;
            rep_MtrItem.DisplayMember = "Code";
            rep_MtrItem.ValueMember = "ID";
            rep_MtrItem.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
            rep_MtrItem.View.PopulateColumns(rep_MtrItem.DataSource);
            rep_MtrItem.View.Columns["ID"].Visible = false;

            sbSQL.Clear();
            sbSQL.Append("SELECT Code AS VendorCode, Name AS VendorName, OIDVEND AS ID FROM Vendor ORDER BY VendorType");
            new ObjDE.setSearchLookUpEdit(slVendor_Mat, sbSQL, "VendorCode", "ID").getData();
            slVendor_Mat.Properties.View.PopulateColumns(slVendor_Mat.Properties.DataSource);
            slVendor_Mat.Properties.View.Columns["ID"].Visible = false;

            rep_MtrVendor.DataSource = slVendor_Mat.Properties.DataSource;
            rep_MtrVendor.DisplayMember = "VendorCode";
            rep_MtrVendor.ValueMember = "ID";
            rep_MtrVendor.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
            rep_MtrVendor.View.PopulateColumns(rep_MtrVendor.DataSource);
            rep_MtrVendor.View.Columns["ID"].Visible = false;

            sbSQL.Clear();
            sbSQL.Append("SELECT ColorName, OIDCOLOR AS ID FROM ProductColor WHERE (ColorType IN (2, 3)) ORDER BY ColorName");
            new ObjDE.setSearchLookUpEdit(slMatColor_Mat, sbSQL, "ColorName", "ID").getData();
            slMatColor_Mat.Properties.View.PopulateColumns(slMatColor_Mat.Properties.DataSource);
            slMatColor_Mat.Properties.View.Columns["ID"].Visible = false;

            rep_MtrColor.DataSource = slMatColor_Mat.Properties.DataSource;
            rep_MtrColor.DisplayMember = "ColorName";
            rep_MtrColor.ValueMember = "ID";
            rep_MtrColor.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
            rep_MtrColor.View.PopulateColumns(rep_MtrColor.DataSource);
            rep_MtrColor.View.Columns["ID"].Visible = false;

            sbSQL.Clear();
            sbSQL.Append("SELECT Currency, OIDCURR AS ID FROM Currency ORDER BY OIDCURR");
            new ObjDE.setSearchLookUpEdit(glCurrency_Mat, sbSQL, "Currency", "ID").getData();
            glCurrency_Mat.Properties.View.PopulateColumns(glCurrency_Mat.Properties.DataSource);
            glCurrency_Mat.Properties.View.Columns["ID"].Visible = false;

            rep_MtrCurrency.DataSource = glCurrency_Mat.Properties.DataSource;
            rep_MtrCurrency.DisplayMember = "Currency";
            rep_MtrCurrency.ValueMember = "ID";
            rep_MtrCurrency.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
            rep_MtrCurrency.View.PopulateColumns(rep_MtrCurrency.DataSource);
            rep_MtrCurrency.View.Columns["ID"].Visible = false;

        }

        public void getGrid_SMPL(GridControl glName, DevExpress.XtraGrid.Views.Grid.GridView gvName, int OIDUser = 0, int showDoc = 1, int showUser = 0)
        {
            StringBuilder sql = new StringBuilder();
            sql.Append("SELECT smpl.OIDSMPL AS ID, smpl.SMPLNo AS [SMPL No.], smpl.Status, (CASE smpl.Status WHEN 0 THEN 'New' WHEN 1 THEN 'Wait Approved' WHEN 2 THEN 'Customer Approved' END) AS [Status Name],  ");
            sql.Append("       (CASE WHEN smpl.SMPLRevise = 0 THEN '' ELSE CONVERT(VARCHAR, smpl.SMPLRevise) END) AS [SMPL Revise], smpl.OIDBranch, b.Name AS Branch, smpl.OIDDEPT, d.Name AS[Sales Section], (CASE WHEN ISNULL(smpl.RequestDate, '') = '' THEN '' ELSE CONVERT(VARCHAR(10), smpl.RequestDate, 103) END) AS RequestDate, smpl.SpecificationSize, (CASE smpl.SpecificationSize WHEN 0 THEN 'Necessary' WHEN 1 THEN 'Unnecessary' END) AS [Specification Size], smpl.Season, smpl.OIDCUST, c.Name AS Customer, ");
            sql.Append("       smpl.UseFor, ISNULL(UF.UseFor, '') AS [Use For], ");
            sql.Append("       smpl.OIDCATEGORY, g.CategoryName AS Category, smpl.OIDSTYLE, p.StyleName AS Style, smpl.SMPLItem AS[SMPL Item], smpl.SMPLPatternNo AS[SMPL Pattern No.], smpl.PatternSizeZone, (CASE smpl.PatternSizeZone WHEN 0 THEN 'Japan' WHEN 1 THEN 'Europe' WHEN 2 THEN 'US' END) AS [Pattern Size Zone], ");
            sql.Append("       smpl.CustApproved AS[Customer Approved], (CASE smpl.CustApproved WHEN 0 THEN '-' WHEN 1 THEN 'Yes' END) AS CustomerApprovedStatus, (CASE WHEN ISNULL(smpl.CustApprovedDate, '') = '' THEN '' ELSE CONVERT(VARCHAR(10), smpl.CustApprovedDate, 103) END) AS CustomerApprovedDate, smpl.ReferenceNo, smpl.ContactName, ");
            sql.Append("       (CASE WHEN ISNULL(smpl.DeliveryRequest, '') = '' THEN '' ELSE CONVERT(VARCHAR(10), smpl.DeliveryRequest, 103) END) AS DeliveryRequest, smpl.ModelName, smpl.Situation, smpl.StateArrangements, smpl.ACPurRecBy, (CASE smpl.ACPurRecBy WHEN 0 THEN '-' WHEN 1 THEN 'Yes' END) AS[Accessory Purchase Received], ");
            sql.Append("       (CASE WHEN ISNULL(smpl.ACPurRecDate, '') = '' THEN '' ELSE CONVERT(VARCHAR(10), smpl.ACPurRecDate, 103) END) AS [Accessory Purchase Received Date], smpl.FBPurRecBy, (CASE smpl.FBPurRecBy WHEN 0 THEN '-' WHEN 1 THEN 'Yes' END) AS[Fabric Purchase Received],  ");
            sql.Append("       (CASE WHEN ISNULL(smpl.FBPurRecDate, '') = '' THEN '' ELSE CONVERT(VARCHAR(10), smpl.FBPurRecDate, 103) END) AS [Fabric Purchase Received Date], ISNULL(smpl.PictureFile, '') AS PictureFile, smpl.CreatedBy AS ByCreated, smpl.CreatedDate AS DateCreated, smpl.UpdatedBy, smpl.UpdatedDate, ");
            sql.Append("       ISNULL((SELECT TOP(1) ITM.Code FROM SMPLRequestFabric AS SRFB INNER JOIN Items AS ITM ON SRFB.OIDITEM = ITM.OIDITEM WHERE(SRFB.OIDSMPLDT IN (SELECT xSQR.OIDSMPLDT FROM SMPLRequest AS xSRQ INNER JOIN SMPLQuantityRequired AS xSQR ON xSRQ.OIDSMPL = xSQR.OIDSMPL WHERE(xSRQ.OIDSMPL = smpl.OIDSMPL))) AND (ITM.MaterialType = '8') AND (ITM.Code LIKE 'TMPFB%')), '') AS ChkFBCode, ");
            sql.Append("       ISNULL((SELECT TOP(1) ITM.Code FROM SMPLRequestMaterial AS SRMT INNER JOIN Items AS ITM ON SRMT.OIDITEM = ITM.OIDITEM WHERE(SRMT.OIDSMPLDT IN (SELECT xSQR.OIDSMPLDT FROM SMPLRequest AS xSRQ INNER JOIN SMPLQuantityRequired AS xSQR ON xSRQ.OIDSMPL = xSQR.OIDSMPL WHERE(xSRQ.OIDSMPL = smpl.OIDSMPL))) AND (ITM.MaterialType = '8') AND (ITM.Code LIKE 'TMPMT%')), '') AS ChkMTCode, u.FullName AS CreatedBy, smpl.CreatedDate AS CreatedDate, smpl.SMPLStatus  ");
            sql.Append("FROM   SMPLRequest AS smpl LEFT OUTER JOIN ");
            sql.Append("       SMPLUseFor AS UF ON smpl.UseFor = UF.OIDUF LEFT OUTER JOIN ");
            sql.Append("       Branchs AS b ON b.OIDBranch = smpl.OIDBranch LEFT OUTER JOIN ");
            sql.Append("       Departments AS d ON d.OIDDEPT = smpl.OIDDEPT LEFT OUTER JOIN ");
            sql.Append("       Customer AS c ON c.OIDCUST = smpl.OIDCUST LEFT OUTER JOIN ");
            sql.Append("       GarmentCategory AS g ON g.OIDGCATEGORY = smpl.OIDCATEGORY LEFT OUTER JOIN ");
            sql.Append("       ProductStyle AS p ON p.OIDSTYLE = smpl.OIDSTYLE LEFT OUTER JOIN ");
            sql.Append("       Users AS u ON smpl.CreatedBy = u.OIDUSER ");
            sql.Append("WHERE  (smpl.SMPLNo <> N'') ");
            if (showDoc == 1)
                sql.Append("AND  (smpl.SMPLStatus = 1) ");
            if (showUser == 0)
                sql.Append("AND  (smpl.CreatedBy = '" + OIDUser + "') ");
            sql.Append("ORDER BY smpl.CreatedDate DESC ");
            new ObjDE.setGridControl(glName, gvName, sql).getData(false, false, false, true);
            //getGc(sql, glName,MDS());

            gvName.Columns[0].Visible = false; //OIDSMPL
            gvName.Columns[2].Visible = false; //Status
            gvName.Columns[5].Visible = false; //OIDBranch
            gvName.Columns[7].Visible = false; //OIDDEPT
            gvName.Columns[10].Visible = false; //SpecificationSize
            gvName.Columns[13].Visible = false; //OIDCUST
            gvName.Columns[15].Visible = false; //UseFor
            gvName.Columns[17].Visible = false; //OIDCATEGORY
            gvName.Columns[19].Visible = false; //OIDSTYLE
            gvName.Columns[23].Visible = false; //PatternSizeZone
            gvName.Columns[25].Visible = false; //CustApproved
            gvName.Columns[34].Visible = false; //ACPurRecBy
            gvName.Columns[37].Visible = false; //FBPurRecBy
            gvName.Columns[40].Visible = false; //PictureFile
            gvName.Columns[41].Visible = false; //ByCreate
            gvName.Columns[42].Visible = false; //CreateDate
            gvName.Columns[43].Visible = false; //UpdateBy
            gvName.Columns[44].Visible = false; //UpdateDate
            gvName.Columns[45].Visible = false; //ChkFBCode
            gvName.Columns[46].Visible = false; //UpdateDate
            gvName.Columns[49].Visible = false; //ChkMTCode

            gvName.Columns[28].VisibleIndex = 4;

            gvName.Columns["SMPL Revise"].Width = 60;
            gvName.Columns["Sales Section"].Width = 80;
            gvName.Columns["RequestDate"].Width = 100;
            gvName.Columns["Specification Size"].Width = 100;
            gvName.Columns["Pattern Size Zone"].Width = 100;
            gvName.Columns["CustomerApprovedStatus"].Width = 120;
            gvName.Columns["CustomerApprovedDate"].Width = 110;
            gvName.Columns["DeliveryRequest"].Width = 100;
            gvName.Columns["DeliveryRequest"].Width = 100;
            gvName.Columns["Accessory Purchase Received"].Width = 130;
            gvName.Columns["Accessory Purchase Received Date"].Width = 130;
            gvName.Columns["Fabric Purchase Received"].Width = 110;
            gvName.Columns["Fabric Purchase Received Date"].Width = 120;

            gvName.Columns["Status Name"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gvName.Columns["SMPL Revise"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gvName.Columns["Sales Section"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gvName.Columns["RequestDate"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gvName.Columns["Season"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gvName.Columns["Pattern Size Zone"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gvName.Columns["CustomerApprovedStatus"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gvName.Columns["CustomerApprovedDate"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gvName.Columns["DeliveryRequest"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gvName.Columns["Accessory Purchase Received"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gvName.Columns["Accessory Purchase Received Date"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gvName.Columns["Fabric Purchase Received"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gvName.Columns["Fabric Purchase Received Date"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;

            gvName.Columns[0].Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
            gvName.Columns[1].Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
            gvName.Columns[2].Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
            gvName.Columns[3].Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
            gvName.Columns[4].Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;

        }

        private void LoadNewData(string Status = "NEW")
        {
            //MessageBox.Show(ObjDE.GlobalVar.DBC.getCONNECTION_STRING());
            SetWrite();

            if (rgDocActive.EditValue == null)
                rgDocActive.EditValue = 1;
            if (rgDocUser.EditValue == null)
                rgDocUser.EditValue = 0;

            radioGroup4.EditValue = 1;

            radioGroup4.EditValue = 0;

            //Tab : Main Load
            getGrid_SMPL(gridControl1, gridView1, UserLogin.OIDUser, Convert.ToInt32(rgDocActive.EditValue.ToString()), Convert.ToInt32(rgDocUser.EditValue.ToString()));
            HideSelectDoc();

            txtSMPLNo.Text = "";
            lblID.Text = "";

            if (Status == "NEW" || Status == "")
                lblStatus.Text = "New SMPL";
            else if (Status == "UPDATE")
                lblStatus.Text = "Update SMPL";
            else if (Status == "REVISE")
                lblStatus.Text = "Revise SMPL";
            else if (Status == "CLONE")
                lblStatus.Text = "Clone SMPL";
            else if (Status == "READ-ONLY")
                lblStatus.Text = "Read Only SMPL";

            radioGroup4.EditValue = "1";
            radioGroup4.EditValue = "0";

            txeQtyDF.Text = "0";

            glSaleSection_Main.EditValue = "";
            txtReferenceNo_Main.Text = "";
            txtContactName_Main.Text = "";
            txtSMPLItemNo_Main.Text = "";
            txtModelName_Main.Text = "";
            txtSMPLPatternNo_Main.Text = "";
            txtPictureFile_Main.Text = "";
            picMain.Image = null;
            txtSituation_Main.Text = "";
            txtStateArrangments_Main.Text = "";

            btnGenSMPLNo.Enabled = false;

            glUseFor.EditValue = "";

            radioGroup1.EditValue = 0;
            radioGroup3.EditValue = 0;
            radioGroup4.EditValue = 0;
            radioGroup5.EditValue = 1;
            radioGroup6.EditValue = 1;
            dtRequestDate_Main.EditValue = DateTime.Now;

            dtDeliveryRequest_Main.EditValue = null;
            dtCustomerApproved_Main.EditValue = null;
            dtACPRBy_Main.EditValue = null;
            dtFBPRBy_Main.EditValue = null;

            glueCreateBy_Main.EditValue = UserLogin.OIDUser;
            txtCreateDate_Main.EditValue = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            glueUpdateBy_Main.EditValue = UserLogin.OIDUser;
            txtUpdateDate_Main.EditValue = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

            lblDescription.Text = "-";
            lblDescription.AppearanceItemCaption.BackColor = Color.Empty;
            lblRowID.Text = "";

            if (dtQtyRequired.Rows.Count > 0)
                dtQtyRequired.Rows.Clear();
            gcQtyRequired.DataSource = dtQtyRequired;

            glBranch_Main.EditValue = "";
            glSeason_Main.EditValue = "";
            slCustomer_Main.EditValue = "";
            glCategoryDivision_Main.EditValue = "";
            slStyleName_Main.EditValue = "";
            slueUnit.EditValue = "";
            slVendor_FB.EditValue = "";
            slFBColor_FB.EditValue = "";
            slFBCode_FB.EditValue = "";
            slFGColor_FB.EditValue = "";
            glCurrency_FB.EditValue = "";


            slueUnit.EditValue = 15;
            slueConsumpUnit.EditValue = 15;


            txtSampleID_FB.Text = "";
            txtFabricRacordID_FB.Text = "";
            slFGColor_FB.ReadOnly = false;
            slFGColor_FB.EditValue = "";
            slFGColor_FB.ReadOnly = true;
            glCurrency_FB.EditValue = "";

            ClearFabric();
            if (dtFBSample.Rows.Count > 0)
                dtFBSample.Rows.Clear();
            gridControl3.DataSource = dtFBSample;
            if (dtFBSize.Rows.Count > 0)
                dtFBSize.Rows.Clear();
            gcSize_Fabric.DataSource = dtFBSize;
            if (dtFabric.Rows.Count > 0)
                dtFabric.Rows.Clear();
            gcList_Fabric.DataSource = dtFabric;

            gridControl5.DataSource = null;

            txtImgUpload_FB.EditValue = "";
            picUpload_FB.Image = null;

            gridView11.ClearSelection();

            // Tab : Materials
            ClearMaterial();

            txtSampleID_Mat.Text = "";
            txtMatRecordID_Mat.Text = "";
            txeMatDescription.Text = "";
            txtVendName_Mat.Text = "";

            if (dtMaterial.Rows.Count > 0)
                dtMaterial.Rows.Clear();
            gridControl6.DataSource = dtMaterial;

            gcList_Fabric.DataSource = null;
            gridControl8.DataSource = null;

        }

        private void bbiSave_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            tabbedControlGroup1.SelectedTabPage = layoutControlGroup2; //Tab Main
            saveSampleRequest();
            //switch (currenTab)
            //{
            //    /* List of , Main , Fabric , Material */
            //    case "Main"     : saveMain(); break;
            //    case "Fabric"   : saveFabric(); break;
            //    case "Material" : saveMaterials(); break;
            //    default         : break;
            //}

        }

        private void gvGarment_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {
            //
        }

        private void selectStatus(int value)
        {
            //switch (value)
            //{
            //    case 0:
            //        rgStatus.SelectedIndex = 0;
            //        break;
            //    case 1:
            //        rgStatus.SelectedIndex = 1;
            //        break;
            //    default:
            //        rgStatus.SelectedIndex = -1;
            //        break;
            //}
        }

        private bool chkDuplicate()
        {
            bool chkDup = true;
            //if (txeName.Text != "")
            //{
            //    txeName.Text = txeName.Text.Trim();
            //    if (txeName.Text.Trim() != "" && lblStatus.Text == "* Add Payment Term")
            //    {
            //        StringBuilder sbSQL = new StringBuilder();
            //        sbSQL.Append("SELECT TOP(1) Name FROM PaymentTerm WHERE (Name = N'" + txeName.Text.Trim() + "') ");
            //        if (DBC.DBQuery(sbSQL).getString() != "")
            //        {
            //            FUNC.msgWarning("Duplicate payment term. !! Please Change.");
            //            txeName.Text = "";
            //            chkDup = false;
            //        }
            //    }
            //    else if (txeName.Text.Trim() != "" && lblStatus.Text == "* Edit Payment Term")
            //    {
            //        StringBuilder sbSQL = new StringBuilder();
            //        sbSQL.Append("SELECT TOP(1) OIDPayment ");
            //        sbSQL.Append("FROM PaymentTerm ");
            //        sbSQL.Append("WHERE (Name = N'" + txeName.Text.Trim() + "') ");
            //        string strCHK = DBC.DBQuery(sbSQL).getString();
            //        if (strCHK != "" && strCHK != txeID.Text.Trim())
            //        {
            //            FUNC.msgWarning("Duplicate payment term. !! Please Change.");
            //            txeName.Text = "";
            //            chkDup = false;
            //        }
            //    }
            //}
            return chkDup;
        }

        private void txeName_KeyDown(object sender, KeyEventArgs e)
        {
            //if (e.KeyCode == Keys.Enter)
            //{
            //    txeDescription.Focus();
            //}
        }

        private void txeName_LostFocus(object sender, EventArgs e)
        {
            //txeName.Text = txeName.Text.ToUpper().Trim();
            //bool chkDup = chkDuplicate();
            //if (chkDup == false)
            //{
            //    txeName.Text = "";
            //    txeName.Focus();
            //}
            //else
            //{
            //    txeDescription.Focus();
            //}
        }

        private void txeDescription_KeyDown(object sender, KeyEventArgs e)
        {
        //    if (e.KeyCode == Keys.Enter)
        //    {
        //        txeDueDate.Focus();
        //    }
        }

        private void txeDueDate_KeyDown(object sender, KeyEventArgs e)
        {
            //if (e.KeyCode == Keys.Enter)
            //{
            //    rgStatus.Focus();
            //}
        }

        private void bbiExcel_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            ////if (gridView1.RowCount > 0)
            ////{
            ////    string pathFile = new ObjSet.Folder(@"C:\MDS\Export\").GetPath() + "PaymentTermList_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
            ////    gridView1.ExportToXlsx(pathFile);
            ////    Process.Start(pathFile);
            ////}

            //// Check Part is Exist
            //string root = @"C:/__MDS/Export/";
            //if (!Directory.Exists(root))
            //{
            //    Directory.CreateDirectory(root);
            //}

            //if (tabbedControlGroup1.SelectedTabPageIndex == 0)
            //{
            //    //List of Sample
            //    if (gridView1.RowCount > 0)
            //    {
            //        if (FUNCT.msgQuiz("Export Excel ?") == true)
            //        {
            //            string filePath = root + "SMPL-"+DateTime.Now.ToString("yyyyMMdd-HHmmss") + ".xlsx";
            //            gridView1.ExportToXlsx(filePath);
            //            Process.Start(filePath);
            //        }
            //    }
            //}
            //if (tabbedControlGroup1.SelectedTabPageIndex == 1)
            //{
            //    //Main
            //    if (gvQtyRequired.RowCount > 0)
            //    {
            //        if (FUNCT.msgQuiz("Export Excel ?") == true)
            //        {
            //            string filePath = root + "SMPLMain-" + DateTime.Now.ToString("yyyyMMdd-HHmmss") + ".xlsx";
            //            gvQtyRequired.ExportToXlsx(filePath);
            //            Process.Start(filePath);
            //        }
            //    }
            //}
            //if (tabbedControlGroup1.SelectedTabPageIndex == 2)
            //{
            //    //Fabric
            //    if (gridView4.RowCount > 0)
            //    {
            //        if (FUNCT.msgQuiz("Export Excel ?") == true)
            //        {
            //            string filePath = root + "SMPLFB-" + DateTime.Now.ToString("yyyyMMdd-HHmmss") + ".xlsx";
            //            gridView4.ExportToXlsx(filePath);
            //            Process.Start(filePath);
            //        }
            //    }
            //}
            //if (tabbedControlGroup1.SelectedTabPageIndex == 3)
            //{
            //    //Material
            //    if (gridView8.RowCount > 0)
            //    {
            //        if (FUNCT.msgQuiz("Export Excel ?") == true)
            //        {
            //            string filePath = root + "SMPLMaterial-" + DateTime.Now.ToString("yyyyMMdd-HHmmss") + ".xlsx";
            //            gridView8.ExportToXlsx(filePath);
            //            Process.Start(filePath);
            //        }
            //    }
            //}
        }

        private void gvPTerm_RowClick(object sender, RowClickEventArgs e)
        {
            //lblStatus.Text = "* Edit Payment Term";
            //lblStatus.ForeColor = Color.Red;

            //txeID.Text = gvPTerm.GetFocusedRowCellValue("No").ToString();
            //txeName.Text = gvPTerm.GetFocusedRowCellValue("Name").ToString();
            //txeDescription.Text = gvPTerm.GetFocusedRowCellValue("Description").ToString();
            //txeDueDate.Text = gvPTerm.GetFocusedRowCellValue("DuedateCalculation").ToString();

            //int status = -1;
            //if (gvPTerm.GetFocusedRowCellValue("Status").ToString() != "")
            //{
            //    status = Convert.ToInt32(gvPTerm.GetFocusedRowCellValue("Status").ToString());
            //}

            //selectStatus(status);

            //txeCREATE.Text = gvPTerm.GetFocusedRowCellValue("CreatedBy").ToString();
            //txeDATE.Text = gvPTerm.GetFocusedRowCellValue("CreatedDate").ToString();
        }

        private void bbiPrintPreview_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            //gcPTerm.ShowPrintPreview();
            if (tabbedControlGroup1.SelectedTabPageIndex == 0)
            {
                gridView1.ShowPrintPreview();
            }
            if (tabbedControlGroup1.SelectedTabPageIndex == 1)
            {
                gvQtyRequired.ShowPrintPreview();
            }
            if (tabbedControlGroup1.SelectedTabPageIndex == 2)
            {
                gridView4.ShowPrintPreview();
            }
            if (tabbedControlGroup1.SelectedTabPageIndex == 3)
            {
                gridView8.ShowPrintPreview();
            }
        }

        private void bbiPrint_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            int[] selectedRowHandles = gridView1.GetSelectedRows();
            if (selectedRowHandles.Length > 0)
            {
                gridView1.FocusedRowHandle = selectedRowHandles[0];
                string SMPLNo = gridView1.GetRowCellDisplayText(selectedRowHandles[0], "SMPL No.");
                if (FUNCT.msgQuiz("Confirm print sample request (excel file)  : " + SMPLNo + " ?") == true)
                {
                    layoutControlItem120.Text = "Print excel file processing ..";
                    layoutControlItem120.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                    pbcEXPORT.Properties.Step = 1;
                    pbcEXPORT.Properties.PercentView = true;
                    pbcEXPORT.Properties.Maximum = 11;
                    pbcEXPORT.Properties.Minimum = 0;

                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT TOP(1) SRQ.OIDSMPL, SRQ.SMPLNo, SRQ.SMPLRevise, CPN.EngName AS Company, BCH.Name AS Branch, SRQ.ContactName, SRQ.RequestDate, ");
                    sbSQL.Append("       SRQ.Season, CUS.Name AS Customer, GC.CategoryName AS Category, SRQ.ModelName, SRQ.SMPLPatternNo, SRQ.SMPLItem, SRQ.StateArrangements, ");
                    sbSQL.Append("       CASE WHEN SRQ.PatternSizeZone = 0 THEN 'Japan' ELSE CASE WHEN SRQ.PatternSizeZone = 1 THEN 'Europe' ELSE CASE WHEN SRQ.PatternSizeZone = 2 THEN 'US' ELSE '' END END END AS SizeZone, ");
                    sbSQL.Append("       SRQ.PictureFile, SRQ.DeliveryRequest, UF.UseFor, CASE WHEN SRQ.SpecificationSize = 0 THEN 'Necessary' ELSE 'Unnecessary' END AS SpecificationSize, ");
                    sbSQL.Append("       (SELECT TOP (1) UN.UnitName AS Unit FROM SMPLQuantityRequired AS SQR INNER JOIN Unit AS UN ON SQR.OIDUnit = UN.OIDUNIT WHERE (SQR.OIDSMPL = SRQ.OIDSMPL)) AS Unit, ");
                    sbSQL.Append("       (SELECT CountColor + 'color ' + CountSize + 'size' AS TTCS FROM(SELECT TOP(1)(SELECT CONVERT(varchar, COUNT(OIDCOLOR)) AS Color FROM(SELECT OIDCOLOR FROM SMPLQuantityRequired AS B WHERE(OIDSMPL = A.OIDSMPL) GROUP BY OIDCOLOR) AS CColor) AS CountColor, (SELECT CONVERT(varchar, COUNT(OIDSIZE)) AS Size FROM(SELECT OIDSIZE FROM SMPLQuantityRequired AS C WHERE(OIDSMPL = A.OIDSMPL) GROUP BY OIDSIZE) AS CSize) AS CountSize FROM SMPLQuantityRequired AS A WHERE(OIDSMPL = SRQ.OIDSMPL)) AS A1) AS TTCS,   ");
                    //sbSQL.Append("       (SELECT TOP (1) Quantity FROM SMPLQuantityRequired WHERE (OIDSMPL = SRQ.OIDSMPL)) AS Pcs, ");
                    sbSQL.Append("       ISNULL((SELECT Quantity + '-' AS 'data()' FROM(SELECT CONVERT(VARCHAR, Quantity) AS Quantity FROM SMPLQuantityRequired WHERE OIDSMPL = SRQ.OIDSMPL) AS XA FOR XML PATH('')), '')  AS Pcs, ");
                    sbSQL.Append("       (SELECT SUM(Quantity) AS TTPcs FROM SMPLQuantityRequired WHERE (OIDSMPL = SRQ.OIDSMPL)) AS TTPcs, SRQ.ReferenceNo, SRQ.Situation, U.FullName ");
                    sbSQL.Append("FROM   SMPLRequest AS SRQ INNER JOIN ");
                    sbSQL.Append("        Branchs AS BCH ON SRQ.OIDBranch = BCH.OIDBranch INNER JOIN ");
                    sbSQL.Append("        Company AS CPN ON BCH.OIDCOMPANY = CPN.OIDCOMPANY INNER JOIN ");
                    sbSQL.Append("        SMPLUseFor AS UF ON SRQ.UseFor = UF.OIDUF LEFT OUTER JOIN ");
                    sbSQL.Append("        Customer AS CUS ON SRQ.OIDCUST = CUS.OIDCUST LEFT OUTER JOIN ");
                    sbSQL.Append("        GarmentCategory AS GC ON SRQ.OIDCATEGORY = GC.OIDGCATEGORY LEFT OUTER JOIN ");
                    sbSQL.Append("        Users AS U ON SRQ.UpdatedBy = U.OIDUSER ");
                    sbSQL.Append("WHERE (SRQ.SMPLNo = N'" + SMPLNo + "') ");
                    string[] SMPL = DBC.DBQuery(sbSQL.ToString()).getMultipleValue();
                    if (SMPL.Length > 0)
                    {
                        //****** BEGIN EXPORT *******

                        String sFilePath = System.IO.Path.Combine(new ObjSet.Folder(@"C:\MDS\Export\").GetPath() + SMPLNo + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");
                        if (File.Exists(sFilePath)) { File.Delete(sFilePath); }
                        bool chkExcel = false;
                        Microsoft.Office.Interop.Excel.Application objApp = new Microsoft.Office.Interop.Excel.Application();
                        Microsoft.Office.Interop.Excel.Worksheet objSheet = new Microsoft.Office.Interop.Excel.Worksheet();
                        Microsoft.Office.Interop.Excel.Workbook objWorkBook = null;
                        //object missing = System.Reflection.Missing.Value;

                        pbcEXPORT.PerformStep();
                        pbcEXPORT.Update();

                        try
                        {
                            int blankCol = 0;
                            objWorkBook = objApp.Workbooks.Add(Type.Missing);
                            objWorkBook = objApp.Workbooks.Open(this.reportPath + "SMPL.xlsx");

                            //objSheet = (Microsoft.Office.Interop.Excel.Worksheet)objWorkBook.ActiveSheet;
                            objSheet = (Microsoft.Office.Interop.Excel.Worksheet)objWorkBook.Sheets[1];
                            objSheet.Name = SMPLNo;

                            objSheet.Cells[2, 1] = SMPL[23].ToUpper().Trim(); //SMPL[1].Replace("-" + SMPL[2], "");
                            if (SMPL[2] != "0")
                                objSheet.Cells[2, 2] = "Revised-" + SMPL[2];

                            objSheet.Cells[2, 11] = SMPL[25];
                            objSheet.Cells[4, 2] = SMPL[3] + " " + SMPL[4];
                            objSheet.Cells[5, 2] = SMPL[5];
                            objSheet.Cells[1, 11] = SMPL[6]==""?"":Convert.ToDateTime(SMPL[6]).ToString("dd/MM/yyyy");
                            objSheet.Cells[7, 2] = SMPL[7] + " " + SMPL[8] + " " + SMPL[9];
                            objSheet.Cells[9, 2] = "- " + SMPL[10] + " -";
                            objSheet.Cells[10, 2] = SMPL[11];
                            objSheet.Cells[11, 2] = SMPL[12];
                            objSheet.Cells[9, 11] = SMPL[13];
                            objSheet.Cells[10, 11] = "ใช้แพทเทิร์น " + SMPL[14];
                            objSheet.Cells[5, 11] = SMPL[17] == "" ? "SMPL" : SMPL[17].ToUpper().Trim();

                            if (SMPL[24] != "")
                            {
                                objSheet.Cells[6, 5] = SMPL[24];
                                objSheet.Range[objSheet.Cells[6, 5], objSheet.Cells[6, 6]].Interior.Color = Color.Yellow;

                            }

                            pbcEXPORT.PerformStep();
                            pbcEXPORT.Update();

                            int LastCol = 10;
                            //********* Set Column ************
                            int CountCol = DBC.DBQuery("SELECT COUNT(OIDSMPLDT) AS CountCol FROM SMPLQuantityRequired WHERE (OIDSMPL = '" + SMPL[0] + "')").getInt();
                            //CountCol = 8;
                            if (CountCol > 6)
                            {
                                for (int ci = 0; ci < CountCol - 6; ci++)
                                {
                                    objSheet.Columns[7].Insert();
                                    LastCol++;
                                }
                            }

                            pbcEXPORT.PerformStep();
                            pbcEXPORT.Update();

                            //Set Column Size
                            float LWCol = 0;
                            float NWCol = 0;

                            if (CountCol < 6)
                            {
                                LWCol = (float)((double)60 / (double)CountCol);
                                NWCol = (float)((double)17.38 / (double)(6-CountCol));
                            }
                            else
                            {
                                LWCol = (float)((double)77.38 / (double)CountCol);
                                NWCol = 0;
                            }

                            int loopCountCol = 0;
                            for (int sc = 4; sc < LastCol; sc++)
                            {
                                if (loopCountCol < CountCol)
                                {
                                    objSheet.Columns[sc].ColumnWidth = LWCol;
                                }
                                else
                                {
                                    objSheet.Columns[sc].ColumnWidth = NWCol;
                                }
                                loopCountCol++;
                            }
                            //******* End Set Column **********

                            Microsoft.Office.Interop.Excel.Range oRange;
                            float Left = 0;
                            float Top = 0;

                            if (SMPL[15] != "")
                            {
                                oRange = (Microsoft.Office.Interop.Excel.Range)objSheet.Cells[38, 3];
                                Left = (float)((double)oRange.Left) + 1;
                                Top = (float)((double)oRange.Top) + 1;
                                string PathImgFile = this.imgPath + SMPL[15];
                                Bitmap original = new Bitmap(PathImgFile);
                                float scaleHeight = 200;
                                float scaleWidth = (scaleHeight * original.Width) / original.Height;
                                objSheet.Shapes.AddPicture(System.IO.Path.Combine(PathImgFile), Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, Left, Top, (int)scaleWidth, (int)scaleHeight);
                                original.Dispose();
                            }
                            objSheet.Cells[28, LastCol + 1] = "※ TAG";
                            objSheet.Cells[30, 2] = SMPL[16]==""?"":Convert.ToDateTime(SMPL[16]).ToString("dd-MMM-yy");
                            objSheet.Cells[30, LastCol + 1] = SMPL[16] == "" ? "" : "ต้องการตัวอย่าง " + Convert.ToDateTime(SMPL[16]).ToString("MMMM-dd-yyyy");
                            objSheet.Cells[30, LastCol + 1].Font.Size = 14;
                            objSheet.Cells[31, 2] = SMPL[17];
                            objSheet.Cells[34, 2] = SMPL[18];

                            objSheet.Cells[29, LastCol] = SMPL[19];
                            objSheet.Cells[28, 2] = SMPL[20];
                            string xPCS = SMPL[21].Replace(" ", "");
                            xPCS = xPCS.Length > 0 ? xPCS.Substring(0, xPCS.Length - 1) : "";
                            objSheet.Cells[28, 5] = xPCS + SMPL[19].ToLower();
                            objSheet.Cells[29, LastCol - 1] = SMPL[22];

                            pbcEXPORT.PerformStep();
                            pbcEXPORT.Update();

                            sbSQL.Clear();
                            sbSQL.Append("SELECT SQR.OIDSIZE, PS.SizeName ");
                            sbSQL.Append("FROM   SMPLQuantityRequired AS SQR INNER JOIN ");
                            sbSQL.Append("       ProductSize AS PS ON SQR.OIDSIZE = PS.OIDSIZE ");
                            sbSQL.Append("WHERE (SQR.OIDSMPL = '" + SMPL[0] + "') ");
                            sbSQL.Append("ORDER BY SQR.OIDSIZE, SQR.OIDCOLOR ");
                            DataTable dtRQ = DBC.DBQuery(sbSQL.ToString()).getDataTable();
                            if (dtRQ.Rows.Count > 0)
                            {
                                int runCell = 4;
                                string chkSizeID = "";
                                foreach (DataRow drRQ in dtRQ.Rows)
                                {
                                    string SizeID = drRQ["OIDSIZE"].ToString();
                                    string SizeName = drRQ["SizeName"].ToString();
                                    if (chkSizeID != SizeID)
                                    {
                                        objSheet.Cells[13, runCell] = SizeName;
                                        chkSizeID = SizeID;
                                    }
                                    else //Merge Cell
                                    {
                                        objSheet.Range[objSheet.Cells[13, runCell - 1], objSheet.Cells[13, runCell]].Merge();
                                        objSheet.Cells[13, runCell] = SizeName;
                                    }
                                    runCell++;
                                }

                                if (runCell <= LastCol)
                                {
                                    objSheet.Range[objSheet.Cells[13, runCell], objSheet.Cells[13, LastCol]].Merge();
                                }
                            }

                            pbcEXPORT.PerformStep();
                            pbcEXPORT.Update();

                            sbSQL.Clear();
                            sbSQL.Append("SELECT SQR.OIDCOLOR, PC.ColorName ");
                            sbSQL.Append("FROM   SMPLQuantityRequired AS SQR INNER JOIN ");
                            sbSQL.Append("       ProductColor AS PC ON SQR.OIDCOLOR = PC.OIDCOLOR ");
                            sbSQL.Append("WHERE (SQR.OIDSMPL = '" + SMPL[0] + "') ");
                            sbSQL.Append("ORDER BY SQR.OIDSIZE, SQR.OIDCOLOR ");
                            DataTable dtRQ2 = DBC.DBQuery(sbSQL.ToString()).getDataTable();
                            if (dtRQ2.Rows.Count > 0)
                            {
                                int runCell = 4;
                                string chkColorID = "";
                                foreach (DataRow drRQ in dtRQ2.Rows)
                                {
                                    string ColorID = drRQ["OIDCOLOR"].ToString();
                                    string ColorName = drRQ["ColorName"].ToString();
                                    if (chkColorID != ColorID)
                                    {
                                        objSheet.Cells[14, runCell] = ColorName;
                                        chkColorID = ColorID;
                                    }
                                    else //Merge Cell
                                    {
                                        objSheet.Range[objSheet.Cells[14, runCell - 1], objSheet.Cells[14, runCell]].Merge();
                                        objSheet.Cells[14, runCell] = ColorName;
                                    }
                                    runCell++;
                                }

                                blankCol = runCell;
                                if (runCell <= LastCol)
                                {
                                    objSheet.Range[objSheet.Cells[14, runCell], objSheet.Cells[14, LastCol]].Merge();
                                }
                            }

                            pbcEXPORT.PerformStep();
                            pbcEXPORT.Update();

                            sbSQL.Clear();
                            sbSQL.Append("SELECT Quantity ");
                            sbSQL.Append("FROM   SMPLQuantityRequired ");
                            sbSQL.Append("WHERE (OIDSMPL = '" + SMPL[0] + "') ");
                            sbSQL.Append("ORDER BY OIDSIZE, OIDCOLOR ");
                            DataTable dtRQ3 = DBC.DBQuery(sbSQL.ToString()).getDataTable();
                            if (dtRQ3.Rows.Count > 0)
                            {
                                int runCell = 4;
                                foreach (DataRow drRQ in dtRQ3.Rows)
                                {
                                    string Quantity = drRQ["Quantity"].ToString();
                                    objSheet.Cells[16, runCell] = Quantity;
                                    runCell++;
                                }

                                if (runCell <= LastCol)
                                {
                                    objSheet.Range[objSheet.Cells[16, runCell], objSheet.Cells[16, LastCol]].Merge();
                                    //Hanger
                                    objSheet.Range[objSheet.Cells[17, runCell], objSheet.Cells[17, LastCol]].Merge();
                                }
                            }

                            pbcEXPORT.PerformStep();
                            pbcEXPORT.Update();

                            int BeginFB = 14;
                            int LastRow = BeginFB;

                            int UseRow = BeginFB;

                            int BeginMT = BeginFB;
                            int BeginFBComp = BeginFB + 4;
                            int BeginMTComp = BeginFB + 7;

                            sbSQL.Clear();
                            sbSQL.Append("SELECT RQ.OIDSMPLDT, PT.OIDGParts, PT.GarmentParts, ");
                            sbSQL.Append("       ISNULL((SELECT ColorName + ',' AS 'data()' FROM(SELECT DISTINCT Z.ColorName FROM SMPLRequestFabric AS A INNER JOIN SMPLRequestFabricParts AS YA ON A.OIDSMPLFB = YA.OIDSMPLFB INNER JOIN ProductColor AS Z ON A.OIDCOLOR = Z.OIDCOLOR WHERE A.OIDSMPLDT = RQ.OIDSMPLDT  AND YA.OIDGParts = PT.OIDGParts) AS XA FOR XML PATH('')), '')  AS FBCodeColor, ");
                            sbSQL.Append("       ISNULL((SELECT VendFBCode + ',' AS 'data()' FROM(SELECT DISTINCT B.VendFBCode FROM SMPLRequestFabric AS B INNER JOIN SMPLRequestFabricParts AS YB ON B.OIDSMPLFB = YB.OIDSMPLFB WHERE B.OIDSMPLDT = RQ.OIDSMPLDT  AND YB.OIDGParts = PT.OIDGParts) AS XB FOR XML PATH('')), '')  AS VendorFBCode, ");
                            sbSQL.Append("       ISNULL((SELECT SMPLotNo + ',' AS 'data()' FROM(SELECT DISTINCT SMPLotNo FROM SMPLRequestFabric AS C INNER JOIN SMPLRequestFabricParts AS YC ON C.OIDSMPLFB = YC.OIDSMPLFB WHERE C.OIDSMPLDT = RQ.OIDSMPLDT  AND YC.OIDGParts = PT.OIDGParts) AS XC FOR XML PATH('')), '')  AS FBLotNo, ");
                            sbSQL.Append("       ISNULL((SELECT Name + ',' AS 'data()' FROM(SELECT DISTINCT REPLACE(REPLACE(VD.Name, ' CO.,LTD.', ''), ' CO.,LTD', '') AS Name FROM SMPLRequestFabric AS D INNER JOIN SMPLRequestFabricParts AS YD ON D.OIDSMPLFB = YD.OIDSMPLFB INNER JOIN Vendor AS VD ON D.OIDVEND = VD.OIDVEND WHERE D.OIDSMPLDT = RQ.OIDSMPLDT  AND YD.OIDGParts = PT.OIDGParts) AS XD FOR XML PATH('')), '')  AS Vendor, ");
                            sbSQL.Append("       ISNULL((SELECT CASE WHEN ISNULL(PathFile, '') = '' THEN '' ELSE PathFile + ',' END AS 'data()' FROM(SELECT DISTINCT E.PathFile FROM SMPLRequestFabric AS E INNER JOIN SMPLRequestFabricParts AS YE ON E.OIDSMPLFB = YE.OIDSMPLFB WHERE E.OIDSMPLDT = RQ.OIDSMPLDT  AND YE.OIDGParts = PT.OIDGParts) AS XE FOR XML PATH('')),'')  AS FBFile, ");
                            sbSQL.Append("       ISNULL((SELECT CASE WHEN ISNULL(Composition, '') = '' THEN '' ELSE Composition + ',' END AS 'data()' FROM(SELECT DISTINCT F.Composition FROM SMPLRequestFabric AS F INNER JOIN SMPLRequestFabricParts AS YF ON F.OIDSMPLFB = YF.OIDSMPLFB WHERE F.OIDSMPLDT = RQ.OIDSMPLDT  AND YF.OIDGParts = PT.OIDGParts) AS XF FOR XML PATH('')), '')  AS Composition, ");
                            sbSQL.Append("       ISNULL((SELECT Remark + ',' AS 'data()' FROM(SELECT DISTINCT Remark FROM SMPLRequestFabric AS G INNER JOIN SMPLRequestFabricParts AS YG ON G.OIDSMPLFB = YG.OIDSMPLFB WHERE G.OIDSMPLDT = RQ.OIDSMPLDT  AND YG.OIDGParts = PT.OIDGParts) AS XG FOR XML PATH('')), '')  AS FBRemark  ");
                            sbSQL.Append("FROM   SMPLQuantityRequired AS RQ CROSS JOIN ");
                            sbSQL.Append("       (SELECT DISTINCT FP.OIDGParts, GP.GarmentParts ");
                            sbSQL.Append("        FROM   SMPLRequestFabricParts AS FP INNER JOIN ");
                            sbSQL.Append("               GarmentParts AS GP ON FP.OIDGParts = GP.OIDGParts ");
                            sbSQL.Append("        WHERE  (FP.OIDSMPLDT IN ");
                            sbSQL.Append("                  (SELECT OIDSMPLDT ");
                            sbSQL.Append("                   FROM   SMPLQuantityRequired AS A ");
                            sbSQL.Append("                   WHERE  (OIDSMPL = '" + SMPL[0] + "')))) AS PT ");
                            sbSQL.Append("WHERE (RQ.OIDSMPL = '" + SMPL[0] + "') ");
                            sbSQL.Append("ORDER BY PT.OIDGParts, RQ.OIDSIZE, RQ.OIDCOLOR ");

                            DataTable dtFB = DBC.DBQuery(sbSQL.ToString()).getDataTable();
                            if (dtFB.Rows.Count > 0)
                            {
                                DataTable chkDtFB = dtFB;
                                string Supplier = "";
                                string Composition = "";

                                string ChkComposition = "";
                                string ChkSupplier = "";

                                string chkParts = "";
                                int runCell = 4;
                                int runLoop = 0;

                                DataTable dtComposition = new DataTable();
                                dtComposition.Columns.Add("Composition", typeof(String));
                                dtComposition.Columns.Add("Supplier", typeof(String));

                                bool chkFBCode = false;
                                bool chkFBLot = false;
                                bool chkFBRemark = false;
                                foreach (DataRow drFB in dtFB.Rows)
                                {
                                    string OIDGParts = drFB["OIDGParts"].ToString();
                                    string GarmentParts = drFB["GarmentParts"].ToString();

                                    if (chkParts != OIDGParts)
                                    {
                                        UseRow++;
                                        if (chkFBCode == true)
                                            UseRow++;
                                        if (chkFBLot == true)
                                            UseRow++;
                                        if (chkFBRemark == true)
                                            UseRow++;

                                        if (UseRow > 15)
                                        {
                                            objSheet.Rows[UseRow].Insert();
                                            objSheet.Range[objSheet.Cells[UseRow, 2], objSheet.Cells[UseRow, 3]].Merge();
                                            LastRow = UseRow;
                                        }
                                        objSheet.Cells[UseRow, 2] = GarmentParts.ToUpper().Trim();
                                        objSheet.Cells[UseRow, 2].Font.Size = 14;
                                        objSheet.Cells[UseRow, 2].Font.Color = Color.Black;
                                        if (blankCol <= LastCol)
                                        {
                                            objSheet.Cells[UseRow, blankCol] = "สีผ้า " + GarmentParts;
                                        }
                                        chkParts = OIDGParts;
                                        
                                        //** FB CODE*************************************
                                        chkFBCode = false;
                                        chkFBLot = false;
                                        chkFBRemark = false;
                                        foreach (DataRow drChkFB in chkDtFB.Rows)
                                        {
                                            string OIDGP = drChkFB["OIDGParts"].ToString();
                                            string FBColor = drChkFB["VendorFBCode"].ToString().Trim();
                                            FBColor = FBColor.Length > 0 ? FBColor.Substring(0, FBColor.Length - 1) : "";
                                            if (OIDGParts == OIDGP && FBColor != "")
                                            {
                                                chkFBCode = true;
                                                break;
                                            }
                                        }

                                        if (chkFBCode == true)
                                        {
                                            objSheet.Rows[UseRow + 1].Insert();
                                            objSheet.Range[objSheet.Cells[UseRow + 1, 2], objSheet.Cells[UseRow + 1, 3]].Merge();
                                            objSheet.Cells[UseRow + 1, 2] = "FABRIC ITEM";
                                            if (blankCol <= LastCol)
                                            {
                                                objSheet.Cells[UseRow + 1, blankCol] = "รหัสผ้า";
                                            }
                                            LastRow = UseRow + 1;
                                        }

                                        //** FB LOT*************************************
                                        foreach (DataRow drChkFB in chkDtFB.Rows)
                                        {
                                            string OIDGP = drChkFB["OIDGParts"].ToString();
                                            string FBLot = drChkFB["FBLotNo"].ToString().Trim();
                                            FBLot = FBLot.Length > 0 ? FBLot.Substring(0, FBLot.Length - 1) : "";
                                            if (OIDGParts == OIDGP && FBLot != "")
                                            {
                                                chkFBLot = true;
                                                break;
                                            }
                                        }

                                        if (chkFBLot == true)
                                        {
                                            if (chkFBCode == true)
                                            {
                                                objSheet.Rows[UseRow + 2].Insert();
                                                objSheet.Range[objSheet.Cells[UseRow + 2, 2], objSheet.Cells[UseRow + 2, 3]].Merge();
                                                objSheet.Cells[UseRow + 2, 2] = "FABRIC LOT";
                                                if (blankCol <= LastCol)
                                                {
                                                    objSheet.Cells[UseRow + 2, blankCol] = "ล็อตผ้า";
                                                }
                                                LastRow = UseRow + 2;
                                            }
                                            else
                                            {
                                                objSheet.Rows[UseRow + 1].Insert();
                                                objSheet.Range[objSheet.Cells[UseRow + 1, 2], objSheet.Cells[UseRow + 1, 3]].Merge();
                                                objSheet.Cells[UseRow + 1, 2] = "FABRIC LOT";
                                                if (blankCol <= LastCol)
                                                {
                                                    objSheet.Cells[UseRow + 1, blankCol] = "ล็อตผ้า";
                                                }
                                                LastRow = UseRow + 1;
                                            }
                                        }

                                        //** FB Remark*************************************
                                        foreach (DataRow drChkFB in chkDtFB.Rows)
                                        {
                                            string OIDGP = drChkFB["OIDGParts"].ToString();
                                            string FBxRemark = drChkFB["FBRemark"].ToString().Trim();
                                            FBxRemark = FBxRemark.Length > 0 ? FBxRemark.Substring(0, FBxRemark.Length - 1) : "";
                                            if (OIDGParts == OIDGP && FBxRemark != "")
                                            {
                                                chkFBRemark = true;
                                                break;
                                            }
                                        }

                                        if (chkFBRemark == true)
                                        {
                                            if (chkFBCode == true)
                                            {
                                                if (chkFBLot == true)
                                                {
                                                    objSheet.Rows[UseRow + 3].Insert();
                                                    objSheet.Range[objSheet.Cells[UseRow + 3, 2], objSheet.Cells[UseRow + 3, 3]].Merge();
                                                    objSheet.Cells[UseRow + 3, 2] = "FABRIC REMARK";
                                                    if (blankCol <= LastCol)
                                                    {
                                                        objSheet.Cells[UseRow + 3, blankCol] = "หมายเหตุ";
                                                    }
                                                    LastRow = UseRow + 3;
                                                }
                                                else
                                                {
                                                    objSheet.Rows[UseRow + 2].Insert();
                                                    objSheet.Range[objSheet.Cells[UseRow + 2, 2], objSheet.Cells[UseRow + 2, 3]].Merge();
                                                    objSheet.Cells[UseRow + 2, 2] = "FABRIC REMARK";
                                                    if (blankCol <= LastCol)
                                                    {
                                                        objSheet.Cells[UseRow + 2, blankCol] = "หมายเหตุ";
                                                    }
                                                    LastRow = UseRow + 2;
                                                }
                                            }
                                            else if (chkFBLot == true)
                                            {
                                                objSheet.Rows[UseRow + 2].Insert();
                                                objSheet.Range[objSheet.Cells[UseRow + 2, 2], objSheet.Cells[UseRow + 2, 3]].Merge();
                                                objSheet.Cells[UseRow + 2, 2] = "FABRIC REMARK";
                                                if (blankCol <= LastCol)
                                                {
                                                    objSheet.Cells[UseRow + 2, blankCol] = "หมายเหตุ";
                                                }
                                                LastRow = UseRow + 2;
                                            }
                                            else
                                            {
                                                objSheet.Rows[UseRow + 1].Insert();
                                                objSheet.Range[objSheet.Cells[UseRow + 1, 2], objSheet.Cells[UseRow + 1, 3]].Merge();
                                                objSheet.Cells[UseRow + 1, 2] = "FABRIC REMARK";
                                                if (blankCol <= LastCol)
                                                {
                                                    objSheet.Cells[UseRow + 1, blankCol] = "หมายเหตุ";
                                                }
                                                LastRow = UseRow + 1;
                                            }
                                        }

                                        runCell = 4;
                                    }

                                    if (GarmentParts.ToUpper().Trim() == "BODY")
                                    {
                                        string FBFile = drFB["FBFile"].ToString().Trim();
                                        FBFile = FBFile.Length > 0 ? FBFile.Substring(0, FBFile.Length - 1) : "";
                                        FBFile = FBFile.Trim().Replace(" ", "");
                                        if (FBFile.Length > 0)
                                        {
                                            if (FBFile.Substring(0, 1) == ",")
                                                FBFile = FBFile.Substring(1);
                                        }

                                        if (FBFile.Length > 0)
                                        {
                                            if (FBFile.Substring(FBFile.Length - 1, 1) == ",")
                                                FBFile = FBFile.Substring(0, FBFile.Length - 1);
                                        }

                                        //objSheet.Cells[12, runCell] = FBFile;
                                        if (FBFile != "")
                                        {
                                            if (FBFile.IndexOf(',') > -1) //มากกว่า 1 รูป
                                            {
                                                string[] imgFile = FBFile.Split(',');
                                                int TTImg = imgFile.Length;

                                                for (int i = 0; i < imgFile.Length; i++)
                                                {
                                                    string FabricFile = "";
                                                    if (imgFile[i] != "")
                                                    {
                                                        if (imgFile[i].IndexOf('/') > -1)
                                                            FabricFile = imgFile[i];
                                                        else
                                                            FabricFile = this.imgPath + imgFile[i];
                                                    }

                                                    if (FabricFile != "")
                                                    {
                                                        oRange = (Microsoft.Office.Interop.Excel.Range)objSheet.Cells[12, runCell];
                                                        if (i == 0)
                                                            Left = (float)((double)oRange.Left) + 1;
                                                        else
                                                            Left = (float)(((double)oRange.Left + 1 + ((double)(oRange.Width / TTImg) * i)));
                                                        Top = (float)((double)oRange.Top) + 1;
                                                        objSheet.Shapes.AddPicture(System.IO.Path.Combine(FabricFile), Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, Left, Top, (int)(oRange.Width - 2) / TTImg, oRange.Height - 2);
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                string FabricFile = "";
                                                if (FBFile != "")
                                                {
                                                    if (FBFile.IndexOf('/') > -1)
                                                        FabricFile = FBFile;
                                                    else
                                                        FabricFile = this.imgPath + FBFile;
                                                }

                                                if (FabricFile != "")
                                                {
                                                    oRange = (Microsoft.Office.Interop.Excel.Range)objSheet.Cells[12, runCell];
                                                    Left = (float)((double)oRange.Left) + 1;
                                                    Top = (float)((double)oRange.Top) + 1;
                                                    objSheet.Shapes.AddPicture(System.IO.Path.Combine(FabricFile), Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, Left, Top, oRange.Width - 2, oRange.Height - 2);
                                                }
                                            }
                                        }
                                    }

                                    string FBCodeColor = drFB["FBCodeColor"].ToString().Trim();
                                    FBCodeColor = FBCodeColor.Length > 0 ? FBCodeColor.Substring(0, FBCodeColor.Length - 1) : "";

                                    string VendorFBCode = drFB["VendorFBCode"].ToString().Trim();
                                    VendorFBCode = VendorFBCode.Length > 0 ? VendorFBCode.Substring(0, VendorFBCode.Length - 1) : "";

                                    string FBLotNo = drFB["FBLotNo"].ToString().Trim();
                                    FBLotNo = FBLotNo.Length > 0 ? FBLotNo.Substring(0, FBLotNo.Length - 1) : "";

                                    string FBRemark = drFB["FBRemark"].ToString().Trim();
                                    FBRemark = FBRemark.Length > 0 ? FBRemark.Substring(0, FBRemark.Length - 1) : "";

                                    Supplier = drFB["Vendor"].ToString().Trim();
                                    Supplier = Supplier.Length > 0 ? Supplier.Substring(0, Supplier.Length - 1) : "";

                                    Composition = drFB["Composition"].ToString().Trim();
                                    Composition = Composition.Length > 0 ? Composition.Substring(0, Composition.Length - 1) : "";

                                    if (ChkComposition != Composition)
                                    {
                                        string xSupplier = "";
                                        if (Supplier != ChkSupplier)
                                        {
                                            xSupplier = Supplier;
                                            ChkSupplier = Supplier;
                                        }
                                        dtComposition.Rows.Add(Composition, xSupplier);
                                        ChkComposition = Composition;   
                                    }

                                    objSheet.Cells[UseRow, runCell] = FBCodeColor;

                                    if (chkFBCode == true)
                                    {
                                        objSheet.Cells[UseRow + 1, runCell] = VendorFBCode;
                                    }

                                    if (chkFBLot == true)
                                    {
                                        if (chkFBCode == true)
                                        {
                                            objSheet.Cells[UseRow + 2, runCell] = FBLotNo;
                                        }
                                        else
                                        {
                                            objSheet.Cells[UseRow + 1, runCell] = FBLotNo;
                                        }
                                    }

                                    if (chkFBRemark == true)
                                    {
                                        if (chkFBCode == true)
                                        {
                                            if (chkFBLot == true)
                                            {
                                                objSheet.Cells[UseRow + 3, runCell] = FBRemark;
                                            }
                                            else
                                            {
                                                objSheet.Cells[UseRow + 2, runCell] = FBRemark;
                                            }
                                        }
                                        else if (chkFBLot == true)
                                        {
                                            objSheet.Cells[UseRow + 2, runCell] = FBRemark;
                                        }
                                        else
                                        {
                                            objSheet.Cells[UseRow + 1, runCell] = FBRemark;
                                        }
                                    }


                                    runCell++;
                                    runLoop++;
                                }

                                for (int xi = BeginFB; xi <= LastRow; xi++)
                                {
                                    if (blankCol <= LastCol)
                                    {
                                        objSheet.Range[objSheet.Cells[xi, blankCol], objSheet.Cells[xi, LastCol]].Merge();
                                        objSheet.Range[objSheet.Cells[xi, blankCol], objSheet.Cells[xi, LastCol]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                                        objSheet.Cells[xi, blankCol].Font.Size = 13;

                                        if (objSheet.Cells[xi, 2].Value.ToString() == "FABRIC ITEM" || objSheet.Cells[xi, 2].Value.ToString() == "FABRIC LOT" || objSheet.Cells[xi, 2].Value.ToString() == "FABRIC REMARK")
                                        {
                                            objSheet.Range[objSheet.Cells[xi, 2], objSheet.Cells[xi, LastCol]].Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDot;
                                            objSheet.Range[objSheet.Cells[xi, 2], objSheet.Cells[xi, LastCol]].Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                                            if (objSheet.Cells[xi, 2].Value.ToString() == "FABRIC REMARK")
                                            {
                                                objSheet.Range[objSheet.Cells[xi, 2], objSheet.Cells[xi, LastCol]].Font.Size = 11;
                                                objSheet.Range[objSheet.Cells[xi, 2], objSheet.Cells[xi, LastCol]].Font.Color = Color.Red;
                                                objSheet.Range[objSheet.Cells[xi, 2], objSheet.Cells[xi, LastCol]].Rows.AutoFit();
                                            }
                                        }
                                        else
                                        {
                                            if (xi == BeginFB + 1)
                                            {
                                                objSheet.Range[objSheet.Cells[xi, 2], objSheet.Cells[xi, LastCol + 3]].Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                                                objSheet.Range[objSheet.Cells[xi, 2], objSheet.Cells[xi, LastCol + 3]].Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick;
                                            }
                                            else
                                            {
                                                objSheet.Range[objSheet.Cells[xi, 2], objSheet.Cells[xi, LastCol]].Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                                                objSheet.Range[objSheet.Cells[xi, 2], objSheet.Cells[xi, LastCol]].Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                                            }
                                        }

                                    }
                                }

                                BeginMT = LastRow;

                                

                                //**** Fabric Composition ****
                                LastRow += 3;
                                BeginFBComp = LastRow;
                                if (dtComposition.Rows.Count > 0)
                                {
                                    int chkRow = 0;
                                    foreach (DataRow drComp in dtComposition.Rows)
                                    {
                                        if (chkRow > 2)
                                        {
                                            objSheet.Rows[BeginFBComp].Insert();
                                            objSheet.Range[objSheet.Cells[BeginFBComp - 1, 1], objSheet.Cells[BeginFBComp, 1]].Merge();
                                            objSheet.Range[objSheet.Cells[BeginFBComp, 2], objSheet.Cells[BeginFBComp, LastCol+3]].Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                        }
                                        objSheet.Cells[BeginFBComp, 2] = drComp[0].ToString();
                                        objSheet.Cells[BeginFBComp, LastCol+1] = drComp[1].ToString();
                                        LastRow++;
                                        BeginFBComp++;
                                        chkRow++;
                                    }

                                }

                            }

                            pbcEXPORT.PerformStep();
                            pbcEXPORT.Update();

                            //**** Material ****
                            sbSQL.Clear();
                            sbSQL.Append("SELECT    RQ.OIDSMPLDT, MT.OIDITEM, MT.Code, MT.Description, ");
                            sbSQL.Append("          ISNULL((SELECT VendMTCode + ',' AS 'data()' FROM(SELECT DISTINCT B.VendMTCode FROM SMPLRequestMaterial AS B INNER JOIN Items AS YB ON B.OIDITEM = YB.OIDITEM WHERE B.OIDSMPLDT = RQ.OIDSMPLDT  AND YB.OIDITEM = MT.OIDITEM) AS XB FOR XML PATH('')), '')  AS VendorMTCode, ");
                            sbSQL.Append("          ISNULL((SELECT Name + ',' AS 'data()' FROM(SELECT DISTINCT REPLACE(REPLACE(VD.Name, ' CO.,LTD.', ''), ' CO.,LTD', '') AS Name FROM SMPLRequestMaterial AS D INNER JOIN Items AS YD ON D.OIDITEM = YD.OIDITEM INNER JOIN Vendor AS VD ON D.OIDVEND = VD.OIDVEND WHERE D.OIDSMPLDT = RQ.OIDSMPLDT  AND YD.OIDITEM = MT.OIDITEM) AS XD FOR XML PATH('')), '')  AS Vendor, ");
                            sbSQL.Append("          ISNULL((SELECT CASE WHEN ISNULL(Composition, '') = '' THEN '' ELSE Composition + ',' END AS 'data()' FROM(SELECT DISTINCT F.Composition FROM SMPLRequestMaterial AS F INNER JOIN Items AS YF ON F.OIDITEM = YF.OIDITEM WHERE F.OIDSMPLDT = RQ.OIDSMPLDT  AND YF.OIDITEM = MT.OIDITEM) AS XF FOR XML PATH('')), '')  AS Composition, ");
                            sbSQL.Append("          ISNULL((SELECT CASE WHEN ISNULL(Situation, '') = '' THEN '' ELSE Situation + ',' END AS 'data()' FROM(SELECT DISTINCT G.Situation FROM SMPLRequestMaterial AS G INNER JOIN Items AS YG ON G.OIDITEM = YG.OIDITEM WHERE G.OIDSMPLDT = RQ.OIDSMPLDT  AND YG.OIDITEM = MT.OIDITEM) AS XG FOR XML PATH('')), '')  AS Situation, ");
                            sbSQL.Append("          ISNULL((SELECT CASE WHEN ISNULL(Comment, '') = '' THEN '' ELSE Comment + ',' END AS 'data()' FROM(SELECT DISTINCT H.Comment FROM SMPLRequestMaterial AS H INNER JOIN Items AS YH ON H.OIDITEM = YH.OIDITEM WHERE H.OIDSMPLDT = RQ.OIDSMPLDT  AND YH.OIDITEM = MT.OIDITEM) AS XH FOR XML PATH('')), '')  AS Comment, ");
                            sbSQL.Append("          ISNULL((SELECT CASE WHEN ISNULL(Remark, '') = '' THEN '' ELSE Remark + ',' END AS 'data()' FROM(SELECT DISTINCT I.Remark FROM SMPLRequestMaterial AS I INNER JOIN Items AS YI ON I.OIDITEM = YI.OIDITEM WHERE I.OIDSMPLDT = RQ.OIDSMPLDT  AND YI.OIDITEM = MT.OIDITEM) AS XI FOR XML PATH('')), '')  AS Remark ");
                            sbSQL.Append("FROM      SMPLQuantityRequired AS RQ CROSS JOIN ");
                            sbSQL.Append("          (SELECT DISTINCT IT.OIDITEM, IT.Code, IT.Description ");
                            sbSQL.Append("           FROM   SMPLRequestMaterial AS SMT INNER JOIN ");
                            sbSQL.Append("                  Items AS IT ON SMT.OIDITEM = IT.OIDITEM ");
                            sbSQL.Append("           WHERE  (SMT.OIDSMPLDT IN ");
                            sbSQL.Append("                             (SELECT OIDSMPLDT ");
                            sbSQL.Append("                              FROM   SMPLQuantityRequired AS A ");
                            sbSQL.Append("                              WHERE  (OIDSMPL = '" + SMPL[0] + "')))) AS MT ");
                            sbSQL.Append("WHERE (RQ.OIDSMPL = '" + SMPL[0] + "') ");
                            sbSQL.Append("ORDER BY MT.OIDITEM, RQ.OIDSIZE, RQ.OIDCOLOR ");
                            DataTable dtMT = DBC.DBQuery(sbSQL.ToString()).getDataTable();
                            if (dtMT.Rows.Count > 0)
                            {
                                LastRow = BeginMT;
                                UseRow = BeginMT;
                                string chkITEM = "";

                                string MTSupplier = "";
                                string MTComposition = "";

                                string ChkMTComposition = "";
                                string ChkMTSupplier = "";

                                DataTable dtMTComposition = new DataTable();
                                dtMTComposition.Columns.Add("Composition", typeof(String));
                                dtMTComposition.Columns.Add("Supplier", typeof(String));

                                int runCell = 4;
                                int runLoop = 0;
                                foreach (DataRow drMT in dtMT.Rows)
                                {
                                    string OIDITEM = drMT["OIDITEM"].ToString();
                                    string Code = drMT["Code"].ToString();
                                    string Description = drMT["Description"].ToString();

                                    if (chkITEM != OIDITEM)
                                    {
                                        UseRow++;
                                        if (UseRow > 15)
                                        {
                                            objSheet.Rows[UseRow].Insert();
                                            objSheet.Range[objSheet.Cells[UseRow, 2], objSheet.Cells[UseRow, 3]].Merge();
                                            LastRow = UseRow;
                                        }
                                        objSheet.Cells[UseRow, 2] = Description.ToUpper().Trim();
                                        objSheet.Cells[UseRow, 2].Font.Size = 14;
                                        objSheet.Cells[UseRow, 2].Font.Color = Color.Black;
                                        //if (blankCol <= LastCol)
                                        //    objSheet.Cells[UseRow, blankCol] = "สีผ้า " + GarmentParts;
                                        chkITEM = OIDITEM;

                                        runCell = 4;
                                    }

                                    string VendorMTCode = drMT["VendorMTCode"].ToString().Trim();
                                    VendorMTCode = VendorMTCode.Length > 0 ? VendorMTCode.Substring(0, VendorMTCode.Length - 1) : "";
                                    objSheet.Cells[UseRow, runCell] = VendorMTCode;

                                    string Situation = drMT["Situation"].ToString().Trim();
                                    Situation = Situation.Length > 0 ? Situation.Substring(0, Situation.Length - 1) : "";

                                    //if (Situation != "")
                                    //{
                                    //    if (blankCol <= LastCol)
                                    //    {
                                    //        objSheet.Cells[UseRow, blankCol] = Situation;
                                    //    }
                                    //}

                                    string Comment = drMT["Comment"].ToString().Trim();
                                    Comment = Comment.Length > 0 ? Comment.Substring(0, Comment.Length - 1) : "";
                                    string Remark = drMT["Remark"].ToString().Trim();
                                    Remark = Remark.Length > 0 ? Remark.Substring(0, Remark.Length - 1) : "";
                                    string Recommend = "";
                                    if (Situation != "")
                                        Recommend += Situation;
                                    if (Comment != "")
                                    {
                                        if (Recommend != "")
                                            Recommend += " / ";
                                        Recommend += Comment;
                                    }
                                    if (Remark != "")
                                    {
                                        if (Recommend != "")
                                            Recommend += " / ";
                                        Recommend += Remark;
                                    }
                                    objSheet.Range[objSheet.Cells[UseRow, LastCol + 1], objSheet.Cells[UseRow, LastCol + 3]].Merge();
                                    objSheet.Range[objSheet.Cells[UseRow, LastCol + 1], objSheet.Cells[UseRow, LastCol + 3]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                                    objSheet.Cells[UseRow, LastCol + 1] = Recommend;

                                    MTSupplier = drMT["Vendor"].ToString().Trim();
                                    MTSupplier = MTSupplier.Length > 0 ? MTSupplier.Substring(0, MTSupplier.Length - 1) : "";

                                    MTComposition = drMT["Composition"].ToString().Trim();
                                    MTComposition = MTComposition.Length > 0 ? MTComposition.Substring(0, MTComposition.Length - 1) : "";

                                    if (ChkMTComposition != MTComposition)
                                    {
                                        string xSupplier = "";
                                        if (MTSupplier != ChkMTSupplier)
                                        {
                                            xSupplier = MTSupplier;
                                            ChkMTSupplier = MTSupplier;
                                        }
                                        dtMTComposition.Rows.Add(MTComposition, xSupplier);
                                        ChkMTComposition = MTComposition;
                                    }

                                    runCell++;
                                    runLoop++;
                                }


                                objSheet.Range[objSheet.Cells[BeginMT + 1, 2], objSheet.Cells[BeginMT + 1, LastCol + 3]].Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                                objSheet.Range[objSheet.Cells[BeginMT + 1, 2], objSheet.Cells[BeginMT + 1, LastCol + 3]].Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick;

                                for (int xi = BeginMT + 1; xi <= LastRow; xi++)
                                {
                                    if (blankCol <= LastCol)
                                    {
                                        objSheet.Range[objSheet.Cells[xi, blankCol], objSheet.Cells[xi, LastCol]].Merge();
                                        objSheet.Range[objSheet.Cells[xi, blankCol], objSheet.Cells[xi, LastCol]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                                        objSheet.Cells[xi, blankCol].Font.Size = 13;

                                        //objSheet.Range[objSheet.Cells[xi, 2], objSheet.Cells[xi, 10]].Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                                        //objSheet.Range[objSheet.Cells[xi, 2], objSheet.Cells[xi, 10]].Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick;

                                    }
                                }

                                //**** Material Composition ****
                                LastRow += 6;
                                BeginMTComp = LastRow;
                                if (dtMTComposition.Rows.Count > 0)
                                {
                                    int chkRow = 0;
                                    foreach (DataRow drComp in dtMTComposition.Rows)
                                    {
                                        if (chkRow > 2)
                                        {
                                            objSheet.Rows[BeginMTComp].Insert();
                                            objSheet.Range[objSheet.Cells[BeginMTComp - 1, 1], objSheet.Cells[BeginMTComp, 1]].Merge();
                                            objSheet.Range[objSheet.Cells[BeginMTComp, 2], objSheet.Cells[BeginMTComp, LastCol + 3]].Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                        }
                                        objSheet.Cells[BeginMTComp, 2] = drComp[0].ToString();
                                        objSheet.Cells[BeginMTComp, LastCol + 1] = drComp[1].ToString();
                                        BeginMTComp++;
                                        LastRow++;
                                        chkRow++;
                                    }
                                    LastRow--;
                                }
                            }

                            pbcEXPORT.PerformStep();
                            pbcEXPORT.Update();

                            sbSQL.Clear();
                            sbSQL.Append("SELECT DISTINCT DT.OIDITEM, DT.Code, DT.Description, DT.MaterialType, PS.SizeName + ' (' + PC.ColorName + ')' AS SizeColor, PC.OIDCOLOR, PS.OIDSIZE, ");
                            sbSQL.Append("       ISNULL((SELECT GarmentParts + ', ' AS 'data()' FROM(SELECT xGP.OIDGParts, xGP.GarmentParts FROM SMPLRequestFabricParts AS xSFBP INNER JOIN GarmentParts AS xGP ON xSFBP.OIDGParts = xGP.OIDGParts AND xSFBP.OIDSMPLFB = SFB.OIDSMPLFB) AS XB FOR XML PATH('')), '')  AS FabricParts ");
                            sbSQL.Append("FROM   SMPLQuantityRequired AS SQR INNER JOIN ");
                            sbSQL.Append("       ProductColor AS PC ON SQR.OIDCOLOR = PC.OIDCOLOR INNER JOIN ");
                            sbSQL.Append("       ProductSize AS PS ON SQR.OIDSIZE = PS.OIDSIZE INNER JOIN ");
                            sbSQL.Append("       SMPLRequestFabric AS SFB ON SQR.OIDSMPLDT = SFB.OIDSMPLDT INNER JOIN ");
                            sbSQL.Append("       (SELECT DISTINCT SRFB.OIDITEM, ITM.Code, ITM.Description, ITM.MaterialType ");
                            sbSQL.Append("        FROM   SMPLRequestFabric AS SRFB INNER JOIN ");
                            sbSQL.Append("               Items AS ITM ON SRFB.OIDITEM = ITM.OIDITEM ");
                            sbSQL.Append("        WHERE  (SRFB.OIDSMPLDT IN ");
                            sbSQL.Append("                  (SELECT OIDSMPLDT ");
                            sbSQL.Append("                   FROM   SMPLQuantityRequired ");
                            sbSQL.Append("                   WHERE  (OIDSMPL = '" + SMPL[0] + "')))) AS DT ON (SQR.OIDSMPL = '" + SMPL[0] + "') AND (SFB.OIDITEM = DT.OIDITEM) ");
                            sbSQL.Append("ORDER BY DT.OIDITEM, PC.OIDCOLOR, PS.OIDSIZE ");
                            DataTable dtITEM = DBC.DBQuery(sbSQL.ToString()).getDataTable();
                            if (dtITEM.Rows.Count > 0)
                            {
                                string chkITEM = "";
                                StringBuilder sbITEM = new StringBuilder();
                                int runLoop = 0;
                                foreach (DataRow drITEM in dtITEM.Rows)
                                {
                                    string ID = drITEM["OIDITEM"].ToString();
                                    string Code = drITEM["Code"].ToString();
                                    string Description = drITEM["Description"].ToString();
                                    string MaterialType = drITEM["MaterialType"].ToString();
                                    string SizeColor = drITEM["SizeColor"].ToString();
                                    string FabricParts = drITEM["FabricParts"].ToString().Trim();
                                    FabricParts = FabricParts.Length > 0 ? FabricParts.Substring(0, FabricParts.Length - 1) : "";

                                    if (chkITEM != ID)
                                    {
                                        chkITEM = ID;
                                        if (runLoop > 0)
                                            sbITEM.Append("\n");

                                        if (MaterialType != "8")
                                            sbITEM.Append(Code + " : " + Description + "\n");
                                        else
                                            sbITEM.Append(Description + "\n");

                                    }
                                    sbITEM.Append(" ※ " + SizeColor + " -> " + FabricParts + "\n");

                                    runLoop++;
                                }

                                if (sbITEM.Length > 0)
                                {
                                    objSheet.Cells[11, LastCol + 1] = "FABRIC CODE";
                                    objSheet.Cells[11, LastCol + 1].Font.Size = 11;
                                    objSheet.Range[objSheet.Cells[12, LastCol + 1], objSheet.Cells[BeginMT, LastCol + 3]].Merge();
                                    objSheet.Range[objSheet.Cells[12 + 1, LastCol + 1], objSheet.Cells[BeginMT, LastCol + 3]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                                    objSheet.Range[objSheet.Cells[12 + 1, LastCol + 1], objSheet.Cells[BeginMT, LastCol + 3]].verticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
                                    objSheet.Cells[12, LastCol + 1] = sbITEM.ToString();
                                }
                            }


                            objWorkBook.SaveAs(sFilePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            objApp.Workbooks.Close();
                            chkExcel = true;
                            
                        }
                        catch (Exception)
                        {
                            //Error Alert
                            chkExcel = false;
                        }
                        finally
                        {
                            objApp.Quit();
                            objWorkBook = null;
                            objApp = null;
                        }

                        pbcEXPORT.PerformStep();
                        pbcEXPORT.Update();

                        if (chkExcel == true)
                        {
                            System.Diagnostics.Process.Start(sFilePath);
                            pbcEXPORT.PerformStep();
                            pbcEXPORT.Update();
                        }
                        //****** END EXPORT *******
                    }
                    else
                    {
                        FUNCT.msgError("ไม่พบข้อมูลเอกสาร Sample Request: " + SMPLNo);
                    }
                    layoutControlItem120.Text = "Status ..";
                    layoutControlItem120.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                }

            }

        }

         private void simpleButton2_Click(object sender, EventArgs e)
        {
            var frm = new DEV01_M04(this.DBC, UserLogin.OIDUser);
            frm.ShowDialog(this);
        }

        private void simpleButton3_Click(object sender, EventArgs e)
        {
            var frm = new DEV01_M09(this.DBC, UserLogin.OIDUser);
            frm.ShowDialog(this);
        }

        private void simpleButton4_Click(object sender, EventArgs e)
        {
            var frm = new DEV01_M11(this.DBC, glCategoryDivision_Main.EditValue.ToString(), UserLogin.OIDUser);
            frm.ShowDialog(this);
        }

        private void tabbedControlGroup1_SelectedPageChanged(object sender, DevExpress.XtraLayout.LayoutTabPageChangedEventArgs e)
        {
            if (tabbedControlGroup1.SelectedTabPage == layoutControlGroup1) //LIST
            {
                bbiNew.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
                bbiEdit.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
                bbiSave.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
                bbiDelete.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
                bbiRefresh.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;

                ribbonPageGroup2.Visible = true;

                if (chkReadWrite == 1)
                    rpgManage.Visible = true;

                if (gridView1.SelectedRowsCount != 0)
                {
                    GridView gv = gridView1;
                    int[] selectedRowHandles = gridView1.GetSelectedRows();
                    if (selectedRowHandles.Length > 0)
                    {
                        bbiPrint.Enabled = true;
                        bbiPrintPDF.Enabled = true;
                        bbiCLONE.Enabled = true;

                        string OIDUSER = gv.GetFocusedRowCellValue("ByCreated").ToString();
                        if (UserLogin.OIDUser.ToString() == OIDUSER)
                        {
                            bbiUPDATE.Enabled = true;
                            bbiREVISE.Enabled = true;
                            //bbiDELBILL.Enabled = true;

                            string SMPLStatus = gv.GetFocusedRowCellValue("SMPLStatus").ToString();
                            if (SMPLStatus == "0")
                                bbiDELBILL.Enabled = false;
                            else
                                bbiDELBILL.Enabled = true;
                        }
                        else
                        {
                            bbiUPDATE.Enabled = false;
                            bbiREVISE.Enabled = false;
                            bbiDELBILL.Enabled = false;
                        }
                    }
                    else
                    {
                        bbiPrint.Enabled = false;
                        bbiPrintPDF.Enabled = false;
                        bbiUPDATE.Enabled = false;
                        bbiREVISE.Enabled = false;
                        bbiCLONE.Enabled = false;
                        bbiDELBILL.Enabled = false;
                    }
                }
                else
                {
                    bbiPrint.Enabled = false;
                    bbiPrintPDF.Enabled = false;
                    bbiUPDATE.Enabled = false;
                    bbiREVISE.Enabled = false;
                    bbiCLONE.Enabled = false;
                    bbiDELBILL.Enabled = false;
                }
            }
            else if (tabbedControlGroup1.SelectedTabPage == layoutControlGroup2) //MAIN
            {
                bbiNew.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
                if (lblStatus.Text.Trim() == "Read Only SMPL")
                {
                    bbiSave.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;

                    if (txtSMPLNo.Text.Trim() != "")
                    {
                        if (chkReadWrite == 1)
                            rpgManage.Visible = true;

                        bbiCLONE.Enabled = true;

                        string OIDUSER = glueCreateBy_Main.EditValue.ToString();
                        if (UserLogin.OIDUser.ToString() == OIDUSER)
                        {
                            bbiUPDATE.Enabled = true;
                            bbiREVISE.Enabled = true;
                            //bbiDELBILL.Enabled = true;

                            string SMPLStatus = DBC.DBQuery("SELECT TOP (1) SMPLStatus FROM SMPLRequest WHERE (SMPLNo = N'" + txtSMPLNo.Text.Trim() + "')").getString();
                            if (SMPLStatus == "0" || SMPLStatus == "")
                                bbiDELBILL.Enabled = false;
                            else
                                bbiDELBILL.Enabled = true;
                        }
                        else
                        {
                            bbiUPDATE.Enabled = false;
                            bbiREVISE.Enabled = false;
                            bbiDELBILL.Enabled = false;
                        }
                    }
                }
                else
                {
                    bbiSave.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
                    rpgManage.Visible = false;
                }
                bbiEdit.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
                bbiDelete.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
                bbiRefresh.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;

                ribbonPageGroup2.Visible = false;
                
                //ribbonPageGroup5.Visible = false;
            }
            else if (tabbedControlGroup1.SelectedTabPage == layoutControlGroup8) //FABRIC
            {
                bbiNew.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
                bbiEdit.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
                if (lblStatus.Text.Trim() == "Read Only SMPL")
                {
                    bbiSave.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;

                    if (txtSMPLNo.Text.Trim() != "")
                    {
                        if (chkReadWrite == 1)
                            rpgManage.Visible = true;

                        bbiCLONE.Enabled = true;

                        string OIDUSER = glueCreateBy_Main.EditValue.ToString();
                        if (UserLogin.OIDUser.ToString() == OIDUSER)
                        {
                            bbiUPDATE.Enabled = true;
                            bbiREVISE.Enabled = true;
                            //bbiDELBILL.Enabled = true;

                            string SMPLStatus = DBC.DBQuery("SELECT TOP (1) SMPLStatus FROM SMPLRequest WHERE (SMPLNo = N'" + txtSMPLNo.Text.Trim() + "')").getString();
                            if (SMPLStatus == "0" || SMPLStatus == "")
                                bbiDELBILL.Enabled = false;
                            else
                                bbiDELBILL.Enabled = true;
                        }
                        else
                        {
                            bbiUPDATE.Enabled = false;
                            bbiREVISE.Enabled = false;
                            bbiDELBILL.Enabled = false;
                        }
                    }
                }
                else
                {
                    bbiSave.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
                    rpgManage.Visible = false;
                }
                bbiDelete.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
                bbiRefresh.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;

                ribbonPageGroup2.Visible = false;
                //rpgManage.Visible = false;
                //ribbonPageGroup5.Visible = false;
            }
            else if (tabbedControlGroup1.SelectedTabPage == layoutControlGroup12) //METERIAL
            {
                bbiNew.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
                bbiEdit.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
                if (lblStatus.Text.Trim() == "Read Only SMPL")
                {
                    bbiSave.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;

                    if (txtSMPLNo.Text.Trim() != "")
                    {
                        if (chkReadWrite == 1)
                            rpgManage.Visible = true;

                        bbiCLONE.Enabled = true;

                        string OIDUSER = glueCreateBy_Main.EditValue.ToString();
                        if (UserLogin.OIDUser.ToString() == OIDUSER)
                        {
                            bbiUPDATE.Enabled = true;
                            bbiREVISE.Enabled = true;
                            //bbiDELBILL.Enabled = true;

                            string SMPLStatus = DBC.DBQuery("SELECT TOP (1) SMPLStatus FROM SMPLRequest WHERE (SMPLNo = N'" + txtSMPLNo.Text.Trim() + "')").getString();
                            if (SMPLStatus == "0" || SMPLStatus == "")
                                bbiDELBILL.Enabled = false;
                            else
                                bbiDELBILL.Enabled = true;
                        }
                        else
                        {
                            bbiUPDATE.Enabled = false;
                            bbiREVISE.Enabled = false;
                            bbiDELBILL.Enabled = false;
                        }
                    }
                }
                else
                {
                    bbiSave.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
                    rpgManage.Visible = false;
                }
                bbiDelete.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
                bbiRefresh.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;

                ribbonPageGroup2.Visible = false;
                //rpgManage.Visible = false;
                //ribbonPageGroup5.Visible = false;
            }
 
        }

        private void btnGenSMPLNo_Click(object sender, EventArgs e)
        {
            DEV01 frmD01 = new DEV01(txtSMPLNo.Text, UserLogin);
            frmD01.Show();

        }

        private void getcell(string cellName)
        {
            gridView1.GetFocusedRowCellValue(cellName).ToString();
        }
        private void gridView1_DoubleClick(object sender, EventArgs e)
        {
            DXMouseEventArgs ea = e as DXMouseEventArgs;
            GridView view       = sender as GridView;
            GridHitInfo info = view.CalcHitInfo(ea.Location);
            if (info.InRow || info.InRowCell)
            {
                GridView gv = gridView1;
                string SMPLNo = gv.GetFocusedRowCellValue("SMPL No.").ToString();
                LoadSampleRequestDocument(SMPLNo, "READ-ONLY");

                //SetReadOnly();

                btnGenSMPLNo.Enabled = true;
                tabbedControlGroup1.SelectedTabPage = layoutControlGroup2;

            }
        }

        private void glSeason_Main_EditValueChanged(object sender, EventArgs e)
        {
            // Function Auto Save :: ไม่เอาจ้าาา
            /*ถ้ามีการเลือก Season ให้ทำการเช็คค่า 1.SaleSection, 2.ReferenceNo, 3.Season | ถ้ามีข้อมูลใน 3 รายการนี้ และ สถารนะเป็น New_Main ให้ทำการ AutoSave เป็น OID ใหม่ในตาราง*/
            //string Season = glSeason_Main.EditValue.ToString();
            //MessageBox.Show(val);

            //string SaleSection  = glSaleSection_Main.Text.ToString();
            //string ReferenceNo  = txtReferenceNo_Main.Text.ToString();
            //string Season       = glSeason_Main.EditValue.ToString();

            //if (SaleSection != "" && ReferenceNo != "" && Season != "")
            //{
            //    //btnGenSMPLNo.Enabled = true;
            //    //MessageBox.Show("Auto Save Success!");
            //}

            slCustomer_Main.Focus();
        }

        private void btnOpenImg_Main_Click(object sender, EventArgs e)
        {
            openFile_Image(xtraOpenFileDialog1,txtPictureFile_Main,picMain);
        }

        //private void updateMain()
        //{
        //    if (FUNCT.msgQuiz("Update Sample Request ?") == true)
        //    {
        //        GridView gv = gvQtyRequired;
        //        /*Test Save SMPLQuantityRequired*/
        //        //Loop Row Item in GL2 And CheckDate PrepareData And Save to tbl SMPLQuantityRequired   \
        //        string Season = glSeason_Main.Text.ToString().Trim().Replace("'", "''");
        //        if (Season == "") { FUNCT.msgWarning("Please select season!"); glSeason_Main.Focus(); return; }
        //        string SaleSection = glSaleSection_Main.EditValue.ToString();
        //        if (SaleSection == "") { FUNCT.msgWarning("Please select sales-section!"); glSaleSection_Main.Focus(); return; }
        //        string Unit = slueUnit.EditValue.ToString();
        //        if (Unit == "") { FUNCT.msgWarning("Please select unit."); slueUnit.Focus(); return; }
        //        int chk_i = 0;
        //        for (int ii = 0; ii < gv.DataRowCount; ii++)
        //        {
        //            //string No = gvQtyRequired.GetRowCellValue(ii, "No").ToString();
        //            if (ct.chkCell_isnull(gv, "Color", 0, "Please select color") == true) { return; }
        //            else if (ct.chkCell_isnull(gv, "Size", 1, "Please select size.") == true) { return; }
        //            // else if (ct.chkCell_isnull(gv, "Unit", 4, "เลือก Unit ด้วยสิคร๊าบ! ขอร้องหละ") == true) { return; }
        //            else if (Convert.ToInt32(gv.GetRowCellValue(ii, "Quantity").ToString()) <= 0)
        //            {
        //                ct.showInfoMessage("Enter a quantity of 1 or more.");
        //                gv.FocusedColumn = gv.VisibleColumns[2];
        //                gv.ShowEditor();
        //                return;
        //            }
        //            else
        //            {
        //                string Color = gvQtyRequired.GetRowCellValue(ii, "Color").ToString();
        //                string Size = gvQtyRequired.GetRowCellValue(ii, "Size").ToString();
        //                string Quantity = gvQtyRequired.GetRowCellValue(ii, "Quantity").ToString();
        //                //string Unit     = gvQtyRequired.GetRowCellValue(ii, "Unit").ToString();

        //                chk_i++;
        //                Console.WriteLine(Color + "," + Size + "," + Quantity + "," + Unit);
        //            }
        //            //MessageBox.Show(Color);
        //        }
        //        if (chk_i == 0)
        //        {
        //            FUNCT.msgWarning("Please fill in the required quantity table.");
        //            gvQtyRequired.Focus();
        //            return;
        //        }

        //        /*TextEdit*/
        //        string SMPLNo               = txtSMPLNo.Text.ToString().Trim().Replace("'","''");
        //        string ReferenceNo          = txtReferenceNo_Main.EditValue.ToString().Trim().Replace("'", "''");
        //        string ContactName          = txtContactName_Main.EditValue.ToString().Trim().Replace("'", "''");
        //        string SMPLItem             = txtSMPLItemNo_Main.EditValue.ToString().Trim().Replace("'", "''");
        //        string ModelName            = txtModelName_Main.EditValue.ToString().Trim().Replace("'", "''");
        //        string SMPLPatternNo        = txtSMPLPatternNo_Main.EditValue.ToString().Trim().Replace("'", "''");
        //        string Situation            = txtSituation_Main.EditValue.ToString().Trim().Replace("'", "''");
        //        string StateArrangements    = txtStateArrangments_Main.EditValue.ToString().Trim().Replace("'", "''");
        //        string PictureFile          = txtPictureFile_Main.Text.ToString().Trim().Replace("'", "''");
        //        string UpdatedBy            = UserLogin.OIDUser.ToString() != "" ? UserLogin.OIDUser.ToString() : "0"; //UpdateBy User Login
        //        string UpdatedDate          = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

        //        ///*GridLookup*/
        //        string OIDBranch    = glBranch_Main.EditValue.ToString();
        //        string OIDCATEGORY  = glCategoryDivision_Main.EditValue.ToString();

        //        ///*SearchLookup*/
        //        string OIDCUST  = slCustomer_Main.EditValue.ToString();
        //        string OIDSTYLE = slStyleName_Main.EditValue.ToString();

        //        ///*RadioGroup*/
        //        int SpecificationSize = Convert.ToInt32(radioGroup1.EditValue.ToString());
        //        int UseFor = Convert.ToInt32(radioGroup2.EditValue.ToString());
        //        int PatternSizeZone = Convert.ToInt32(radioGroup3.EditValue.ToString());
        //        int CustApproved = Convert.ToInt32(radioGroup4.EditValue.ToString());
        //        int ACPurRecBy = Convert.ToInt32(radioGroup5.EditValue.ToString());
        //        int FBPurRecBy = Convert.ToInt32(radioGroup6.EditValue.ToString());

        //        ///*DateTime*/
        //        string RequestDate = dtRequestDate_Main.Text.Trim() != "" ? "'" + Convert.ToDateTime(dtRequestDate_Main.EditValue.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL";
        //        string DeliveryRequest = dtDeliveryRequest_Main.Text.Trim() != "" ? "'" + Convert.ToDateTime(dtDeliveryRequest_Main.EditValue.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL";
        //        string CustApprovedDate = dtCustomerApproved_Main.Text.Trim() != "" ? "'" + Convert.ToDateTime(dtCustomerApproved_Main.EditValue.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL";
        //        string ACPurRecDate = dtACPRBy_Main.Text.Trim() != "" ? "'" + Convert.ToDateTime(dtACPRBy_Main.EditValue.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL";
        //        string FBPurRecDate = dtFBPRBy_Main.Text.Trim() != "" ? "'" + Convert.ToDateTime(dtFBPRBy_Main.EditValue.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL";

        //        // Check is null or Empty Data :: Not Null
        //        if (CustApproved == 1)
        //        {
        //            if (OIDCUST == "") { FUNCT.msgWarning("Please select customer!"); slCustomer_Main.Focus(); return; }
        //            if (OIDCATEGORY == "") { FUNCT.msgWarning("Please select category!"); glCategoryDivision_Main.Focus(); return; }
        //            if (OIDSTYLE == "") { FUNCT.msgWarning("Please select style!"); slStyleName_Main.Focus(); return; }
        //            if (OIDBranch == "") { FUNCT.msgWarning("Please select branch!"); glBranch_Main.Focus(); return; }
        //            if (ContactName == "") { FUNCT.msgWarning("Please select contact-name!"); txtContactName_Main.Focus(); return; }
        //            if (SMPLItem == "") { FUNCT.msgWarning("Please select SMPL Item No.!"); txtSMPLItemNo_Main.Focus(); return; }
        //            if (ModelName == "") { FUNCT.msgWarning("Please select model name!"); txtModelName_Main.Focus(); return; }
        //            if (SMPLPatternNo == "") { FUNCT.msgWarning("Please select SMPL Pattern No.!"); txtSMPLPatternNo_Main.Focus(); return; }
        //        }

        //        string newFileName = ct.uploadImg(txtPictureFile_Main, "FG");

        //        sql = "Update SMPLRequest Set OIDBranch="+ OIDBranch + ",RequestDate=" + RequestDate + ",SpecificationSize=" + SpecificationSize + ",ContactName='" + ContactName + "',DeliveryRequest=" + DeliveryRequest + ",UseFor=" + UseFor + ",SMPLItem='" + SMPLItem + "',ModelName='" + ModelName + "',OIDCATEGORY=" + OIDCATEGORY + ",Situation='" + Situation + "',StateArrangements='" + StateArrangements + "',CustApproved=" + CustApproved + ",CustApprovedDate=" + CustApprovedDate + ",ACPurRecBy=" + ACPurRecBy + ",ACPurRecDate=" + ACPurRecDate + ",FBPurRecBy=" + FBPurRecBy + ",FBPurRecDate=" + FBPurRecDate + " /*,PictureFile=" + newFileName + "*/ WHERE SMPLNo = '"+SMPLNo+"' ";

        //        Console.WriteLine(sql);
        //        int i = db.Query(sql, mainConn);
        //        if (i > 0)
        //        {
        //            bool chkSave = false;
        //            //Step1 : Get OIDSMPL ที่บันทึกเมื่อกี้มา
        //            //string maxOIDSMPL = db.get_oneParameter("Select MAX(OIDSMPL) as maxOID From SMPLRequest",mainConn, "maxOID");
        //            //string maxOIDSMPL = db.get_oneParameter("SELECT TOP (1) OIDSMPL FROM SMPLRequest WHERE (SMPLNo = N'" + SMPLNo + "')", mainConn, "OIDSMPL");
        //            //Step2 : Loop List of Quantity Required And Save Data to tbl SMPLQuantityRequired
        //            //string Unit = slueUnit.EditValue.ToString();

        //            string sql2 = "";
        //            string listID = "";
        //            int chkLoop = 0;
        //            //string sql2 = "DELETE FROM SMPLQuantityRequired WHERE (OIDSMPL = '" + maxOIDSMPL + "') ";
        //            for (int j = 0; j < gv.DataRowCount; j++)
        //            {
        //                string Color = gv.GetRowCellValue(j, "Color").ToString();
        //                string Size = gv.GetRowCellValue(j, "Size").ToString();
        //                string Quantity = gv.GetRowCellValue(j, "Quantity").ToString();
        //                string ID = gv.GetRowCellValue(j, "ID").ToString();

        //                if (ID != "")
        //                {
        //                    if (chkLoop > 0)
        //                        listID += ", ";
        //                    listID += "'" + ID + "'";
        //                    chkLoop++;
        //                    sql2 += "UPDATE SMPLQuantityRequired SET OIDCOLOR='" + Color + "', OIDSIZE='" + Size + "', Quantity='" + Quantity + "', OIDUnit='" + Unit + "' WHERE (OIDSMPLDT = '" + ID + "')  ";
        //                }
        //                else
        //                {
        //                    sql2 += "INSERT INTO SMPLQuantityRequired(OIDSMPL, OIDCOLOR, OIDSIZE, Quantity, OIDUnit) VALUES('" + SMPLNo + "', '" + Color + "', '" + Size + "', '" + Quantity + "', '" + Unit + "')  ";
        //                }


        //            }

        //            if (sql2 != "")
        //            {
        //                string xSQL = "";
        //                if (listID != "")
        //                    xSQL = "DELETE FROM SMPLQuantityRequired WHERE (OIDSMPL = '" + SMPLNo + "') AND (OIDSMPLDT NOT IN (" + listID + "))   ";

        //                xSQL += sql2;
        //                int i2 = db.Query(xSQL, mainConn);
        //                if (i2 > 0)
        //                {
        //                    chkSave = true;
        //                }
        //            }

        //            if (chkSave == true)
        //            {
        //                ct.showInfoMessage("Update is Successfull.");
        //                db.getGrid_SMPL(gridControl1, gridView1);
        //                newMain();
        //                //Next to TabPage Fabric
        //                //tabbedControlGroup1.SelectedTabPageIndex = 2;
        //            }
        //            else
        //            {
        //                ct.showErrorMessage("Save SMPL is Failed!");
        //                db.getGrid_SMPL(gridControl1, gridView1);
        //                //newMain();
        //            }
        //        }


        //    }
        //}

        //private void updateFabric()
        //{
        //    bool isUpdate = false;

        //    if (FUNCT.msgQuiz("Updat Fabric?") == true && txtSampleID_FB.Text.ToString() != "" && txtFabricRacordID_FB.Text.ToString() != "")
        //    {
        //        string fbid = txtFabricRacordID_FB.Text.ToString();
        //        string QDT  = db.get_oneParameter("Select OIDSMPLDT From SMPLRequestFabric Where OIDSMPLFB = "+ fbid + " ",mainConn, "OIDSMPLDT");

        //        //Not null
        //        string VendFBCode   = ct.getVal_text(txtVendorFBCode_FB); //return null
        //        string SampleLotNo  = ct.getVal_text(txtSampleLotNo_FB);
        //        string Vendor       = ct.getVal_sl(slVendor_FB);
        //        string FBColor      = ct.getVal_sl(slFBColor_FB);
        //        string FBCode       = ct.getVal_sl(slFBCode_FB);
        //        string FGColor      = ct.getVal_sl(slFGColor_FB);

        //        //Accept Null
        //        string Composition  = ct.getVal_text(txtComposition_FB);
        //        string weight       = ct.getVal_text(txtWeightFB_FB);
        //        string widthCut     = ct.getVal_text(txtWidthCuttable_FB);
        //        string price        = ct.getVal_num(txtPrice_FB);
        //        string Currency     = ct.getVal_sl(glCurrency_FB);
        //        string TotalWidth   = ct.getVal_num(txtTotalWidth_FB);
        //        string UsableWidth  = ct.getVal_num(txtUsableWidth_FB);

        //        //chkNull
        //        if (VendFBCode == "null") {         FUNCT.msgWarning("Please Key Vendor Fabric Code!"); txtVendorFBCode_FB.Focus(); return; }
        //        else if (SampleLotNo == "null") {   FUNCT.msgWarning("Please Key SampleLotNo!"); txtSampleLotNo_FB.Focus(); return; }
        //        else if (Vendor == "null") {        FUNCT.msgWarning("Please Select Vendor!"); slVendor_FB.Focus(); return; }
        //        else if (FBColor == "null") {       FUNCT.msgWarning("Please Select FBColor!"); slFBColor_FB.Focus(); return; }
        //        else if (FBCode == "null") {        FUNCT.msgWarning("Please Select FBCode!"); slFBCode_FB.Focus(); return; }
        //        else if (FGColor == "null") {       FUNCT.msgWarning("Please Select FGColor!"); slFGColor_FB.Focus(); return; }
        //        else
        //        {
        //            string sql = "Update SMPLRequestFabric set VendFBCode = "+ VendFBCode + ", SMPLotNo = "+ SampleLotNo + ", OIDVEND = "+ Vendor + ",OIDCOLOR = "+ FBColor + ",OIDITEM = "+ FBCode + " /*FGColor ต้องไป Join SMPLQuantity*/ ,Composition = "+ Composition + ", FBWeight = "+ weight + ",WidthCuttable = "+ widthCut + ",Price = "+ price + ",TotalWidth = "+ TotalWidth + ",UsableWidth = "+ UsableWidth + ",OIDCURR = "+ Currency + " ";
        //            sql += " Where OIDSMPLFB = "+ fbid + " ";
        //            Console.WriteLine(sql);
        //            db.Query(sql,mainConn);

        //            string sql2 = "Delete SMPLRequestFabricParts Where OIDSMPLFB = "+ fbid + " ";
        //            Console.WriteLine(sql2);
        //            db.Query(sql2,mainConn);

        //            ArrayList arow = ct.getList_isChecked(gridView11);
        //            if (arow.Count > 0)
        //            {
        //                try
        //                {
        //                    for (int i = 0; i < arow.Count; i++)
        //                    {
        //                        DataRow r = arow[i] as DataRow;
        //                        string OIDGpart = r["OIDGParts"].ToString();
        //                        string sql3 = "Insert Into SMPLRequestFabricParts(OIDSMPLFB,OIDSMPLDT,OIDGParts) Values("+ fbid + ","+ QDT + ","+ OIDGpart + ")";
        //                        Console.WriteLine(sql3);
        //                        int qi = db.Query(sql3,mainConn);
        //                        if (qi > 0)
        //                        {
        //                            isUpdate = true;
        //                        }
        //                    }
        //                }
        //                catch { }
        //            }
        //        }
        //        refreshFabric();
        //        db.getListofFabric(gcList_Fabric,dosetOIDSMPL);

        //        if (isUpdate == true)
        //        {
        //            ct.showInfoMessage("Update Success.");
        //        }
        //        else
        //        {
        //            ct.showErrorMessage("Can't Update. Please Contact Administrator!");
        //        }
        //    }
        //}

        //private void updateMaterials()
        //{
        //    bool statusUpdate = false;

        //    if (FUNCT.msgQuiz("Update Material?")==true)
        //    {
        //        GridView g = gridView7;
        //        //Not null :
        //        if (glWorkStation_Mat.Text.ToString() != "" || slVendor_Mat.Text.ToString() != "")
        //        {
        //            //Var Set Update
        //            string MatID        = txtMatRecordID_Mat.Text.ToString();
        //            string WorkStation  = glWorkStation_Mat.EditValue.ToString();
        //            string Vendor       = slVendor_Mat.EditValue.ToString();
        //            string vendMatCode  = ct.getVal_text(txtVendorMatCode_Mat);
        //            string Lotno        = ct.getVal_text(txtSampleLotNo_Mat);
        //            string Composition  = ct.getVal_text(txtMatComposition_Mat);
        //            //string matColor     = ct.getVal_sl(slMatColor_Mat);
        //            string matCode      = ct.getVal_sl(slMatCode_Mat);
        //            string price        = ct.getVal_num(txtPrice_Mat);
        //            string currency     = ct.getVal_sl(glCurrency_Mat);
        //            string situation    = ct.getVal_text(txtSituation_Mat);
        //            string Comment      = ct.getVal_text(txtComment_Mat);
        //            string Remark       = ct.getVal_text(txtRemark_Mat);
        //            string pathFile     = ct.getVal_text(txtPathFile_Mat); /*ยังไม่ Update อันนี้นะจ๊ะ*/

        //            // Special Var in Gridview
        //            string matColor     = (g.GetRowCellValue(0,"Color").ToString() == "") ? "null" : g.GetRowCellValue(0, "Color").ToString();
        //            string Consumption  = (g.GetRowCellValue(0, "Consumption").ToString() == "") ? "null" : g.GetRowCellValue(0, "Consumption").ToString();
        //            string Size         = g.GetRowCellValue(0, "Size").ToString(); //จะต้องไม่เป็นค่า Null แน่นอน เพราะดึงมาจาก Master

        //            // SqlUpdate
        //            string sql = "Update SMPLRequestMaterial Set OIDDEPT = " + WorkStation + ",OIDVEND = " + Vendor + ",VendMTCode=" + vendMatCode + ",SMPLotNo=" + Lotno + ",Composition=" + Composition + ",MTColor=" + matColor + ",Consumption=" + Consumption + ",OIDITEM=" + matCode + ",Price=" + price + ",OIDCURR=" + currency + ",Situation=" + situation + ",Comment=" + Comment + ",Remark=" + Remark + " ";
        //            sql += " Where OIDSMPLMT = " + txtMatRecordID_Mat.Text.ToString() + " ";

        //            // CheckUpdate รับค่าจาก Form มาเช็คใน Database 3 ตัว ถ้าตรงกันหมด = ไม่มีอะไรเปลี่ยนแปลง >> แล้วถ้ามีอันไหนไม่ตรงกัน = ให้ chkDup ก่อน Update
        //            string db_vendMatCode   = db.get_oneParameter("Select VendMTCode From SMPLRequestMaterial Where OIDSMPLMT = "+ MatID + " ",mainConn, "VendMTCode");
        //            db_vendMatCode          = (db_vendMatCode == "") ? "null" : "N'"+db_vendMatCode+"'";
        //            string db_matColor      = db.get_oneParameter("Select MTColor From SMPLRequestMaterial Where OIDSMPLMT = " + MatID + " ",mainConn, "MTColor");
        //            db_matColor             = (db_matColor == "") ? "null" : db_matColor;
        //            string db_matSize       = db.get_oneParameter("Select MTSize From SMPLRequestMaterial Where OIDSMPLMT = " + MatID + " ",mainConn, "MTSize");
        //            Console.WriteLine(db_vendMatCode+","+ vendMatCode+"\n" +db_matColor+","+ matColor+"\n" +db_matSize+","+ Size);

        //            if (db_vendMatCode == vendMatCode && db_matColor == matColor && db_matSize == Size)
        //            {
        //                //ct.showInfoMessage("Math and Normal Update");
        //                Console.WriteLine(sql);
        //                int i = db.Query(sql, mainConn);
        //                if (i > 0)
        //                {
        //                    statusUpdate = true;
        //                }
        //            }
        //            else
        //            {
        //                //ct.showInfoMessage("Not Math Some Field is Changed! > chkDuplicate");
        //                // chkDup :: VenMatCode , MatColor , MatSize
        //                string eq = (matColor == "null") ? "is" : "=";
        //                string sql_chkDup = "Select VendMTCode From SMPLRequestMaterial Where (VendMTCode = " + vendMatCode + " and MTColor " + eq + " " + matColor + " and MTSize = " + Size + ")";
        //                Console.WriteLine(sql_chkDup);
        //                if (db.get(sql_chkDup, mainConn) == true) { FUNCT.msgWarning("MatCode or MatColor or MatSize is Duplicate!"); return; }
        //                else
        //                {
        //                    Console.WriteLine(sql);
        //                    int i = db.Query(sql, mainConn);
        //                    if (i > 0)
        //                    {
        //                        statusUpdate = true;
        //                    }
        //                }
        //            }
        //        }
        //        else
        //        {
        //            FUNCT.msgWarning("Please Select WorkStation and Vendor");
        //        }

        //        // Check Update is Successfull
        //        if (statusUpdate == true)
        //        {
        //            ct.showInfoMessage("Update Completed");
        //            newMaterials();
        //        }
        //        else
        //        {
        //            ct.showErrorMessage("Update is Failed. Please Contact Administrator!");
        //        }
        //    }
        //}

        private void bbiEdit_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            //switch (currenTab)
            //{
            //    /* List of Sample , Main , Fabric , Material */
            //    case "Main": updateMain(); break;
            //    case "Fabric": updateFabric(); break;
            //    case "Material": updateMaterials(); break;
            //    default: updateMain(); break;
            //}
        }

        private void gridView3_SelectionChanged(object sender, DevExpress.Data.SelectionChangedEventArgs e)
        {
            //ArrayList rows = ct.getList_isChecked(gridView3);

            //if (rows.Count > 0)
            //{
            //    PageFBVal = true;

            //    try
            //    {
            //        // Create DataTable
            //        DataTable dt = new DataTable();

            //        // Add Columns Header into DataTable
            //        dt.Columns.Add("No", typeof(string));
            //        dt.Columns.Add("VendFBCode", typeof(string));
            //        dt.Columns.Add("Composition", typeof(string));
            //        dt.Columns.Add("FBWeight", typeof(string));
            //        dt.Columns.Add("ColorName", typeof(string));
            //        dt.Columns.Add("SMPLotNo", typeof(string));
            //        dt.Columns.Add("Supplier", typeof(string));
            //        dt.Columns.Add("NAVCode", typeof(string));
            //        dt.Columns.Add("Description", typeof(string));

            //        int listfbNo = 1;
            //        string sqlcmd = string.Empty;
            //        for (int i = 0; i < rows.Count; i++)
            //        {
            //            DataRow row = rows[i] as DataRow;
            //            string PatternNo = row["SMPLPatternNo"].ToString();
            //            string Color = row["ColorName"].ToString();
            //            string Size = row["SizeName"].ToString();
            //            string Quantity = row["Quantity"].ToString();

            //            /*Set size*/
            //            sqlcmd += "Select '" + Size + "' as SizeName Union ";

            //            /* Add to List Fabric*/
            //            dt.Rows.Add(new object[] {
            //                listfbNo++
            //                ,txtVendorFBCode_FB.Text.Trim().ToString().Replace("'","''")
            //                ,txtComposition_FB.Text.Trim().ToString().Replace("'","''")
            //                ,txtWeightFB_FB.Text.Trim().ToString().Replace("'","''")
            //                ,slFBColor_FB.Text.ToString()
            //                ,txtSampleLotNo_FB.Text.Trim().ToString().Replace("'","''")
            //                ,slVendor_FB.Text.ToString()
            //                ,slFBCode_FB.Text.ToString()
            //                ,lblDescription.Text
            //            });
            //        }
            //        gcList_Fabric.DataSource = dt;

            //        int length = sqlcmd.Length;
            //        sqlcmd = sqlcmd.Substring(0, length - 6);
            //        Console.WriteLine(sqlcmd);
            //        db.getDgv(sqlcmd, gcSize_Fabric, mainConn);
            //    }
            //    catch { }
            //}
            //else
            //{
            //    PageFBVal = false;
            //    gcSize_Fabric.DataSource = null;
            //    gcList_Fabric.DataSource = null;
            //}
        }

        private void gridView3_DoubleClick(object sender, EventArgs e)
        {
            DXMouseEventArgs ea = e as DXMouseEventArgs;
            GridView view = sender as GridView;
            GridHitInfo info = view.CalcHitInfo(ea.Location);
            if (info.InRow || info.InRowCell)
            {
                GridView gv = gridView3;
                string SMPLNo = gv.GetFocusedRowCellValue("OIDSMPL").ToString();
                string PatternNo = gv.GetFocusedRowCellValue("SMPLPatternNo").ToString();
                string ColorNo = gv.GetFocusedRowCellValue("ColorName").ToString();

                if (ColorNo != slFGColor_FB.EditValue.ToString())
                {
                    lblFBStatus.Text = "Status : New";
                    lblFBStatus.BackColor = Color.Green;
                    lblRowID.Text = "";
                }

                //MessageBox.Show(ColorNo);
                slFGColor_FB.ReadOnly = false;
                slFGColor_FB.EditValue = ColorNo;
                slFGColor_FB.ReadOnly = true;

                txtSampleID_FB.Text = SMPLNo;
                if (dtFBSize.Rows.Count > 0)
                    dtFBSize.Rows.Clear();
                DataTable dtCS = new DataTable();
                dtCS = (DataTable)gcQtyRequired.DataSource;
                if (dtCS.Rows.Count > 0)
                {
                    foreach (DataRow rCS in dtCS.Rows)
                    {
                        string Color = rCS["Color"].ToString().Trim();
                        if (Color != "")
                        {
                            string Size = rCS["Size"].ToString().Trim();
                            string Quantity = rCS["Quantity"].ToString().Trim();
                            string DT_ID = rCS["ID"].ToString().Trim();

                            if (ColorNo == Color)
                            {
                                dtFBSize.Rows.Add(Size, Quantity, DT_ID);
                            }

                        }
                    }

                    gcSize_Fabric.DataSource = dtFBSize;
                    gridView12.SelectAll();
                    gridView12.BestFitColumns();

                    //Clear Data
                    //ClearFabric();
                }
            }
        }

        private void btngetListFB_FB_Click(object sender, EventArgs e)
        {
            if (slFGColor_FB.Text != "")
            {
                if (slFBCode_FB.Text.Trim() == "")
                {
                    FUNCT.msgWarning("Please select fabric code"); slFBCode_FB.Focus(); return;
                }
                if (gridView11.SelectedRowsCount == 0)
                {
                    FUNCT.msgWarning("Please select fabric parts"); gridView11.Focus(); return;
                }
                else
                {
                    //****** Check Duplicate Value ******
                    string chkFGColor = slFGColor_FB.Text.Trim() == "" ? "" : slFGColor_FB.EditValue.ToString();
                    string chkFBCode = slFBCode_FB.Text.Trim();
                    DataTable dtFB = (DataTable)gcList_Fabric.DataSource;
                    if (dtFB != null)
                    {
                        int runLoop = 0;
                        foreach (DataRow drFB in dtFB.Rows)
                        {
                            string xFGColor = drFB["ColorName"].ToString();
                            string xFBCode = drFB["ChkFBCode"].ToString();

                            //MessageBox.Show("chkFGColor:" + chkFGColor + "-xFGColor:" + xFGColor + ", chkFBCode:" + chkFBCode + "-xFBCode:" + xFBCode);

                            if (lblFBStatus.Text.Replace("Status : ", "") == "New")
                            {
                                if (xFGColor == chkFGColor && xFBCode == chkFBCode)
                                {
                                    if (FUNCT.msgQuiz("Found duplicate fabric code. Do you want to update ?\nพบโค้ดผ้ามีอยู่แล้วในตาราง ต้องการจะแทนที่ข้อมูลเดิมหรือไม่") == true)
                                    {
                                        lblRowID.Text = runLoop.ToString();
                                        lblFBStatus.Text = "Status : Update";
                                        lblFBStatus.BackColor = Color.Navy;
                                        break;
                                    }
                                    else
                                    {
                                        return;
                                    }
                                }
                            }
                            else if (lblFBStatus.Text.Replace("Status : ", "") == "Update")
                            {
                                if (xFGColor == chkFGColor && xFBCode == chkFBCode && runLoop.ToString() != lblRowID.Text)
                                {
                                    FUNCT.msgError("Cannot update. Because the fabric code is duplicated with other rows ?\nไม่สามารถแทนที่ข้อมูลได้ เนื่องจากพบโค้ดผ้าซ้ำกับแถวอื่นในตาราง");
                                    return;
                                }
                            }

                            runLoop++;
                        }
                    }
                    //***********************************



                    if (lblFBStatus.Text.Replace("Status : ", "") == "New")
                    {
                        bool chkDUP = false;
                        DataTable dtFabricX = (DataTable)gcList_Fabric.DataSource;
                        if (dtFabricX == null)
                            dtFabricX = dtFabric;


                        if (chkDUP == false)
                        {
                            string strFBID = "";
                            string strFBParts = "";

                            int[] selectedRowHandles = gridView11.GetSelectedRows();
                            if (selectedRowHandles.Length > 0)
                            {
                                int xLoop = 0;
                                gridView11.FocusedRowHandle = selectedRowHandles[0];
                                for (int i = 0; i < selectedRowHandles.Length; i++)
                                {
                                    string PartsID = gridView11.GetRowCellDisplayText(selectedRowHandles[i], "ID");
                                    string GarmentParts = gridView11.GetRowCellDisplayText(selectedRowHandles[i], "GarmentParts");

                                    if (xLoop > 0)
                                    {
                                        strFBID += ",";
                                        strFBParts += ", ";
                                    }

                                    strFBID += PartsID;
                                    strFBParts += GarmentParts;

                                    xLoop++;
                                }
                            }

                            string Price = txtPrice_FB.Text.Trim() == "" ? "0" : txtPrice_FB.Text.Trim();
                            string TotalWidth = txtTotalWidth_FB.Text.Trim() == "" ? "0" : txtTotalWidth_FB.Text.Trim();
                            string UsableWidth = txtUsableWidth_FB.Text.Trim() == "" ? "0" : txtUsableWidth_FB.Text.Trim();
                            //MessageBox.Show("2: " + dtFabric.Rows.Count.ToString());

                            string FGColor = slFGColor_FB.Text.Trim() == "" ? "" : slFGColor_FB.EditValue.ToString();
                            string Vendor = slVendor_FB.Text.Trim() == "" ? "" : slVendor_FB.EditValue.ToString();
                            string FBColor = slFBColor_FB.Text.Trim() == "" ? "" : slFBColor_FB.EditValue.ToString();
                            string FBCode = slFBCode_FB.Text.Trim() == "" ? "" : slFBCode_FB.EditValue.ToString();
                            string Currency = glCurrency_FB.Text.Trim() == "" ? "" : glCurrency_FB.EditValue.ToString();

                            dtFabricX.Rows.Add(
                                FGColor
                                , txtVendorFBCode_FB.Text.Trim().ToString().Replace("'", "''")
                                , txtSampleLotNo_FB.Text.Trim().ToString().Replace("'", "''")
                                , Vendor
                                , FBColor
                                , FBCode
                                , lblDescription.Text
                                , txtComposition_FB.Text.Trim().ToString().Replace("'", "''")
                                , txtWeightFB_FB.Text.Trim().ToString().Replace("'", "''")
                                , txtWidthCuttable_FB.Text.Trim().ToString().Replace("'", "''")
                                , Price
                                , Currency
                                , TotalWidth
                                , UsableWidth
                                , txtImgUpload_FB.Text.Trim()
                                , slFBCode_FB.Text.Trim()
                                , strFBID
                                , strFBParts
                                , ""
                                , txeRemark_FB.Text.Trim()
                            );
                            gcList_Fabric.DataSource = dtFabricX;
                            gridView4.BestFitColumns();

                            //Clear Data
                            if (FUNCT.msgQuiz("Do you want to clear data in field above?\nต้องการเคลียร์ข้อมูลด้านบนหรือไม่") == true)
                            {
                                ClearFabric();
                            }
                            else
                            {
                                lblRowID.Text = "";
                                lblFBStatus.Text = "Status : New";
                                lblFBStatus.BackColor = Color.Green;
                                sbDeleteRow.Visible = false;
                            }

                        }
                        else
                        {
                            FUNCT.msgWarning("Fabric Code is duplicate. Please change."); slFBCode_FB.Focus(); return;
                        }
                    }
                    else if (lblFBStatus.Text.Replace("Status : ", "") == "Update")
                    {
                        if (FUNCT.msgQuiz("Confirm update this data ?") == true)
                        {
                            bool chkDUP = false;
                            //copy datasource to datatable
                            dtFabric = (DataTable)gcList_Fabric.DataSource;
                            //chk fabric code duplicate
                            //int RowID = 0;
                            //foreach (DataRow rDup in dtFabric.Rows)
                            //{
                            //    if (RowID != Convert.ToInt32(lblRowID.Text))
                            //    {
                            //        string ColorName = rDup["ColorName"].ToString();
                            //        string FabricCode = rDup["FabricCode"].ToString();
                            //        if (ColorName == slFGColor_FB.EditValue.ToString() && FabricCode == slFBCode_FB.EditValue.ToString())
                            //        {
                            //            chkDUP = true;
                            //            break;
                            //        }
                            //    }
                            //    RowID++;
                            //}


                            if (chkDUP == false)
                            {
                                string strFBID = "";
                                string strFBParts = "";

                                int[] selectedRowHandles = gridView11.GetSelectedRows();
                                if (selectedRowHandles.Length > 0)
                                {
                                    int xLoop = 0;
                                    gridView11.FocusedRowHandle = selectedRowHandles[0];
                                    for (int i = 0; i < selectedRowHandles.Length; i++)
                                    {
                                        string PartsID = gridView11.GetRowCellDisplayText(selectedRowHandles[i], "ID");
                                        string GarmentParts = gridView11.GetRowCellDisplayText(selectedRowHandles[i], "GarmentParts");

                                        if (xLoop > 0)
                                        {
                                            strFBID += ",";
                                            strFBParts += ", ";
                                        }

                                        strFBID += PartsID;
                                        strFBParts += GarmentParts;

                                        xLoop++;
                                    }
                                }

                                string FGColor = slFGColor_FB.Text.Trim() == "" ? "" : slFGColor_FB.EditValue.ToString();
                                string Vendor = slVendor_FB.Text.Trim() == "" ? "" : slVendor_FB.EditValue.ToString();
                                string FBColor = slFBColor_FB.Text.Trim() == "" ? "" : slFBColor_FB.EditValue.ToString();
                                string FBCode = slFBCode_FB.Text.Trim() == "" ? "" : slFBCode_FB.EditValue.ToString();
                                string Currency = glCurrency_FB.Text.Trim() == "" ? "" : glCurrency_FB.EditValue.ToString();

                                dtFabric.Rows[Convert.ToInt32(lblRowID.Text)]["ColorName"] = FGColor;
                                dtFabric.Rows[Convert.ToInt32(lblRowID.Text)]["VendorFBCode"] = txtVendorFBCode_FB.Text.Trim().ToString().Replace("'", "''");
                                dtFabric.Rows[Convert.ToInt32(lblRowID.Text)]["SMPLotNo"] = txtSampleLotNo_FB.Text.Trim().ToString().Replace("'", "''");
                                dtFabric.Rows[Convert.ToInt32(lblRowID.Text)]["Supplier"] = Vendor;
                                dtFabric.Rows[Convert.ToInt32(lblRowID.Text)]["FabricColor"] = FBColor;
                                dtFabric.Rows[Convert.ToInt32(lblRowID.Text)]["FabricCode"] = FBCode;
                                dtFabric.Rows[Convert.ToInt32(lblRowID.Text)]["Description"] = lblDescription.Text;
                                dtFabric.Rows[Convert.ToInt32(lblRowID.Text)]["Composition"] = txtComposition_FB.Text.Trim().ToString().Replace("'", "''");
                                dtFabric.Rows[Convert.ToInt32(lblRowID.Text)]["FBWeight"] = txtWeightFB_FB.Text.Trim().ToString().Replace("'", "''");
                                dtFabric.Rows[Convert.ToInt32(lblRowID.Text)]["WidthCut"] = txtWidthCuttable_FB.Text.Trim().ToString().Replace("'", "''");
                                dtFabric.Rows[Convert.ToInt32(lblRowID.Text)]["Price"] = txtPrice_FB.Text.Trim() == "" ? "0" : txtPrice_FB.Text.Trim();
                                dtFabric.Rows[Convert.ToInt32(lblRowID.Text)]["Currency"] = Currency;
                                dtFabric.Rows[Convert.ToInt32(lblRowID.Text)]["TTWidth"] = txtTotalWidth_FB.Text.Trim() == "" ? "0" : txtTotalWidth_FB.Text.Trim();
                                dtFabric.Rows[Convert.ToInt32(lblRowID.Text)]["UsableWidth"] = txtUsableWidth_FB.Text.Trim() == "" ? "0" : txtUsableWidth_FB.Text.Trim();
                                dtFabric.Rows[Convert.ToInt32(lblRowID.Text)]["PicFile"] = txtImgUpload_FB.Text.Trim();
                                dtFabric.Rows[Convert.ToInt32(lblRowID.Text)]["ChkFBCode"] = slFBCode_FB.Text.Trim();
                                dtFabric.Rows[Convert.ToInt32(lblRowID.Text)]["FBPartsID"] = strFBID;
                                dtFabric.Rows[Convert.ToInt32(lblRowID.Text)]["FBPartsName"] = strFBParts;
                                dtFabric.Rows[Convert.ToInt32(lblRowID.Text)]["Remark"] = txeRemark_FB.Text.Trim();

                                gcList_Fabric.DataSource = dtFabric;
                                gridView4.BestFitColumns();

                                //Clear Data
                                if (FUNCT.msgQuiz("Do you want to clear data in field above?\nต้องการเคลียร์ข้อมูลด้านบนหรือไม่") == true)
                                {
                                    ClearFabric();
                                }
                                else
                                {
                                    lblRowID.Text = "";
                                    lblFBStatus.Text = "Status : New";
                                    lblFBStatus.BackColor = Color.Green;
                                    sbDeleteRow.Visible = false;
                                }
                            }
                            else
                            {
                                FUNCT.msgWarning("Fabric Code is duplicate. Please change."); slFBCode_FB.Focus(); return;
                            }
                        }
                    }
                }
            }
            else
            {
                FUNCT.msgWarning("Please select color from list of sample table by double click.\nกรุณาเลือกสีจากตาราง List of Sample โดยการดับเบิ้ลคลิ๊ก"); gridControl3.Focus(); return;
            }


        }

        private void simpleButton5_Click(object sender, EventArgs e)
        {
            openFile_Image(xtraOpenFileDialog1,txtImgUpload_FB,picUpload_FB);
        }

        public ArrayList getList_isChecked(GridView gvName)
        {
            ArrayList rows = new ArrayList();
            // Add the selected rows to the list.
            Int32[] selectedRowHandles = gvName.GetSelectedRows();  //getSelectedRow
            for (int i = 0; i < selectedRowHandles.Length; i++)     //Loop SelectedRow
            {
                int selectedRowHandle = selectedRowHandles[i];
                if (selectedRowHandle >= 0)                         //if getSelectedRow >= 0
                {
                    rows.Add(gvName.GetDataRow(selectedRowHandle)); //Add SelectedRow to ArrayList
                }
            }
            return rows;
        }

        private void gridView6_SelectionChanged(object sender, DevExpress.Data.SelectionChangedEventArgs e)
        {
            txtSampleID_Mat.Text = lblID.Text;

            txtMatRecordID_Mat.Text = "";
            lblMTStatus.Text = "Status : New";
            lblMTStatus.BackColor = Color.Green;
            layoutControlItem101.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;

            ArrayList row = getList_isChecked(gridView6);
            if (row.Count > 0)
            {
                // Create DataTable
                DataTable dt = new DataTable();

                // Add Columns Header into DataTable
                dt.Columns.Add("Color", typeof(string));
                dt.Columns.Add("Size", typeof(string));
                dt.Columns.Add("Consumption", typeof(string));

                for (int i = 0; i < row.Count; i++)
                {
                    DataRow r = row[i] as DataRow;
                    /* Add to List Mat*/
                    dt.Rows.Add(new object[] {
                        r["ColorName"].ToString()
                        , r["SizeName"].ToString()
                        , "0"
                    });
                }

                gridControl7.DataSource = dt;
                gridView7.BestFitColumns();
 
            }
            else
            {
                gridControl7.DataSource = null;
                gridView7.BestFitColumns();
            }
        }

        private void slVendor_Mat_EditValueChanged(object sender, EventArgs e)
        {
            txtVendName_Mat.Text = "";
            if (slVendor_Mat.Text.Trim() != "")
            {
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT Name FROM Vendor WHERE (OIDVEND = '" + slVendor_Mat.EditValue.ToString() + "') ");
                txtVendName_Mat.Text = DBC.DBQuery(sbSQL.ToString()).getString();
            }

            txtVendorMatCode_Mat.Focus();
        }

        private void btnUploadMat_Click(object sender, EventArgs e)
        {
            openFile_Image(xtraOpenFileDialog1,txtPathFile_Mat,picMat);
        }

        private void btnGettoLlist_Mat_Click(object sender, EventArgs e)
        {
            if (glWorkStation_Mat.Text.Trim() == "")
            {
                FUNCT.msgWarning("Please select work status."); glWorkStation_Mat.Focus(); return;
            }
            else if (slVendor_Mat.Text.Trim() == "")
            {
                FUNCT.msgWarning("Please select vendor"); slVendor_Mat.Focus(); return;
            }
            else if (gridView7.RowCount == 0)
            {
                FUNCT.msgWarning("Please select sample"); gridControl6.Focus(); return;
            }
            else if (slMatCode_Mat.Text.Trim() == "")
            {
                FUNCT.msgWarning("Please select material code"); slMatCode_Mat.Focus(); return;
            }
            else
            {
                //****** Check Duplicate Value ******
                //string chkFGColor = slFGColor_FB.Text.Trim() == "" ? "" : slFGColor_FB.EditValue.ToString();
                //string chkFBCode = slFBCode_FB.Text.Trim();
                //DataTable dtMTR = (DataTable)gridControl8.DataSource;
                //if (dtMTR != null)
                //{
                //    int runLoop = 0;
                //    foreach (DataRow drMTR in dtMTR.Rows)
                //    {
                //        string xFGColor = drMTR["ColorName"].ToString();
                //        string xFBCode = drMTR["ChkFBCode"].ToString();

                //        //MessageBox.Show("chkFGColor:" + chkFGColor + "-xFGColor:" + xFGColor + ", chkFBCode:" + chkFBCode + "-xFBCode:" + xFBCode);

                //        if (lblFBStatus.Text.Replace("Status : ", "") == "New")
                //        {
                //            if (xFGColor == chkFGColor && xFBCode == chkFBCode)
                //            {
                //                if (FUNCT.msgQuiz("Found duplicate fabric code. Do you want to update ?\nพบโค้ดผ้ามีอยู่แล้วในตาราง ต้องการจะแทนที่ข้อมูลเดิมหรือไม่") == true)
                //                {
                //                    lblRowID.Text = runLoop.ToString();
                //                    lblFBStatus.Text = "Status : Update";
                //                    lblFBStatus.BackColor = Color.Navy;
                //                    break;
                //                }
                //                else
                //                {
                //                    return;
                //                }
                //            }
                //        }
                //        else if (lblFBStatus.Text.Replace("Status : ", "") == "Update")
                //        {
                //            if (xFGColor == chkFGColor && xFBCode == chkFBCode && runLoop.ToString() != lblRowID.Text)
                //            {
                //                FUNCT.msgError("Cannot update. Because the fabric code is duplicated with other rows ?\nไม่สามารถแทนที่ข้อมูลได้ เนื่องจากพบโค้ดผ้าซ้ำกับแถวอื่นในตาราง");
                //                return;
                //            }
                //        }

                //        runLoop++;
                //    }
                //}
                //***********************************

                if (lblMTStatus.Text.Replace("Status : ", "") == "New")
                {
                    bool chkDUP = false;
                    //copy datasource to datatable
                    DataTable dtMTListX = (DataTable)gridControl8.DataSource;
                    //MessageBox.Show("1: " + dtFabric.Rows.Count.ToString());
                    if (dtMTListX != null)
                    {
                        DataTable dtMT = (DataTable)gridControl7.DataSource;
                        if (dtMT != null)
                        {
                            foreach (DataRow rMT in dtMT.Rows)
                            {
                                string MTColor = rMT["Color"].ToString();
                                string MTSize = rMT["Size"].ToString();

                                //chk Material code duplicate
                                foreach (DataRow rDup in dtMTListX.Rows)
                                {
                                    string ColorName = rDup["ColorName"].ToString();
                                    string SizeName = rDup["MatSize"].ToString();
                                    string MatCode = rDup["NAVCode"].ToString();

                                    string chkMatCode = slMatCode_Mat.Text.Trim() == "" ? "" : slMatCode_Mat.EditValue.ToString();
                                    if (ColorName == MTColor && SizeName == MTSize && MatCode == chkMatCode)
                                    {
                                        chkDUP = true;
                                        break;
                                    }
                                }

                                if (chkDUP == true)
                                    break;
                            }
                        }
                    }
                    else
                    {
                        dtMTListX = dtMTList;
                    }

                    //MessageBox.Show(chkDUP.ToString());

                    if (chkDUP == false)
                    {
                        //MessageBox.Show("2: " + dtFabric.Rows.Count.ToString());
                        DataTable dtMS = (DataTable)gridControl7.DataSource;
                        if (dtMS != null)
                        {
                            foreach (DataRow drMS in dtMS.Rows)
                            {
                                string msColor = drMS["Color"].ToString();
                                string msSize = drMS["Size"].ToString();
                                string msConsumption = drMS["Consumption"].ToString();
                                string Price = txtPrice_Mat.Text.Trim() == "" ? "0" : txtPrice_Mat.Text.Trim();
                                //MessageBox.Show("Color:" + msColor + ", Size:" + msSize);

                                string WorkStation = glWorkStation_Mat.Text.Trim() == "" ? "" : glWorkStation_Mat.EditValue.ToString();
                                string Vendor = slVendor_Mat.Text.Trim() == "" ? "" : slVendor_Mat.EditValue.ToString();
                                string MatColor = slMatColor_Mat.Text.Trim() == "" ? "" : slMatColor_Mat.EditValue.ToString();
                                string Unit = slueConsumpUnit.Text.Trim() == "" ? "" : slueConsumpUnit.EditValue.ToString();
                                string Currency = glCurrency_Mat.Text.Trim() == "" ? "" : glCurrency_Mat.EditValue.ToString();
                                string MatCode = slMatCode_Mat.Text.Trim() == "" ? "" : slMatCode_Mat.EditValue.ToString();

                                //if (slMatCode_Mat.Text.Trim() != "" && MatCode == "")
                                //{
                                //    MatCode = DBC.DBQuery("SELECT TOP(1) OIDITEM AS ID FROM Items WHERE (Code = N'" + slMatCode_Mat.Text.Trim() + "') ").getString();
                                //}

                                dtMTListX.Rows.Add(
                                    "",
                                    txtSampleID_Mat.Text.Trim()
                                    , WorkStation
                                    , txtVendorMatCode_Mat.Text.Trim()
                                    , txtSampleLotNo_Mat.Text.Trim()
                                    , Vendor
                                    , MatColor
                                    , msColor
                                    , msSize
                                    , msConsumption
                                    , Unit
                                    , txtMatComposition_Mat.Text.Trim()
                                    , "" //detail
                                    , Price
                                    , Currency
                                    , MatCode
                                    , txeMatDescription.Text.Trim()
                                    , txtSituation_Mat.Text.Trim()
                                    , txtComment_Mat.Text.Trim()
                                    , txtRemark_Mat.Text.Trim()
                                    , txtPathFile_Mat.Text.Trim()
                                    , slMatCode_Mat.Text.Trim()
                                    , ""
                                );
                            }
                        }
                        else
                        {
                            dtMS = dtCSConsumption;
                        }


                        gridControl8.DataSource = dtMTListX;
                        gridView8.BestFitColumns();

                        //Clear Data
                        if (FUNCT.msgQuiz("Do you want to clear data in field above?\nต้องการเคลียร์ข้อมูลด้านบนหรือไม่") == true)
                        {
                            ClearMaterial();
                        } 
                    }
                    else
                    {
                        FUNCT.msgWarning("Material Code is duplicate. Please change."); slMatCode_Mat.Focus(); return;
                    }
                }
                else if (lblMTStatus.Text.Replace("Status : ", "") == "Update")
                {
                    if (FUNCT.msgQuiz("Confirm update this data ?") == true)
                    {
                        bool chkDUP = false;
                        //copy datasource to datatable
                        DataTable dtMTListX = (DataTable)gridControl8.DataSource;
                        //chk fabric code duplicate
                        int RowID = 0;
                        DataTable dtMT = (DataTable)gridControl7.DataSource;
                        if (dtMT != null)
                        {
                            foreach (DataRow rMT in dtMT.Rows)
                            {
                                string MTColor = rMT["Color"].ToString();
                                string MTSize = rMT["Size"].ToString();

                                //chk Material code duplicate
                                foreach (DataRow rDup in dtMTListX.Rows)
                                {
                                    if (RowID != Convert.ToInt32(txtMatRecordID_Mat.Text))
                                    {
                                        string ColorName = rDup["ColorName"].ToString();
                                        string SizeName = rDup["MatSize"].ToString();
                                        string MatCode = rDup["NAVCode"].ToString();

                                        string chkMatCode = slMatCode_Mat.Text.Trim() == "" ? "" : slMatCode_Mat.EditValue.ToString();

                                        //MessageBox.Show("MTColor:" + MTColor + " - ColorName:" + ColorName + "\nMTSize:" + MTSize + " - SizeName:" + SizeName + "\nchkMatCode:" + chkMatCode + " - MatCode:" + MatCode);

                                        if (ColorName == MTColor && SizeName == MTSize && MatCode == chkMatCode)
                                        {
                                            chkDUP = true;
                                            break;
                                        }
                                    }
                                    RowID++;
                                }

                                if (chkDUP == true)
                                    break;
                            }
                        }
                        else
                        {
                            dtMT = dtCSConsumption;
                        }

                        if (chkDUP == false)
                        {
                            string WorkStation = glWorkStation_Mat.Text.Trim() == "" ? "" : glWorkStation_Mat.EditValue.ToString();
                            string Vendor = slVendor_Mat.Text.Trim() == "" ? "" : slVendor_Mat.EditValue.ToString();
                            string MatColor = slMatColor_Mat.Text.Trim() == "" ? "" : slMatColor_Mat.EditValue.ToString();
                            string Unit = slueConsumpUnit.Text.Trim() == "" ? "" : slueConsumpUnit.EditValue.ToString();
                            string Currency = glCurrency_Mat.Text.Trim() == "" ? "" : glCurrency_Mat.EditValue.ToString();
                            string MatCode = slMatCode_Mat.Text.Trim() == "" ? "" : slMatCode_Mat.EditValue.ToString();
                            //if (slMatCode_Mat.Text.Trim() != "" && MatCode == "")
                            //{
                            //    MatCode = DBC.DBQuery("SELECT TOP(1) OIDITEM AS ID FROM Items WHERE (Code = N'" + slMatCode_Mat.Text.Trim() + "') ").getString();
                            //}

                            dtMTListX.Rows[Convert.ToInt32(txtMatRecordID_Mat.Text)]["WorkStation"] = WorkStation;
                            dtMTListX.Rows[Convert.ToInt32(txtMatRecordID_Mat.Text)]["VendMTCode"] = txtVendorMatCode_Mat.Text.Trim();
                            dtMTListX.Rows[Convert.ToInt32(txtMatRecordID_Mat.Text)]["SMPLotNo"] = txtSampleLotNo_Mat.Text.Trim();
                            dtMTListX.Rows[Convert.ToInt32(txtMatRecordID_Mat.Text)]["Vendor"] = Vendor;
                            dtMTListX.Rows[Convert.ToInt32(txtMatRecordID_Mat.Text)]["MatColor"] = MatColor;

                            if (gridView7.RowCount > 0)
                            {
                                dtMTListX.Rows[Convert.ToInt32(txtMatRecordID_Mat.Text)]["Consumption"] = gridView7.GetRowCellValue(0, "Consumption");
                            }

                            dtMTListX.Rows[Convert.ToInt32(txtMatRecordID_Mat.Text)]["Unit"] = Unit;
                            dtMTListX.Rows[Convert.ToInt32(txtMatRecordID_Mat.Text)]["Composition"] = txtMatComposition_Mat.Text.Trim();
                            dtMTListX.Rows[Convert.ToInt32(txtMatRecordID_Mat.Text)]["Price"] = txtPrice_Mat.Text.Trim() == "" ? "0" : txtPrice_Mat.Text.Trim();
                            dtMTListX.Rows[Convert.ToInt32(txtMatRecordID_Mat.Text)]["Currency"] = Currency;
                            dtMTListX.Rows[Convert.ToInt32(txtMatRecordID_Mat.Text)]["NAVCode"] = MatCode;
                            dtMTListX.Rows[Convert.ToInt32(txtMatRecordID_Mat.Text)]["Description"] = txeMatDescription.Text.Trim();
                            dtMTListX.Rows[Convert.ToInt32(txtMatRecordID_Mat.Text)]["Situation"] = txtSituation_Mat.Text.Trim();
                            dtMTListX.Rows[Convert.ToInt32(txtMatRecordID_Mat.Text)]["Comment"] = txtComment_Mat.Text.Trim();
                            dtMTListX.Rows[Convert.ToInt32(txtMatRecordID_Mat.Text)]["Remark"] = txtRemark_Mat.Text.Trim();
                            dtMTListX.Rows[Convert.ToInt32(txtMatRecordID_Mat.Text)]["PathFile"] = txtPathFile_Mat.Text.Trim();
                            dtMTListX.Rows[Convert.ToInt32(txtMatRecordID_Mat.Text)]["ChkMTCode"] = slMatCode_Mat.Text.Trim();

                            gridControl8.DataSource = dtMTListX;
                            gridView8.BestFitColumns();

                            //Clear Data
                            if (FUNCT.msgQuiz("Do you want to clear data in field above?\nต้องการเคลียร์ข้อมูลด้านบนหรือไม่") == true)
                            {
                                ClearMaterial();
                            }
                            //else
                            //{
                            //    lblMTStatus.Text = "Status : Update";
                            //    lblMTStatus.BackColor = Color.Navy;
                            //}
                        }
                        else
                        {
                            FUNCT.msgWarning("Material Code is duplicate. Please change."); slMatCode_Mat.Focus(); return;
                        }
                    }

                }
            }
            
        }

        private void gridControl7_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            GridControl gridControl = (GridControl)sender;
            GridView currentView = (GridView)gridControl.FocusedView;
            if (e.KeyCode == Keys.Delete) { 
                currentView.DeleteRow(currentView.FocusedRowHandle);
            }
        }

        private void gridView8_DoubleClick(object sender, EventArgs e)
        {
            DXMouseEventArgs ea = e as DXMouseEventArgs;
            GridView view = sender as GridView;
            GridHitInfo info = view.CalcHitInfo(ea.Location);
            if (info.InRow || info.InRowCell)
            {
                GridView gv = gridView8 as GridView;
                int index = gv.FocusedRowHandle;

                string MatID = gv.GetFocusedRowCellValue("MatID").ToString();
                string SampleID = gv.GetFocusedRowCellValue("SampleID").ToString();
                string WorkStation = gv.GetFocusedRowCellValue("WorkStation").ToString();
                string VendMTCode = gv.GetFocusedRowCellValue("VendMTCode").ToString();
                string SMPLotNo = gv.GetFocusedRowCellValue("SMPLotNo").ToString();
                string Vendor = gv.GetFocusedRowCellValue("Vendor").ToString();
                string MatColor = gv.GetFocusedRowCellValue("MatColor").ToString();
                string ColorName = gv.GetFocusedRowCellValue("ColorName").ToString();
                string MatSize = gv.GetFocusedRowCellValue("MatSize").ToString();
                string Consumption = gv.GetFocusedRowCellValue("Consumption").ToString();
                string MatUnit = gv.GetFocusedRowCellValue("Unit").ToString();
                string Composition = gv.GetFocusedRowCellValue("Composition").ToString();
                string Price = gv.GetFocusedRowCellValue("Price").ToString();
                string Currency = gv.GetFocusedRowCellValue("Currency").ToString();
                string NAVCode = gv.GetFocusedRowCellValue("NAVCode").ToString();
                string Situation = gv.GetFocusedRowCellValue("Situation").ToString();
                string Comment = gv.GetFocusedRowCellValue("Comment").ToString();
                string Remark = gv.GetFocusedRowCellValue("Remark").ToString();
                string PathFile = gv.GetFocusedRowCellValue("PathFile").ToString();
                string ChkMTCode = gv.GetFocusedRowCellValue("ChkMTCode").ToString();

                gridView6.ClearSelection();

                DataTable dtMaterial = (DataTable)gridControl6.DataSource;
                if (dtMaterial != null)
                {
                    int iRow = 0;
                    foreach (DataRow drMaterial in dtMaterial.Rows)
                    {
                        string Color = drMaterial["ColorName"].ToString();
                        string Size = drMaterial["SizeName"].ToString();
                        if (Color == ColorName && Size == MatSize)
                        {
                            gridView6.SelectRow(iRow);
                            break;
                        }
                        iRow++;
                    }
                }

                if (gridView7.RowCount > 0)
                {
                    gridView7.SetRowCellValue(0, "Consumption", Consumption);
                }

                txtMatRecordID_Mat.Text = index.ToString();

                lblMTStatus.Text = "Status : Update";
                lblMTStatus.BackColor = Color.Navy;
                layoutControlItem101.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;

                txtSampleID_Mat.Text = SampleID;
                glWorkStation_Mat.EditValue = WorkStation;
                slVendor_Mat.EditValue = Vendor;
                txtVendorMatCode_Mat.Text = VendMTCode;
                txtSampleLotNo_Mat.Text = SMPLotNo;
                txtMatComposition_Mat.Text = Composition;
                slMatColor_Mat.EditValue = MatColor;
                slueConsumpUnit.EditValue = MatUnit;
                slMatCode_Mat.EditValue = NAVCode;
                txtPrice_Mat.Text = Price;
                glCurrency_Mat.Text = Currency;
                txtSituation_Mat.Text = Situation;
                txtComment_Mat.Text = Comment;
                txtRemark_Mat.Text = Remark;

                if (PathFile != "")
                {
                    txtPathFile_Mat.Text = PathFile;
                    picMat.Image = null;
                    try
                    {
                        picMat.Image = Image.FromFile(PathFile);
                    }
                    catch (Exception) { }
                }
                else
                {
                    txtPathFile_Mat.Text = "";
                    picMat.Image = null;
                }

            }

            //if (gridView8.RowCount > 0)
            //{
            //    status_Mat = "update";
            //    bbiSave.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;   //hide
            //    bbiEdit.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;  //show
            //    gridControl6.Enabled        = false;
            //    btnGettoLlist_Mat.Enabled   = false;
            //    slMatColor_Mat.Enabled      = false;

            //    // Set Update FormDetails
            //    var s = sender;
            //    string SampleID = ct.getCellVal(s, "SampleID");
            //    string MatID    = ct.getCellVal(s, "MatID");
            //    string sql = "Select ROW_NUMBER() Over(Order by OIDSMPLMT) as No, c.OIDCOLOR as Color,s.OIDSIZE as Size, Consumption as Consumption, u.OIDUNIT as Unit , q.OIDSMPL as OIDSMPL,";
            //    sql += " m.OIDSMPLMT as MatID,q.OIDSMPL as SampleID,d.Name as WorkStation,VendMTCode,SMPLotNo,v.Name as Vendor,c.ColorName as MatColor,s.SizeName as MatSize,m.Composition,Details,Price,cr.Currency as Currency,i.Code as NAVCode,m.Situation,Comment,Remark,m.PathFile,Consumption/*,m.OIDUNIT *//*Special OIDValue*/,d.OIDDEPT,v.OIDVEND,c.OIDCOLOR,i.OIDITEM,cr.OIDCURR From SMPLRequestMaterial m inner join SMPLQuantityRequired q on q.OIDSMPLDT = m.OIDSMPLDT inner join Departments d on d.OIDDEPT = m.OIDDEPT inner join Vendor v on v.OIDVEND = m.OIDVEND left join ProductColor c on c.OIDCOLOR = m.MTColor inner join ProductSize s on s.OIDSIZE = m.MTSize left join Currency cr on cr.OIDCURR = m.OIDCURR left join Items i on i.OIDITEM = m.OIDITEM inner join Unit u on u.OIDUNIT = m.OIDUNIT Where q.OIDSMPL = " + SampleID + " And m.OIDSMPLMT = "+ MatID + " ";

            //    txtSampleID_Mat.EditValue       = SampleID;
            //    txtMatRecordID_Mat.EditValue    = MatID;
            //    glWorkStation_Mat.EditValue     = db.get_oneParameter(sql,mainConn, "OIDDEPT");
            //    slVendor_Mat.EditValue          = db.get_oneParameter(sql, mainConn, "OIDVEND");
            //    slMatColor_Mat.EditValue        = db.get_oneParameter(sql,mainConn, "OIDCOLOR");
            //    slMatCode_Mat.EditValue         = db.get_oneParameter(sql,mainConn, "OIDITEM");
            //    glCurrency_Mat.EditValue        = db.get_oneParameter(sql,mainConn, "OIDCURR");

            //    txtVendorMatCode_Mat.EditValue  = ct.getCellVal(s, "VendMTCode");
            //    txtSampleLotNo_Mat.EditValue    = ct.getCellVal(s, "SMPLotNo");
            //    txtMatComposition_Mat.EditValue = ct.getCellVal(s, "Composition");
            //    txtPrice_Mat.EditValue          = ct.getCellVal(s, "Price");
            //    txtSituation_Mat.EditValue      = ct.getCellVal(s, "Situation");
            //    txtComment_Mat.EditValue        = ct.getCellVal(s, "Comment");
            //    txtRemark_Mat.EditValue         = ct.getCellVal(s, "Remark");
            //    txtPathFile_Mat.EditValue       = (ct.getCellVal(s, "PathFile") == "") ? "" : picPath+ct.getCellVal(s, "PathFile");
            //    //picMat.Image                    = (ct.getCellVal(s, "PathFile") == "") ? null : Image.FromFile(picPath + ct.getCellVal(s, "PathFile"));

            //    if (ct.getCellVal(s, "PathFile") != "")
            //    {
            //        try
            //        {
            //            picMat.Image = Image.FromFile(picPath + ct.getCellVal(s, "PathFile"));
            //        }
            //        catch { ct.showInfoMessage("Can't Find Image File Destination!"); }
            //    }
            //    else
            //    {
            //        picMat.Image = null;
            //    }

            //    // Set Grid
            //    db.getDgv(sql,gridControl7,mainConn);
            //}
        }

        private void gridView4_DoubleClick(object sender, EventArgs e)
        {
            DXMouseEventArgs ea = e as DXMouseEventArgs;
            GridView view = sender as GridView;
            GridHitInfo info = view.CalcHitInfo(ea.Location);
            if (info.InRow || info.InRowCell)
            {
                GridView gv = gridView4 as GridView;
                int index = gv.FocusedRowHandle;
                lblRowID.Text = index.ToString();

                lblFBStatus.Text = "Status : Update";
                lblFBStatus.BackColor = Color.Navy;

                string ColorName = gv.GetFocusedRowCellValue("ColorName").ToString();
                string VendorFBCode = gv.GetFocusedRowCellValue("VendorFBCode").ToString();
                string SMPLotNo = gv.GetFocusedRowCellValue("SMPLotNo").ToString();
                string Supplier = gv.GetFocusedRowCellValue("Supplier").ToString();
                string FabricColor = gv.GetFocusedRowCellValue("FabricColor").ToString();
                string FabricCode = gv.GetFocusedRowCellValue("FabricCode").ToString();
                string Description = gv.GetFocusedRowCellValue("Description").ToString();
                string Composition = gv.GetFocusedRowCellValue("Composition").ToString();
                string FBWeight = gv.GetFocusedRowCellValue("FBWeight").ToString();
                string WidthCut = gv.GetFocusedRowCellValue("WidthCut").ToString();
                string Price = gv.GetFocusedRowCellValue("Price").ToString();
                string Currency = gv.GetFocusedRowCellValue("Currency").ToString();
                string TTWidth = gv.GetFocusedRowCellValue("TTWidth").ToString();
                string UsableWidth = gv.GetFocusedRowCellValue("UsableWidth").ToString();
                string PicFile = gv.GetFocusedRowCellValue("PicFile").ToString();
                string FBPartsID = gv.GetFocusedRowCellValue("FBPartsID").ToString();
                string FBPartsName = gv.GetFocusedRowCellValue("FBPartsName").ToString();
                string Remark = gv.GetFocusedRowCellValue("Remark").ToString();

                slFGColor_FB.EditValue = ColorName;
                txtVendorFBCode_FB.Text = VendorFBCode;
                txtSampleLotNo_FB.Text = SMPLotNo;
                slVendor_FB.EditValue = Supplier;
                slFBColor_FB.EditValue = FabricColor;
                slFBCode_FB.EditValue = FabricCode;
                lblDescription.Text = Description;
                txtComposition_FB.Text = Composition;
                txtWeightFB_FB.Text = FBWeight;
                txtWidthCuttable_FB.Text = WidthCut;
                txtPrice_FB.Text = Price;
                glCurrency_FB.EditValue = Currency;
                txtTotalWidth_FB.Text = TTWidth;
                txtUsableWidth_FB.Text = UsableWidth;
                txeRemark_FB.Text = Remark;

                if (PicFile != "")
                {
                    txtImgUpload_FB.Text = PicFile;
                    picUpload_FB.Image = null;
                    try
                    {
                        picUpload_FB.Image = Image.FromFile(PicFile);
                    }
                    catch (Exception) { }
                }
                else
                {
                    txtImgUpload_FB.Text = "";
                    picUpload_FB.Image = null;
                }


                gridView11.ClearSelection();
                FBPartsID = FBPartsID.Trim().Replace(" ", "");
                if (FBPartsID != "")
                {
                    DataTable dtGPart = (DataTable)gcPart_Fabric.DataSource;
                    if (FBPartsID.IndexOf(',') != -1)
                    {
                        string[] ID = FBPartsID.Split(',');
                        if (ID.Length > 0)
                        {
                            foreach (string idPart in ID)
                            {
                                int iRow = 0;
                                foreach (DataRow drPart in dtGPart.Rows)
                                {
                                    string Part = drPart["ID"].ToString();
                                    if (idPart == Part)
                                    {
                                        gridView11.SelectRow(iRow);
                                        break;
                                    }
                                    iRow++;
                                }
                            }
                        }

                    }
                    else
                    {
                        int iRow = 0;
                        foreach (DataRow drPart in dtGPart.Rows)
                        {
                            string Part = drPart["ID"].ToString();
                            if (FBPartsID == Part)
                            {
                                gridView11.SelectRow(iRow);
                                break;
                            }
                            iRow++;
                        }
                    }
                    
                }


                DataTable dtFB = new DataTable();
                dtFB = (DataTable)gridControl3.DataSource;
                if (dtFB.Rows.Count > 0)
                {
                    string ColorNo = ColorName;
                    int iFB = 0;
                    foreach (DataRow rFB in dtFB.Rows)
                    {
                        string Color = rFB["ColorName"].ToString().Trim();
                        if (ColorNo == Color)
                        {
                            gridView3.FocusedRowHandle = iFB;
                            break;
                        }
                        iFB++;
                    }
                }

                //Load Size
                gcSize_Fabric.DataSource = null;
                DataTable dtCS = new DataTable();
                dtCS = (DataTable)gcQtyRequired.DataSource;
                if (dtCS.Rows.Count > 0)
                {
                    string ColorNo = ColorName;

                    DataTable dtFBSize = new DataTable();
                    dtFBSize.Columns.Add("SizeName", typeof(String));
                    dtFBSize.Columns.Add("Quantity", typeof(String));
                    dtFBSize.Columns.Add("OIDSMPLDT", typeof(String));

                    foreach (DataRow rCS in dtCS.Rows)
                    {
                        string Color = rCS["Color"].ToString().Trim();
                        if (Color != "")
                        {
                            string Size = rCS["Size"].ToString().Trim();
                            string Quantity = rCS["Quantity"].ToString().Trim();
                            string DT_ID = rCS["ID"].ToString().Trim();

                            if (ColorNo == Color)
                                dtFBSize.Rows.Add(Size, Quantity, DT_ID);
                        }
                    }

                    gcSize_Fabric.DataSource = dtFBSize;
                    gridView12.SelectAll();
                    gridView12.BestFitColumns();

                }
            }

            
        }

        //public void refreshFabric()
        //{
        //    bbiSave.Enabled         = false;
        //    bbiEdit.Enabled         = false;
        //    gridControl3.Enabled    = false;
        //    btngetListFB_FB.Enabled = false;
        //    //bbiSave.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
        //    //bbiEdit.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;

        //    db.getGrid_FBListSample(gridControl3, " And smplQR.OIDSMPL = " + dosetOIDSMPL + " ");
        //    //db.getDgv("Select OIDGParts,GarmentParts From GarmentParts", gcPart_Fabric, mainConn);
        //    txtSampleID_FB.Text = dosetOIDSMPL;
        //    txtFabricRacordID_FB.Text = db.get_oneParameter("Select case When ISNULL( MAX(OIDSMPLFB),'') = '' Then 1 Else MAX(OIDSMPLFB) End as maxFB From SMPLRequestFabric", mainConn, "maxFB");

        //    //get List of Fabric
        //    db.getListofFabric(gcList_Fabric, dosetOIDSMPL);

        //    //Set New OIDFB
        //    txtFabricRacordID_FB.EditValue = null;//db.get_newOIDFB();

        //    //Clear Data
        //    ClearFabric();
        //}

        private void bbiRefresh_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            //MessageBox.Show(this.ConnectionString);
            if (rgDocActive.EditValue == null)
                rgDocActive.EditValue = 1;
            if (rgDocUser.EditValue == null)
                rgDocUser.EditValue = 0;

            tabbedControlGroup1.SelectedTabPage = layoutControlGroup1;
            getGrid_SMPL(gridControl1, gridView1, UserLogin.OIDUser, Convert.ToInt32(rgDocActive.EditValue.ToString()), Convert.ToInt32(rgDocUser.EditValue.ToString()));
            HideSelectDoc();


            //string Season = glSeason_Main.Text.ToString().Trim().Replace("'", "''");
            //string SaleSection = glSaleSection_Main.EditValue.ToString();
            //string DEPCode = DBC.DBQuery("SELECT Code FROM Departments WHERE (OIDDEPT = '" + SaleSection + "') ").getString();

            //StringBuilder sbGEN = new StringBuilder();
            //sbGEN.Append("SELECT TOP (1) FORMAT(CAST(SUBSTRING(SMPLNo, CASE WHEN CHARINDEX('-', SMPLNo) > 0 THEN CHARINDEX('-', SMPLNo) - 4 ELSE LEN(SMPLNo) - 3 END, 4) + 1 AS Int), '0000') AS genD4 ");
            //sbGEN.Append("FROM SMPLRequest ");
            //sbGEN.Append("WHERE (SMPLNo LIKE N'" + Season + DEPCode + "%') AND (LEN(SMPLNo) > 9) ");
            //sbGEN.Append("ORDER BY genD4 DESC ");
            //string strRUN = DBC.DBQuery(sbGEN).getString();
            //if (strRUN == "")
            //    strRUN = "0001";
            //string SMPLNo = Season + DEPCode + strRUN;
            //MessageBox.Show(SMPLNo);

        }

        private void gridView1_PrintInitialize(object sender, DevExpress.XtraGrid.Views.Base.PrintInitializeEventArgs e)
        {
            PrintingSystemBase pb = e.PrintingSystem as PrintingSystemBase;
            pb.PageSettings.Landscape = true;
        }

        private void gridView8_PrintInitialize(object sender, DevExpress.XtraGrid.Views.Base.PrintInitializeEventArgs e)
        {
            PrintingSystemBase pb = e.PrintingSystem as PrintingSystemBase;
            pb.PageSettings.Landscape = true;
        }

        //private void gvQtyRequired_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        //{
        //    if (e.Info.IsRowIndicator) e.Info.DisplayText = (e.RowHandle + 1).ToString();
        //}

        private void gridView7_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator) e.Info.DisplayText = (e.RowHandle + 1).ToString();
        }

        private void gridView4_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator) e.Info.DisplayText = (e.RowHandle + 1).ToString();
        }

        private void gridView5_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator) e.Info.DisplayText = (e.RowHandle + 1).ToString();
        }

        private void glCategoryDivision_Main_EditValueChanged(object sender, EventArgs e)
        {
            if (glCategoryDivision_Main.Text.Trim() != "")
            {
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT StyleName, OIDSTYLE AS ID From ProductStyle Where (OIDGCATEGORY = '" + glCategoryDivision_Main.EditValue.ToString() + "') ORDER BY StyleName ");
                new ObjDE.setSearchLookUpEdit(slStyleName_Main, sbSQL, "StyleName", "ID").getData();
                slStyleName_Main.Properties.View.PopulateColumns(slStyleName_Main.Properties.DataSource);
                slStyleName_Main.Properties.View.Columns["ID"].Visible = false;
            }
            slStyleName_Main.Focus();
        }

        private void rep_slUnit_EditValueChanged(object sender, EventArgs e)
        {
           
        }

        //private void gvQtyRequired_ValidateRow(object sender, DevExpress.XtraGrid.Views.Base.ValidateRowEventArgs e)
        //{
        //    GridView view = sender as GridView;
        //    DevExpress.XtraGrid.Columns.GridColumn ColorCol = view.Columns["Color"];
        //    DevExpress.XtraGrid.Columns.GridColumn SizeCol = view.Columns["Size"];
        //    string strColor = "";
        //    try { strColor = view.GetRowCellValue(e.RowHandle, ColorCol) != null ? (String)view.GetRowCellValue(e.RowHandle, ColorCol) : ""; } catch (Exception) { }
        //    string strSize = "";
        //    try { strSize = view.GetRowCellValue(e.RowHandle, SizeCol) != null ? (String)view.GetRowCellValue(e.RowHandle, SizeCol) : ""; } catch (Exception) { }

        //    bool chkSizeColor = chkDupSizeColor(strSize, strColor, e.RowHandle);
        //    //Validity criterion
        //    if (chkSizeColor == false)
        //    {
        //        e.Valid = false;
        //        //Set errors with specific descriptions for the columns
        //        view.SetColumnError(ColorCol, "Duplicate color & size. !! Please change.");
        //        view.SetColumnError(SizeCol, "Duplicate color & size. !! Please change.");
        //    }

        //}


        private void radioGroup4_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Customer Approved
            if (radioGroup4.EditValue.ToString() == "1")
            {
                layoutControlItem23.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;

                //emptySpaceItem6.AppearanceItemCaption.ForeColor = Color.Maroon;
                //emptySpaceItem10.AppearanceItemCaption.ForeColor = Color.Maroon;
                //emptySpaceItem25.AppearanceItemCaption.ForeColor = Color.Maroon;

                layoutControlItem4.AppearanceItemCaption.ForeColor = Color.Maroon;
                layoutControlItem5.AppearanceItemCaption.ForeColor = Color.Maroon;
                layoutControlItem8.AppearanceItemCaption.ForeColor = Color.Maroon;
                layoutControlItem13.AppearanceItemCaption.ForeColor = Color.Maroon;
                //emptySpaceItem5.AppearanceItemCaption.ForeColor = Color.Maroon;
                //emptySpaceItem11.AppearanceItemCaption.ForeColor = Color.Maroon;
                //emptySpaceItem21.AppearanceItemCaption.ForeColor = Color.Maroon;
                //emptySpaceItem24.AppearanceItemCaption.ForeColor = Color.Maroon;

                layoutControlItem68.AppearanceItemCaption.ForeColor = Color.Maroon;
                layoutControlItem69.AppearanceItemCaption.ForeColor = Color.Maroon;

                //emptySpaceItem6.Text = "* " + emptySpaceItem6.Text;
                //emptySpaceItem10.Text = "* " + emptySpaceItem10.Text;
                //emptySpaceItem25.Text = "* " + emptySpaceItem25.Text;

                layoutControlItem4.Text = "* " + layoutControlItem4.Text;
                layoutControlItem5.Text = "* " + layoutControlItem5.Text;
                layoutControlItem8.Text = "* " + layoutControlItem8.Text;
                layoutControlItem13.Text = "* " + layoutControlItem13.Text;
                //emptySpaceItem5.Text = "* " + emptySpaceItem5.Text;
                //emptySpaceItem11.Text = "* " + emptySpaceItem11.Text;
                //emptySpaceItem21.Text = "* " + emptySpaceItem21.Text;
                //emptySpaceItem24.Text = "* " + emptySpaceItem24.Text;

                layoutControlItem68.Text = "* " + layoutControlItem68.Text;
                layoutControlItem69.Text = "* " + layoutControlItem69.Text;
            }
            else if (radioGroup4.EditValue.ToString() == "0")
            {
                layoutControlItem23.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;

                //emptySpaceItem6.AppearanceItemCaption.ForeColor = Color.Maroon;
                //emptySpaceItem10.AppearanceItemCaption.ForeColor = Color.Maroon;
                //emptySpaceItem25.AppearanceItemCaption.ForeColor = Color.Maroon;

                layoutControlItem4.AppearanceItemCaption.ForeColor = Color.Empty;
                layoutControlItem5.AppearanceItemCaption.ForeColor = Color.Empty;
                layoutControlItem8.AppearanceItemCaption.ForeColor = Color.Empty;
                layoutControlItem13.AppearanceItemCaption.ForeColor = Color.Empty;
                //emptySpaceItem5.AppearanceItemCaption.ForeColor = Color.Empty;
                //emptySpaceItem11.AppearanceItemCaption.ForeColor = Color.Empty;
                //emptySpaceItem21.AppearanceItemCaption.ForeColor = Color.Empty;
                //emptySpaceItem24.AppearanceItemCaption.ForeColor = Color.Empty;

                layoutControlItem68.AppearanceItemCaption.ForeColor = Color.Empty;
                layoutControlItem69.AppearanceItemCaption.ForeColor = Color.Empty;

                //emptySpaceItem6.Text = "* " + emptySpaceItem6.Text;
                //emptySpaceItem10.Text = "* " + emptySpaceItem10.Text;
                //emptySpaceItem25.Text = "* " + emptySpaceItem25.Text;

                layoutControlItem4.Text = layoutControlItem4.Text.Replace("* ", "");
                layoutControlItem5.Text = layoutControlItem5.Text.Replace("* ", "");
                layoutControlItem8.Text = layoutControlItem8.Text.Replace("* ", "");
                layoutControlItem13.Text = layoutControlItem13.Text.Replace("* ", "");
                //emptySpaceItem5.Text = emptySpaceItem5.Text.Replace("* ", "");
                //emptySpaceItem11.Text = emptySpaceItem11.Text.Replace("* ", "");
                //emptySpaceItem21.Text = emptySpaceItem21.Text.Replace("* ", "");
                //emptySpaceItem24.Text = emptySpaceItem24.Text.Replace("* ", "");

                layoutControlItem68.Text = layoutControlItem68.Text.Replace("* ", "");
                layoutControlItem69.Text = layoutControlItem69.Text.Replace("* ", "");
            }
        }

        private void radioGroup5_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (radioGroup5.EditValue.ToString() == "0")
                layoutControlItem25.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
            else if (radioGroup5.EditValue.ToString() == "1")
                layoutControlItem25.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
        }

        private void radioGroup6_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (radioGroup6.EditValue.ToString() == "0")
                layoutControlItem27.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
            else if (radioGroup6.EditValue.ToString() == "1")
                layoutControlItem27.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
        }

        private void gvQtyRequired_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator) e.Info.DisplayText = (e.RowHandle + 1).ToString();
        }

        private void gvQtyRequired_ValidateRow(object sender, DevExpress.XtraGrid.Views.Base.ValidateRowEventArgs e)
        {

            GridView view = sender as GridView;
            DevExpress.XtraGrid.Columns.GridColumn ColorCol = view.Columns["Color"];
            DevExpress.XtraGrid.Columns.GridColumn SizeCol = view.Columns["Size"];
            string strColor = "";
            try { strColor = view.GetRowCellValue(e.RowHandle, ColorCol) != null ? (String)view.GetRowCellValue(e.RowHandle, ColorCol) : ""; } catch (Exception) { }
            string strSize = "";
            try { strSize = view.GetRowCellValue(e.RowHandle, SizeCol) != null ? (String)view.GetRowCellValue(e.RowHandle, SizeCol) : ""; } catch (Exception) { }

            bool chkSizeColor = chkDupSizeColor(strSize, strColor, e.RowHandle);
            //Validity criterion
            if (chkSizeColor == false)
            {
                e.Valid = false;
                //Set errors with specific descriptions for the columns
                view.SetColumnError(ColorCol, "Duplicate color & size. !! Please change.");
                view.SetColumnError(SizeCol, "Duplicate color & size. !! Please change.");
            }

        }

        private bool chkDupSizeColor(string strSize, string strColor, int rowIndex)
        {
            gvQtyRequired.CloseEditor();
            gvQtyRequired.UpdateCurrentRow();

            strSize = strSize.Trim();
            strColor = strColor.Trim();
            //MessageBox.Show("Size:" + strSize + ", Color:" + strColor);
            bool chkDup = true;
            //chkDup = false;

            if (strSize != "" && strColor != "")
            {
                int countCol = 0;
                DataTable dtFind = (DataTable)gcQtyRequired.DataSource;
                if (dtFind != null)
                {
                    if (dtFind.Rows.Count > 0)
                    {
                        int xRow = 0;
                        foreach (DataRow row in dtFind.Rows)
                        {
                            string chkColor = "";
                            try { chkColor = row["Color"] != null ? row["Color"].ToString().Trim() : ""; } catch (Exception) { }
                            string chkSize = "";
                            try { chkSize = row["Size"] != null ? row["Size"].ToString().Trim() : ""; } catch (Exception) { }
                            //MessageBox.Show("CSize:" + chkSize + "|Size:" + strSize + ", CColor:" + chkColor + "|Color:" + strColor);
                            if (chkColor == strColor && chkSize == strSize && xRow != rowIndex)
                                countCol++;
                            xRow++;
                        }
                        //MessageBox.Show(countCol.ToString());
                        if (countCol > 0)
                            chkDup = false;
                    }
                }
            }
            return chkDup;
        }

        private void ribbonControl_Click(object sender, EventArgs e)
        {

        }

        private void gvQtyRequired_InvalidRowException(object sender, DevExpress.XtraGrid.Views.Base.InvalidRowExceptionEventArgs e)
        {
            e.ExceptionMode = DevExpress.XtraEditors.Controls.ExceptionMode.NoAction;
            //MessageBox.Show(e.ErrorText);
        }

        private void gridView1_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator) e.Info.DisplayText = (e.RowHandle + 1).ToString();
        }

        private void gridView1_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            string ChkFBCode = Convert.ToString(gridView1.GetRowCellValue(e.RowHandle, "ChkFBCode"));
            string ChkMTCode = Convert.ToString(gridView1.GetRowCellValue(e.RowHandle, "ChkMTCode"));

            if (ChkFBCode.Length > 5)
            {
                if (ChkFBCode.Substring(0, 5).ToUpper().Trim() == "TMPFB")
                    e.Appearance.BackColor = Color.FromArgb(255, 255, 192);
            }
            else if (ChkMTCode.Length > 5)
            {
                if (ChkMTCode.Substring(0, 5).ToUpper().Trim() == "TMPMT")
                    e.Appearance.BackColor = Color.FromArgb(255, 255, 192);
            }

            GridView currentView = sender as GridView;
            if (e.Column.FieldName == "Status Name")
            {
                int value = Convert.ToInt32(currentView.GetRowCellValue(e.RowHandle, "Status"));
                if (value == 0)
                    e.Appearance.BackColor = Color.FromArgb(255, 200, 200);
                else if (value == 1)
                    e.Appearance.BackColor = Color.FromArgb(255, 224, 192);
                else if (value == 2)
                    e.Appearance.BackColor = Color.FromArgb(200, 255, 200);
            }

            
        }

        private void lblStatus_TextChanged(object sender, EventArgs e)
        {
            lblStatus.ForeColor = Color.White;
            bbiSave.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
            if (lblStatus.Text.Trim() == "New SMPL")
            {
                lblStatus.BackColor = Color.Green;
            }
            else if (lblStatus.Text.Trim() == "Update SMPL")
            {
                lblStatus.BackColor = Color.Navy;
            }
            else if (lblStatus.Text.Trim() == "Revise SMPL")
            {
                lblStatus.BackColor = Color.Purple;
            }
            else if (lblStatus.Text.Trim() == "Clone SMPL")
            {
                lblStatus.BackColor = Color.Teal;  
            }
            else if (lblStatus.Text.Trim() == "Read Only SMPL")
            {
                lblStatus.ForeColor = Color.Red;
                lblStatus.BackColor = Color.White;
                bbiSave.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
            }
            else
            {
                lblStatus.BackColor = Color.FromArgb(240, 240, 240);
                lblStatus.ForeColor = Color.Black;
            }

            if (tabbedControlGroup1.SelectedTabPage == layoutControlGroup1) //LIST
                if (chkReadWrite == 1)
                    rpgManage.Visible = true;
            else
            {
                if (lblStatus.Text.Trim() == "Read Only SMPL")
                        if (chkReadWrite == 1)
                            rpgManage.Visible = true;
                else
                    rpgManage.Visible = false;
            }
        }

        private void glBranch_Main_EditValueChanged(object sender, EventArgs e)
        {
            if (glBranch_Main.Text.Trim() != "")
            {
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT Name AS Department, OIDDEPT AS ID FROM Departments WHERE (OIDBRANCH = '" + glBranch_Main.EditValue.ToString() + "') AND (Status = 1) ORDER BY ID");
                new ObjDE.setSearchLookUpEdit(glSaleSection_Main, sbSQL, "Department", "ID").getData();
                glSaleSection_Main.Properties.View.PopulateColumns(glSaleSection_Main.Properties.DataSource);
                glSaleSection_Main.Properties.View.Columns["ID"].Visible = false;

                new ObjDE.setSearchLookUpEdit(glWorkStation_Mat, sbSQL, "Department", "ID").getData();
                glSaleSection_Main.Properties.View.PopulateColumns(glSaleSection_Main.Properties.DataSource);
                glSaleSection_Main.Properties.View.Columns["ID"].Visible = false;

                rep_MtrDept.DataSource = glWorkStation_Mat.Properties.DataSource;
                rep_MtrDept.DisplayMember = "Department";
                rep_MtrDept.ValueMember = "ID";
                rep_MtrDept.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
                //rep_MtrDept.View.PopulateColumns(rep_MtrDept.DataSource);
                //rep_MtrDept.View.Columns["ID"].Visible = false;

            }
            glSaleSection_Main.Focus();
        }

        private void sbColor_Click(object sender, EventArgs e)
        {
            var frm = new DEV01_M03(this.DBC, "FG", UserLogin.OIDUser);
            frm.ShowDialog();

        }

        private void sbSize_Click(object sender, EventArgs e)
        {
            var frm = new DEV01_M10(this.DBC, UserLogin.OIDUser);
            frm.ShowDialog(this);
        }

        private void gvQtyRequired_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {

            GridView view = sender as GridView;
            if (e.Column.ColumnHandle == 0 || e.Column.ColumnHandle == 1)
            {
                string Color = view.GetRowCellValue(e.RowHandle, view.Columns["Color"]) == null ? "" : view.GetRowCellValue(e.RowHandle, view.Columns["Color"]).ToString();
                string Size = view.GetRowCellValue(e.RowHandle, view.Columns["Size"]) == null ? "" : view.GetRowCellValue(e.RowHandle, view.Columns["Size"]).ToString();
                string Quantity = view.GetRowCellValue(e.RowHandle, view.Columns["Quantity"]) == null ? "0" : view.GetRowCellValue(e.RowHandle, view.Columns["Quantity"]).ToString();

                if (Color != tmpColor || Size != tmpSize || Quantity != tmpQuantity)
                {
                    tmpColor = Color;
                    tmpSize = Size;
                    tmpQuantity = Quantity;

                    string xQty = "0";
                    //MessageBox.Show(txeQtyDF.Text.Trim());
                    if (txeQtyDF.Text.Trim() != "" || txeQtyDF.Text.Trim() != "0")
                    {
                        xQty = txeQtyDF.Text.Trim();
                    }

                    if (Color != "" || Size != "")
                    {
                        if (Quantity == "" || Quantity == "0")
                            gvQtyRequired.SetRowCellValue(e.RowHandle, "Quantity", xQty);
                    }

                    //MessageBox.Show("Color:" + Color + ", Size:" + Size + ", Qty:" + Quantity);
                    if (Color != "" && Size != "" && Quantity != "")
                    {
                        bool chkDRQ = false;
                        //เช็คก่อนว่าใน gcQtyRequired ยังมีสีก่อนเปลี่ยนอยู่หรือป่าว ถ้ามี(chkRQ=true)ไม่ต้องแก้ใน gcList_Fabric แต่ถ้าไม่มีแล้ว(chkRQ=false)ให้แก้ไขใน gridControl8 ด้วย
                        DataTable dtDRQ = (DataTable)gcQtyRequired.DataSource;
                        if (dtDRQ != null)
                        {
                            int xRow = 0;
                            foreach (DataRow drDRQ in dtDRQ.Rows)
                            {
                                if (xRow != e.RowHandle)
                                {
                                    string chkColor = drDRQ["Color"].ToString();
                                    string chkSize = drDRQ["Size"].ToString();
                                    if (Color == chkColor && Size == chkSize)
                                    {
                                        chkDRQ = true;
                                        break;
                                    }
                                }
                                xRow++;
                            }
                        }

                        if (chkDRQ == false)
                        {
                            transFabricMaterial();

                            //*** EDIT SIZE&COLOR ****
                            if (this.BF_Color != Color) //Edit Fabric
                            {
                                bool chkRQ = false;
                                //เช็คก่อนว่าใน gcQtyRequired ยังมีสีก่อนเปลี่ยนอยู่หรือป่าว ถ้ามี(chkRQ=true)ไม่ต้องแก้ใน gcList_Fabric แต่ถ้าไม่มีแล้ว(chkRQ=false)ให้แก้ไขใน gcList_Fabric ด้วย
                                DataTable dtRQ = (DataTable)gcQtyRequired.DataSource;
                                if (dtRQ != null)
                                {
                                    int xRow = 0;
                                    foreach (DataRow drRQ in dtRQ.Rows)
                                    {
                                        if (xRow != e.RowHandle)
                                        {
                                            string chkColor = drRQ["Color"].ToString();
                                            if (this.BF_Color == chkColor)
                                            {
                                                chkRQ = true;
                                                break;
                                            }
                                        }
                                        xRow++;
                                    }
                                }

                                if (chkRQ == false)
                                {
                                    DataTable dtFB = (DataTable)gcList_Fabric.DataSource;
                                    if (dtFB != null)
                                    {
                                        int iRow = 0;
                                        foreach (DataRow drFB in dtFB.Rows)
                                        {
                                            string chkColor = drFB["ColorName"].ToString();
                                            if (this.BF_Color == chkColor)
                                            {
                                                gridView4.SetRowCellValue(iRow, "ColorName", Color);
                                            }
                                            iRow++;
                                        }
                                    }
                                }
                            }

                            if (this.BF_Color != Color || this.BF_Size != Size) //Edit Material
                            {
                                bool chkRQ = false;
                                //เช็คก่อนว่าใน gcQtyRequired ยังมีสีก่อนเปลี่ยนอยู่หรือป่าว ถ้ามี(chkRQ=true)ไม่ต้องแก้ใน gcList_Fabric แต่ถ้าไม่มีแล้ว(chkRQ=false)ให้แก้ไขใน gridControl8 ด้วย
                                DataTable dtRQ = (DataTable)gcQtyRequired.DataSource;
                                if (dtRQ != null)
                                {
                                    int xRow = 0;
                                    foreach (DataRow drRQ in dtRQ.Rows)
                                    {
                                        if (xRow != e.RowHandle)
                                        {
                                            string chkColor = drRQ["Color"].ToString();
                                            string chkSize = drRQ["Size"].ToString();
                                            if (this.BF_Color == chkColor && this.BF_Size == chkSize)
                                            {
                                                chkRQ = true;
                                                break;
                                            }
                                        }
                                        xRow++;
                                    }
                                }

                                if (chkRQ == false)
                                {
                                    DataTable dtMT = (DataTable)gridControl8.DataSource;
                                    if (dtMT != null)
                                    {
                                        int iRow = 0;
                                        foreach (DataRow drMT in dtMT.Rows)
                                        {
                                            string chkColor = drMT["ColorName"].ToString();
                                            string chkSize = drMT["MatSize"].ToString();
                                            if (this.BF_Color == chkColor && this.BF_Size == chkSize)
                                            {
                                                gridView8.SetRowCellValue(iRow, "ColorName", Color);
                                                gridView8.SetRowCellValue(iRow, "MatSize", Size);
                                            }
                                            iRow++;
                                        }
                                    }
                                }

                            }
                        }
                        //else
                        //{
                        //    gvQtyRequired.SetRowCellValue(e.RowHandle, "Color", this.BF_Color);
                        //    gvQtyRequired.SetRowCellValue(e.RowHandle, "Size", this.BF_Size);
                        //}
                    }
                }
            }
        }

        private void transFabricMaterial()
        {
            //gvQtyRequired.CloseEditor();
            gvQtyRequired.UpdateCurrentRow();
            if(dtFBSample.Rows.Count > 0)
                dtFBSample.Rows.Clear();
            if (dtMaterial.Rows.Count > 0)
                dtMaterial.Rows.Clear();
            DataTable dtCS = new DataTable();
            dtCS = (DataTable)gcQtyRequired.DataSource;
            if (dtCS.Rows.Count > 0)
            {
                string ID = lblID.Text;
                string PatternNo = txtSMPLPatternNo_Main.Text.Trim();
                string Unit = "";
                if (slueUnit.Text != "")
                    Unit = slueUnit.EditValue.ToString();


                DataTable dtChkDup = new DataTable();
                dtChkDup.Columns.Add("SMPLPatternNo", typeof(String));
                dtChkDup.Columns.Add("ColorName", typeof(String));

                foreach (DataRow rCS in dtCS.Rows)
                {
                    string Color = rCS["Color"].ToString().Trim();
                    string chkQty = rCS["Quantity"].ToString() == "" ? "0" : rCS["Quantity"].ToString();
                    if (Color != "" && chkQty != "0")
                    {
                        string Size = rCS["Size"].ToString().Trim();
                        string Quantity = rCS["Quantity"].ToString().Trim();
                        string DT_ID = rCS["ID"].ToString().Trim();

                        bool chkDup = false;
                        if (dtChkDup.Rows.Count > 0)
                        {
                            foreach (DataRow rDup in dtChkDup.Rows)
                            {
                                string xSMPLPatternNo = rDup["SMPLPatternNo"].ToString().Trim();
                                string xColorName = rDup["ColorName"].ToString().Trim();
                                //MessageBox.Show("xColorName:" + xColorName + ", Color:" + Color);
                                if (xColorName == Color)
                                {
                                    chkDup = true;
                                    break;
                                }
                            }
                        }
                        else
                        {
                            chkDup = false;
                        }

                        if(chkDup==false)
                        {
                            dtFBSample.Rows.Add(ID, PatternNo, Color, DT_ID);
                            dtChkDup.Rows.Add(PatternNo, Color);
                        }
                        
                        dtMaterial.Rows.Add(ID, PatternNo, Color, Size, Quantity, Unit, DT_ID);
                    }
                }

                gridControl3.DataSource = dtFBSample;
                gridView3.BestFitColumns();

                gridControl6.DataSource = dtMaterial;
                gridView6.BestFitColumns();

                gcSize_Fabric.DataSource = null;
                gridView12.BestFitColumns();
            }
        }

        private void slueUnit_EditValueChanged(object sender, EventArgs e)
        {
            transFabricMaterial();
            //slueConsumpUnit.ReadOnly = false;
            slueConsumpUnit.EditValue = slueUnit.EditValue;
            //slueConsumpUnit.ReadOnly = true;
        }

        private void txtSMPLPatternNo_Main_Leave(object sender, EventArgs e)
        {
            transFabricMaterial();
        }

        private void slFBCode_FB_EditValueChanged(object sender, EventArgs e)
        {
            lblDescription.Text = "-";
            lblDescription.AppearanceItemCaption.BackColor = Color.Empty;
            if (slFBCode_FB.Text.Trim() != "")
            {
                string[] arrFB = DBC.DBQuery("SELECT Description, Code FROM Items WHERE (MaterialType IN ('" + TYPE_FABRIC + "', '" + TYPE_TEMPORARY + "')) AND (OIDITEM = '" + slFBCode_FB.EditValue.ToString() + "') ").getMultipleValue();
                if (arrFB.Length > 0)
                {
                    lblDescription.Text = arrFB[0];
                    if (arrFB[1].Trim().Length > 5)
                    {
                        if (arrFB[1].Trim().Substring(0, 5).ToUpper() == "TMPFB")
                            lblDescription.AppearanceItemCaption.BackColor = Color.FromArgb(255, 255, 192);
                        else
                            lblDescription.AppearanceItemCaption.BackColor = Color.Empty;
                    }
                    else
                        lblDescription.AppearanceItemCaption.BackColor = Color.Empty;
                }

            }
            txtComposition_FB.Focus();
            //MessageBox.Show(slFBCode_FB.Text);
        }

        private void sbTempCode_Click(object sender, EventArgs e)
        {
            var frm = new DEV01_M07(this.DBC, "FB", UserLogin.OIDUser);
            frm.ShowDialog();
        }

        private void lblFBStatus_TextChanged(object sender, EventArgs e)
        {
            if (lblFBStatus.Text.Replace("Status : ", "") == "New")
                layoutControlItem97.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            else if (lblFBStatus.Text.Replace("Status : ", "") == "Update")
                layoutControlItem97.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
            else
                layoutControlItem97.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
        }

        private void sbDeleteRow_Click(object sender, EventArgs e)
        {
            if (lblRowID.Text != "")
            {
                if (FUNCT.msgQuiz("Confirm delete this data ?") == true)
                {
                    DataTable dtFB = (DataTable)gcList_Fabric.DataSource;
                    if (dtFB != null)
                    {
                        int runRow = 0;
                        foreach (DataRow drFB in dtFB.Rows)
                        {
                            if (runRow.ToString() == lblRowID.Text)
                            {
                                dtFB.Rows.Remove(drFB);
                                break;
                            }
                            runRow++;
                        }

                        dtFB.AcceptChanges();

                        gcList_Fabric.DataSource = dtFB;
                        gcList_Fabric.Update();
                        gcList_Fabric.Refresh();
                    }

                    //GridView gv = gridView4 as GridView;
                    //gv.DeleteRow(Convert.ToInt32(lblRowID.Text));
                    //gcList_Fabric.Update();
                    ClearFabric();
                }
            }
        }


        private void ClearFabric()
        {
            //Clear Data
            lblFBStatus.Text = "Status : New";
            lblFBStatus.BackColor = Color.Green;

            txtVendorFBCode_FB.Text = "";
            txtSampleLotNo_FB.Text = "";
            slVendor_FB.EditValue = "";
            slFBColor_FB.EditValue = "";
            slFBCode_FB.EditValue = "";
            lblDescription.Text = "-";
            txtComposition_FB.Text = "";
            txtWeightFB_FB.Text = "";
            txtWidthCuttable_FB.Text = "";
            txtPrice_FB.Text = "";
            txeRemark_FB.Text = "";
            //glCurrency_FB.EditValue = "";
            txtTotalWidth_FB.Text = "";
            txtUsableWidth_FB.Text = "";
            lblRowID.Text = "";

            txtImgUpload_FB.Text = "";
            picUpload_FB.Image = null;

            gridView11.ClearSelection();

            txtVendorFBCode_FB.Focus();
        }

        private void HideSelectDoc()
        {
            bbiPrint.Enabled = false;
            bbiPrintPDF.Enabled = false;
            bbiUPDATE.Enabled = false;
            bbiREVISE.Enabled = false;
            bbiCLONE.Enabled = false;
            bbiDELBILL.Enabled = false;
        }

        private void ClearMaterial()
        {
            //Clear Data
            lblMTStatus.Text = "Status : New";
            lblMTStatus.BackColor = Color.Green;

            txtMatRecordID_Mat.Text = "";
            glWorkStation_Mat.EditValue = "";
            slVendor_Mat.EditValue = "";
            txtVendName_Mat.Text = "";
            txtVendorMatCode_Mat.Text = "";
            txtSampleLotNo_Mat.Text = "";
            txtMatComposition_Mat.Text = "";
            slMatColor_Mat.EditValue = "";
            slueConsumpUnit.EditValue = "";
            slMatCode_Mat.EditValue = "";
            txeMatDescription.Text = "";
            txtPrice_Mat.Text = "";
            glCurrency_Mat.EditValue = "";
            txtSituation_Mat.Text = "";
            txtComment_Mat.Text = "";
            txtRemark_Mat.Text = "";
            txtPathFile_Mat.Text = "";
            picMat.Image = null;

            layoutControlItem101.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;

            gridView6.ClearSelection();

            glWorkStation_Mat.Focus();
        }

        private void sbClear_Click(object sender, EventArgs e)
        {
            if (FUNCT.msgQuiz("Confirm clear data in field ?") == true)
            {
                ClearFabric();
            }
        }

        private void gridView3_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator) e.Info.DisplayText = (e.RowHandle + 1).ToString();
        }

        private void DEV01_Shown(object sender, EventArgs e)
        {
            if (this.DBC.chkCONNECTION_STING() == false)
            {
                FUNCT.msgError("Connection string is null.");
                return;
            }

            rgDocActive.EditValue = 1;
            rgDocUser.EditValue = 0;


            //if (this.cloneSMPLNo != "")
            //{
            //    btnGenSMPLNo.Enabled = false;
            //    txtSMPLNo.Text = "";
            //    lblID.Text = "";
            //    lblStatus.Text = "New";
            //    tabbedControlGroup1.SelectedTabPage = layoutControlGroup2;
            //}
            //else
            //{
            tabbedControlGroup1.SelectedTabPage = layoutControlGroup1;
            // }


        }

        private void rgDocActive_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (rgDocActive.EditValue == null)
                rgDocActive.EditValue = 1;
            if (rgDocUser.EditValue == null)
                rgDocUser.EditValue = 0;

            getGrid_SMPL(gridControl1, gridView1, UserLogin.OIDUser, Convert.ToInt32(rgDocActive.EditValue.ToString()), Convert.ToInt32(rgDocUser.EditValue.ToString()));
            HideSelectDoc();
        }

        private void glSaleSection_Main_EditValueChanged(object sender, EventArgs e)
        {
            txtReferenceNo_Main.Focus();
        }

        private void txtReferenceNo_Main_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                dtRequestDate_Main.Focus();
        }

        private void slCustomer_Main_EditValueChanged(object sender, EventArgs e)
        {
            txtContactName_Main.Focus();
        }

        private void txtContactName_Main_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                dtDeliveryRequest_Main.Focus();
        }

        private void txtSMPLItemNo_Main_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txtModelName_Main.Focus();
        }

        private void txtModelName_Main_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                glCategoryDivision_Main.Focus();
        }

        private void slStyleName_Main_EditValueChanged(object sender, EventArgs e)
        {
            txtSMPLPatternNo_Main.Focus();
        }

        private void txtSMPLPatternNo_Main_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                radioGroup3.Focus();
        }

        private void txtSituation_Main_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txtStateArrangments_Main.Focus();
        }

        private void txtStateArrangments_Main_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                radioGroup5.Focus();
        }

        private void dtRequestDate_Main_EditValueChanged(object sender, EventArgs e)
        {

        }

        private void txtVendorFBCode_FB_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txtSampleLotNo_FB.Focus();
        }

        private void txtSampleLotNo_FB_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                slVendor_FB.Focus();
        }

        private void slVendor_FB_EditValueChanged(object sender, EventArgs e)
        {
            slFBColor_FB.Focus();
        }

        private void slFBColor_FB_EditValueChanged(object sender, EventArgs e)
        {
            slFBCode_FB.Focus();
        }

        private void txtComposition_FB_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txtWeightFB_FB.Focus();
        }

        private void txtWeightFB_FB_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txtWidthCuttable_FB.Focus();
        }

        private void txtWidthCuttable_FB_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txtPrice_FB.Focus();
        }

        private void txtPrice_FB_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                glCurrency_FB.Focus();
        }

        private void glCurrency_FB_EditValueChanged(object sender, EventArgs e)
        {
            txtTotalWidth_FB.Focus();
        }

        private void txtTotalWidth_FB_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txtUsableWidth_FB.Focus();
        }

        private void txtUsableWidth_FB_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                simpleButton5.Focus();
        }

        private void glWorkStation_Mat_EditValueChanged(object sender, EventArgs e)
        {
            slVendor_Mat.Focus();
        }

        private void txtVendorMatCode_Mat_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txtSampleLotNo_Mat.Focus();
        }

        private void txtSampleLotNo_Mat_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txtMatComposition_Mat.Focus();
        }

        private void txtMatComposition_Mat_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                slMatColor_Mat.Focus();
        }

        private void slMatColor_Mat_EditValueChanged(object sender, EventArgs e)
        {
            //slMatPart_Mat.Focus();
            slMatCode_Mat.Focus();
        }

        private void slMatPart_Mat_EditValueChanged(object sender, EventArgs e)
        {
            //slMatCode_Mat.Focus();
        }

        private void slMatCode_Mat_EditValueChanged(object sender, EventArgs e)
        {
            txeMatDescription.Text = "";
            txeMatDescription.BackColor = Color.Empty;

            //string MatCode = slMatCode_Mat.Text.Trim() == "" ? "" : slMatCode_Mat.EditValue.ToString();
            //MessageBox.Show(MatCode);

            if (slMatCode_Mat.Text.Trim() != "")
            {
                string[] arrFB = DBC.DBQuery("SELECT Description, Code FROM Items WHERE (MaterialType IN ('" + TYPE_ACCESSORY + "', '" + TYPE_PACKAGING + "', '" + TYPE_TEMPORARY + "')) AND (OIDITEM = '" + slMatCode_Mat.EditValue.ToString() + "') ").getMultipleValue();
                if (arrFB.Length > 0)
                {
                    txeMatDescription.Text = arrFB[0];
                    if (arrFB[1].Trim().Length > 5)
                    {
                        if (arrFB[1].Trim().Substring(0, 5).ToUpper() == "TMPMT")
                            txeMatDescription.BackColor = Color.FromArgb(255, 255, 192);
                        else
                            txeMatDescription.BackColor = Color.Empty;
                    }
                    else
                        txeMatDescription.BackColor = Color.Empty;
                }

            }
            txtPrice_Mat.Focus();
        }

        private void txtPrice_Mat_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                glCurrency_Mat.Focus();
        }

        private void glCurrency_Mat_EditValueChanged(object sender, EventArgs e)
        {
            txtSituation_Mat.Focus();
        }

        private void txtSituation_Mat_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txtComment_Mat.Focus();
        }

        private void txtComment_Mat_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txtRemark_Mat.Focus();
        }

        private void txtRemark_Mat_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                btnUploadMat.Focus();
        }

        private void gridView4_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            string ChkFBCode = Convert.ToString(gridView4.GetRowCellValue(e.RowHandle, "ChkFBCode"));

            if (ChkFBCode.Length > 5)
            {
                if (ChkFBCode.Substring(0, 5).ToUpper().Trim() == "TMPFB")
                    e.Appearance.BackColor = Color.FromArgb(255, 255, 192);
            }

        }

        private void gridView5_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            

            //if (e.Column.FieldName == "FabricCode")
            //{
            //    string xVal = currentView.GetRowCellValue(e.RowHandle, "FabricCode").ToString();
            //    if (xVal.Length > 5)
            //    { 
            //        if(xVal.Substring(0, 5).ToUpper().Trim() == "TMPFB")

            //    }

            //    int value = Convert.ToInt32(currentView.GetRowCellValue(e.RowHandle, "Status"));
            //    if (value == 0)
            //        e.Appearance.BackColor = Color.FromArgb(255, 200, 200);
            //    else if (value == 1)
            //        e.Appearance.BackColor = Color.FromArgb(255, 255, 200);
            //    else if (value == 2)
            //        e.Appearance.BackColor = Color.FromArgb(200, 255, 200);
            //}
        }

        private void sbTempCodeMat_Click(object sender, EventArgs e)
        {
            var frm = new DEV01_M07(this.DBC, "MT", UserLogin.OIDUser);
            frm.ShowDialog();
        }

        private void btnMatDelete_Click(object sender, EventArgs e)
        {
            if (txtMatRecordID_Mat.Text != "")
            {
                if (FUNCT.msgQuiz("Confirm delete this data ?") == true)
                {
                    DataTable dtMT = (DataTable)gridControl8.DataSource;
                    if (dtMT != null)
                    {
                        int runRow = 0;
                        foreach (DataRow drMT in dtMT.Rows)
                        {
                            if (runRow.ToString() == txtMatRecordID_Mat.Text)
                            {
                                dtMT.Rows.Remove(drMT);
                                break;
                            }
                            runRow++;
                        }

                        dtMT.AcceptChanges();

                        gridControl8.DataSource = dtMT;
                        gridControl8.Update();
                        gridControl8.Refresh();
                    }

                    //GridView gv = gridView8 as GridView;
                    //gv.DeleteRow(Convert.ToInt32(txtMatRecordID_Mat.Text));

                    ClearMaterial();
                }
            }
        }

        private void gridView8_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            string ChkMTCode = Convert.ToString(gridView8.GetRowCellValue(e.RowHandle, "ChkMTCode"));

            if (ChkMTCode.Length > 5)
            {
                if (ChkMTCode.Substring(0, 5).ToUpper().Trim() == "TMPMT")
                    e.Appearance.BackColor = Color.FromArgb(255, 255, 192);
            }
        }

        private void gridView8_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator) e.Info.DisplayText = (e.RowHandle + 1).ToString();
        }

        private void gvQtyRequired_CellValueChanging(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            this.BF_Color = Convert.ToString(gvQtyRequired.GetRowCellValue(e.RowHandle, "Color"));
            this.BF_Size = Convert.ToString(gvQtyRequired.GetRowCellValue(e.RowHandle, "Size"));
            this.BF_Qty = Convert.ToString(gvQtyRequired.GetRowCellValue(e.RowHandle, "Quantity"));

            this.New_Color = Convert.ToString(gvQtyRequired.GetRowCellValue(e.RowHandle, "Color"));
            this.New_Size = Convert.ToString(gvQtyRequired.GetRowCellValue(e.RowHandle, "Size"));
            this.New_Qty = Convert.ToString(gvQtyRequired.GetRowCellValue(e.RowHandle, "Quantity"));

            if (e.Column.FieldName == "Color")
                this.New_Color = Convert.ToString(e.Value);
            else if (e.Column.FieldName == "Size")
                this.New_Size = Convert.ToString(e.Value);
            else if (e.Column.FieldName == "Quantity")
                this.New_Qty = Convert.ToString(e.Value);

            if (this.New_Color != "" && this.New_Size != "")
            {
                if (gvQtyRequired.RowCount > 1)
                {
                    bool chkDRQ = false;
                    //เช็คก่อนว่าใน gcQtyRequired ยังมีสีก่อนเปลี่ยนอยู่หรือป่าว ถ้ามี(chkRQ=true)ไม่ต้องแก้ใน gcList_Fabric แต่ถ้าไม่มีแล้ว(chkRQ=false)ให้แก้ไขใน gridControl8 ด้วย
                    DataTable dtDRQ = (DataTable)gcQtyRequired.DataSource;
                    if (dtDRQ != null)
                    {
                        int xRow = 0;
                        foreach (DataRow drDRQ in dtDRQ.Rows)
                        {
                            if (xRow != e.RowHandle)
                            {
                                string chkColor = drDRQ["Color"].ToString();
                                string chkSize = drDRQ["Size"].ToString();
                                if (this.New_Color == chkColor && this.New_Size == chkSize)
                                {
                                    chkDRQ = true;
                                    break;
                                }
                            }
                            xRow++;
                        }
                    }

                    if (chkDRQ == true)
                    {
                        FUNCT.msgWarning("Cannot enter the same color and size !!");
                        if (e.Column.FieldName == "Color")
                            gvQtyRequired.SetRowCellValue(e.RowHandle, "Color", this.BF_Color);
                        if (e.Column.FieldName == "Size")
                            gvQtyRequired.SetRowCellValue(e.RowHandle, "Size", this.BF_Size);

                    }
                }
            }
        }


        private void rep_bteDel_Click(object sender, EventArgs e)
        {
            int RowHandle = gvQtyRequired.FocusedRowHandle;
            if (RowHandle > -1)
            {
                string Color = "";
                try { Color = (string)gvQtyRequired.GetFocusedRowCellValue("Color"); } catch (Exception) { FUNCT.msgWarning("Color do not be null !!\nสีห้ามเป็นค่าว่าง"); }
                string Size = "";
                try { Size = (string)gvQtyRequired.GetFocusedRowCellValue("Size"); } catch (Exception) { FUNCT.msgWarning("Size do not be null !!\nขนาดห้ามเป็นค่าว่าง"); }

                if (Color != "" && Size != "")
                {
                    if (FUNCT.msgQuiz("Confirm delete this row ?") == true)
                    {

                        //*** EDIT SIZE&COLOR ****
                        if (Color != "") //Edit Fabric
                        {
                            bool chkRQ = false;
                            //เช็คก่อนว่าใน gcQtyRequired ยังมีสีก่อนเปลี่ยนอยู่หรือป่าว ถ้ามี(chkRQ=true)ไม่ต้องลบใน gcList_Fabric แต่ถ้าไม่มีแล้ว(chkRQ=false)ให้ลบใน gcList_Fabric ด้วย
                            DataTable dtRQ = (DataTable)gcQtyRequired.DataSource;
                            if (dtRQ != null)
                            {
                                int xRow = 0;
                                foreach (DataRow drRQ in dtRQ.Rows)
                                {
                                    if (xRow != RowHandle)
                                    {
                                        string chkColor = drRQ["Color"].ToString();
                                        if (Color == chkColor)
                                        {
                                            chkRQ = true;
                                            break;
                                        }
                                    }
                                    xRow++;
                                }
                            }

                            if (chkRQ == false)
                            {
                                DataTable dtFB = (DataTable)gcList_Fabric.DataSource;
                                if (dtFB != null)
                                {
                                    int iRow = 0;
                                    foreach (DataRow drFB in dtFB.Rows)
                                    {
                                        string chkColor = drFB["ColorName"].ToString();
                                        if (Color == chkColor)
                                        {
                                            dtFB.Rows.RemoveAt(iRow);
                                            break;
                                        }
                                        iRow++;
                                    }
                                    gcList_Fabric.DataSource = dtFB;
                                }

                                DataTable dtFBM = (DataTable)gridControl3.DataSource;
                                if (dtFBM != null)
                                {
                                    int iRow = 0;
                                    foreach (DataRow drFB in dtFBM.Rows)
                                    {
                                        string chkColor = drFB["ColorName"].ToString();
                                        if (Color == chkColor)
                                        {
                                            dtFBM.Rows.RemoveAt(iRow);
                                            break;
                                        }
                                        iRow++;
                                    }
                                    gridControl3.DataSource = dtFBM;
                                }
                            }
                        }

                        if (Color != "" || Size != "") //Edit Material
                        {
                            bool chkRQ = false;
                            //เช็คก่อนว่าใน gcQtyRequired ยังมีสีก่อนเปลี่ยนอยู่หรือป่าว ถ้ามี(chkRQ=true)ไม่ต้องลบใน gcList_Fabric แต่ถ้าไม่มีแล้ว(chkRQ=false)ให้ลบใน gridControl8 ด้วย
                            DataTable dtRQ = (DataTable)gcQtyRequired.DataSource;
                            if (dtRQ != null)
                            {
                                int xRow = 0;
                                foreach (DataRow drRQ in dtRQ.Rows)
                                {
                                    if (xRow != RowHandle)
                                    {
                                        string chkColor = drRQ["Color"].ToString();
                                        string chkSize = drRQ["Size"].ToString();
                                        if (Color == chkColor && Size == chkSize)
                                        {
                                            chkRQ = true;
                                            break;
                                        }
                                    }
                                    xRow++;
                                }
                            }

                            if (chkRQ == false)
                            {
                                DataTable dtMT = (DataTable)gridControl8.DataSource;
                                if (dtMT != null)
                                {
                                    int iRow = 0;
                                    foreach (DataRow drMT in dtMT.Rows)
                                    {
                                        string chkColor = drMT["ColorName"].ToString();
                                        string chkSize = drMT["MatSize"].ToString();
                                        if (Color == chkColor && Size == chkSize)
                                        {
                                            dtMT.Rows.RemoveAt(iRow);
                                            break;
                                        }
                                        iRow++;
                                    }
                                    gridControl8.DataSource = dtMT;
                                }

                                DataTable dtMTM = (DataTable)gridControl6.DataSource;
                                if (dtMTM != null)
                                {
                                    int iRow = 0;
                                    foreach (DataRow drMT in dtMTM.Rows)
                                    {
                                        string chkColor = drMT["ColorName"].ToString();
                                        string chkSize = drMT["SizeName"].ToString();
                                        if (Color == chkColor && Size == chkSize)
                                        {
                                            dtMTM.Rows.RemoveAt(iRow);
                                            break;
                                        }
                                        iRow++;
                                    }
                                    gridControl6.DataSource = dtMTM;
                                }
                            }
                        }

                        //DELETE REQUIRED
                        DataTable dtQR = (DataTable)gcQtyRequired.DataSource;
                        if (dtQR != null)
                        {
                            int iRow = 0;
                            foreach (DataRow drQR in dtQR.Rows)
                            {
                                string chkColor = drQR["Color"].ToString();
                                string chkSize = drQR["Size"].ToString();
                                if (Color == chkColor && Size == chkSize)
                                {
                                    dtQR.Rows.RemoveAt(iRow);
                                    break;
                                }
                                iRow++;
                            }

                            gcQtyRequired.DataSource = dtQR;
                        }


                    }
                }
                
            }
        }

        private void sbFBColor_Click(object sender, EventArgs e)
        {
            var frm = new DEV01_M03(this.DBC, "FB", UserLogin.OIDUser);
            frm.ShowDialog();
        }

        private void sbMTColor_Click(object sender, EventArgs e)
        {
            var frm = new DEV01_M03(this.DBC, "MT", UserLogin.OIDUser);
            frm.ShowDialog();
        }

        private void sbDelete_S_Click(object sender, EventArgs e)
        {
            txtPictureFile_Main.Text = "";
            picMain.Image = null;
        }

        private void sbDelete_F_Click(object sender, EventArgs e)
        {
            txtImgUpload_FB.Text = "";
            picUpload_FB.Image = null;
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            txtPathFile_Mat.Text = "";
            picMat.Image = null;
        }

        private void glUseFor_EditValueChanged(object sender, EventArgs e)
        {
            txtSMPLItemNo_Main.Focus();
        }

        private void sbFBSupplier_Click(object sender, EventArgs e)
        {
            var frm = new DEV01_M12(this.DBC, "FB", UserLogin.OIDUser);
            frm.ShowDialog();
        }

        private void sbMTSupplier_Click(object sender, EventArgs e)
        {
            var frm = new DEV01_M12(this.DBC, "MT", UserLogin.OIDUser);
            frm.ShowDialog();
        }

        private void sbUseFor_Click(object sender, EventArgs e)
        {
            var frm = new DEV01_UF(this.DBC, UserLogin.OIDUser);
            frm.ShowDialog(this);
        }

        private void sbUnit_Click(object sender, EventArgs e)
        {
            var frm = new DEV01_M13(this.DBC, UserLogin.OIDUser);
            frm.ShowDialog(this);
        }

        private void sbPart_Click(object sender, EventArgs e)
        {
            var frm = new DEV01_M06(this.DBC, UserLogin.OIDUser);
            frm.ShowDialog(this);
        }

        private void slFGColor_FB_EditValueChanged(object sender, EventArgs e)
        {
            //if (slFGColor_FB.Text.Trim() == "")
            //{
            //    layoutControlGroup10.Enabled = false;
            //    layoutControlGroup11.Enabled = false;
            //}
            //else
            //{
            //    layoutControlGroup10.Enabled = true;
            //    layoutControlGroup11.Enabled = true;
            //}
        }

        private void slFGColor_FB_Validated(object sender, EventArgs e)
        {
            
        }

        private void slFGColor_FB_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void gridView1_SelectionChanged(object sender, DevExpress.Data.SelectionChangedEventArgs e)
        {
            GridView gv = gridView1;
            int RowSelect = gv.GetFocusedDataSourceRowIndex();

            for (int i = 0; i < gridView1.DataRowCount; i++)
            {
                //DataRow row = gridView1.GetDataRow(i);
                if(i != RowSelect)
                    gridView1.UnselectRow(i);
            }

            int[] selectedRowHandles = gridView1.GetSelectedRows();
            if (selectedRowHandles.Length > 0)
            {
                bbiPrint.Enabled = true;
                bbiPrintPDF.Enabled = true;
                bbiCLONE.Enabled = true;

                string OIDUSER = gv.GetFocusedRowCellValue("ByCreated").ToString();
                if (UserLogin.OIDUser.ToString() == OIDUSER)
                {
                    bbiUPDATE.Enabled = true;
                    bbiREVISE.Enabled = true;
                    //bbiDELBILL.Enabled = true;

                    string SMPLStatus = gv.GetFocusedRowCellValue("SMPLStatus").ToString();
                    if (SMPLStatus == "0")
                        bbiDELBILL.Enabled = false;
                    else
                        bbiDELBILL.Enabled = true;
                }
                else
                {
                    bbiUPDATE.Enabled = false;
                    bbiREVISE.Enabled = false;
                    bbiDELBILL.Enabled = false;
                }  
            }
            else
            {
                bbiPrint.Enabled = false;
                bbiPrintPDF.Enabled = false;
                bbiUPDATE.Enabled = false;
                bbiREVISE.Enabled = false;
                bbiCLONE.Enabled = false;
                bbiDELBILL.Enabled = false;
            }
        }

        private void bbiUPDATE_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                string SMPLNo = "";
                if (tabbedControlGroup1.SelectedTabPage == layoutControlGroup1) //LIST
                {
                    GridView gv = gridView1;
                    SMPLNo = gv.GetFocusedRowCellValue("SMPL No.").ToString();
                }
                else
                {
                    SMPLNo = txtSMPLNo.Text.Trim();
                }

                LoadSampleRequestDocument(SMPLNo, "UPDATE");
                lblStatus.Text = "Update SMPL";
                tabbedControlGroup1.SelectedTabPage = layoutControlGroup2;
                //SetWrite();
                txtSMPLNo.Focus();
            }
            catch (Exception exc)
            {
                FUNCT.msgError(exc.ToString());
            }
        }

        private void bbiCLONE_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                string SMPLNo = "";
                if (tabbedControlGroup1.SelectedTabPage == layoutControlGroup1) //LIST
                {
                    GridView gv = gridView1;
                    SMPLNo = gv.GetFocusedRowCellValue("SMPL No.").ToString();
                }
                else
                {
                    SMPLNo = txtSMPLNo.Text.Trim();
                }
                LoadSampleRequestDocument(SMPLNo, "CLONE");
                txtSMPLNo.Text = "";
                lblStatus.Text = "Clone SMPL";
                tabbedControlGroup1.SelectedTabPage = layoutControlGroup2;
                txtSMPLNo.Focus();
            }
            catch (Exception exc)
            {
                FUNCT.msgError(exc.ToString());
            }
        }

        private void bbiDELBILL_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            string SMPLID = "";
            string SMPLNo = "";
            if (tabbedControlGroup1.SelectedTabPage == layoutControlGroup1) //LIST
            {
                GridView gv = gridView1;
                SMPLID = gv.GetFocusedRowCellValue("ID").ToString();
                SMPLNo = gv.GetFocusedRowCellValue("SMPL No.").ToString();
            }
            else
            {
                SMPLID = lblID.Text.Trim();
                SMPLNo = txtSMPLNo.Text.Trim();
            }

            if (SMPLID != "")
            {
                if (FUNCT.msgQuiz("Confirm delete this sample request ?\nยืนยันลบเอกสารนี้") == true)
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT TOP(1) OIDMARK FROM Marking WHERE (OIDSMPL = '" + SMPLID + "')");
                    string chkMK = DBC.DBQuery(sbSQL.ToString()).getString();
                    if (chkMK == "")
                    {
                        sbSQL.Clear();
                        sbSQL.Append("DELETE FROM SMPLRequestMaterial WHERE (OIDSMPLDT IN (SELECT OIDSMPLDT FROM SMPLQuantityRequired WHERE (OIDSMPL = '" + SMPLID + "')))   ");
                        sbSQL.Append("DELETE FROM SMPLRequestFabricParts WHERE (OIDSMPLDT IN (SELECT OIDSMPLDT FROM SMPLQuantityRequired WHERE (OIDSMPL = '" + SMPLID + "')))   ");
                        sbSQL.Append("DELETE FROM SMPLRequestFabric WHERE (OIDSMPLDT IN (SELECT OIDSMPLDT FROM SMPLQuantityRequired WHERE (OIDSMPL = '" + SMPLID + "')))   ");
                        sbSQL.Append("DELETE FROM SMPLQuantityRequired WHERE (OIDSMPL = '" + SMPLID + "')   ");

                        sbSQL.Append("UPDATE SMPLRequest  ");
                        sbSQL.Append("SET    SMPLStatus = 1  ");
                        sbSQL.Append("WHERE  (SMPLNo = ");
                        sbSQL.Append("            (SELECT CASE WHEN SMPLRevise > 1 THEN CONCAT(SUBSTRING(SMPLNo, 1, CASE WHEN CHARINDEX('-', SMPLNo) > 0 THEN CHARINDEX('-', SMPLNo) - 1 ELSE LEN(SMPLNo) END), '-' + CONVERT(VARCHAR, SMPLRevise - 1)) ");
                        sbSQL.Append("                    ELSE SUBSTRING(SMPLNo, 1, CASE WHEN CHARINDEX('-', SMPLNo) > 0 THEN CHARINDEX('-', SMPLNo) - 1 ELSE LEN(SMPLNo) END) END AS NewSMPLNo ");
                        sbSQL.Append("             FROM   SMPLRequest AS SRQ ");
                        sbSQL.Append("             WHERE  (OIDSMPL = ");
                        sbSQL.Append("                           (SELECT MAX(OIDSMPL) AS OIDSMPL ");
                        sbSQL.Append("                            FROM   SMPLRequest AS SR ");
                        sbSQL.Append("                            WHERE  (SUBSTRING(SMPLNo, 1, CASE WHEN CHARINDEX('-', SMPLNo) > 0 THEN CHARINDEX('-', SMPLNo) ELSE LEN(SMPLNo) END) = ");
                        sbSQL.Append("                                          (SELECT SUBSTRING(SMPLNo, 1, CASE WHEN CHARINDEX('-', SMPLNo) > 0 THEN CHARINDEX('-', SMPLNo) ELSE LEN(SMPLNo) END) AS SMPL ");
                        sbSQL.Append("                                           FROM   SMPLRequest AS xSR ");
                        sbSQL.Append("                                           WHERE  (OIDSMPL = '" + SMPLID + "')))))))   ");

                        sbSQL.Append("DELETE FROM SMPLRequest WHERE (OIDSMPL = '" + SMPLID + "')   ");

                        try
                        {
                            bool chkSave = DBC.DBQuery(sbSQL.ToString()).runSQL();
                            if (chkSave == true)
                            {
                                FUNCT.msgInfo("Delete complete.");
                                //Delete Success
                                if (SMPLNo == txtSMPLNo.Text.Trim())
                                    LoadNewData();
                                else
                                {
                                    getGrid_SMPL(gridControl1, gridView1, UserLogin.OIDUser, Convert.ToInt32(rgDocActive.EditValue.ToString()), Convert.ToInt32(rgDocUser.EditValue.ToString()));
                                    HideSelectDoc();
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            FUNCT.msgError("Error : " + ex.ToString());
                        }

                    }
                    else
                    {
                        FUNCT.msgError("Can not be deleted. Because of this sample request has been used to create marking.\nไม่สามารถลบเอกสารได้ เนื่องจากถูกนำไปสร้างมาร์คกิ้งแล้ว");
                    }
                }
            }
            else
            {
                FUNCT.msgWarning("Please select sample request document.");
            }
        }

        private void bbiREVISE_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                string SMPLNo = "";
                if (tabbedControlGroup1.SelectedTabPage == layoutControlGroup1) //LIST
                {
                    GridView gv = gridView1;
                    SMPLNo = gv.GetFocusedRowCellValue("SMPL No.").ToString();
                }
                else
                {
                    SMPLNo = txtSMPLNo.Text.Trim();
                }
                LoadSampleRequestDocument(SMPLNo, "REVISE");
                lblStatus.Text = "Revise SMPL";
                tabbedControlGroup1.SelectedTabPage = layoutControlGroup2;
                txtSMPLNo.Focus();
            }
            catch (Exception exc)
            {
                FUNCT.msgError(exc.ToString());
            }
        }

        private void rgDocUser_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (rgDocActive.EditValue == null)
                rgDocActive.EditValue = 1;
            if (rgDocUser.EditValue == null)
                rgDocUser.EditValue = 0;

            getGrid_SMPL(gridControl1, gridView1, UserLogin.OIDUser, Convert.ToInt32(rgDocActive.EditValue.ToString()), Convert.ToInt32(rgDocUser.EditValue.ToString()));
            HideSelectDoc();
        }

        private void txeQtyDF_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeQtyDF.Text = txeQtyDF.Text.Trim() == "" ? "0" : txeQtyDF.Text.Trim();
                if (txeQtyDF.Text.Trim() != "0")
                {
                    DataTable dtQTY = (DataTable)gcQtyRequired.DataSource;
                    if (dtQTY != null)
                    {
                        int runLoop = 0;
                        foreach (DataRow drQTY in dtQTY.Rows)
                        {
                            string Color = drQTY["Color"].ToString().Trim();
                            string Size = drQTY["Size"].ToString().Trim();
                            string Quantity = drQTY["Quantity"].ToString().Trim();
                            Quantity = Quantity == "" ? "0" : Quantity;
                            if (Color != "" || Size != "")
                            {
                                dtQTY.Rows[runLoop].SetField("Quantity", txeQtyDF.Text);
                            }
                            runLoop++;
                        }

                        gcQtyRequired.DataSource = dtQTY;
                        gcQtyRequired.Update();
                        gcQtyRequired.Refresh();
                    }
                }
            }
        }

        private void txeQtyDF_Validated(object sender, EventArgs e)
        {
            txeQtyDF.Text = txeQtyDF.Text.Trim() == "" ? "0" : txeQtyDF.Text.Trim();
            if (txeQtyDF.Text.Trim() != "0")
            {
                DataTable dtQTY = (DataTable)gcQtyRequired.DataSource;
                if (dtQTY != null)
                {
                    int runLoop = 0;
                    foreach (DataRow drQTY in dtQTY.Rows)
                    {
                        string Color = drQTY["Color"].ToString().Trim();
                        string Size = drQTY["Size"].ToString().Trim();
                        string Quantity = drQTY["Quantity"].ToString().Trim();
                        Quantity = Quantity == "" ? "0" : Quantity;
                        if (Color != "" || Size != "")
                        {
                            if (Quantity == "0")
                            {
                                dtQTY.Rows[runLoop].SetField("Quantity", txeQtyDF.Text);
                            }
                        }
                        runLoop++;
                    }

                    gcQtyRequired.DataSource = dtQTY;
                    gcQtyRequired.Update();
                    gcQtyRequired.Refresh();
                }
            }
        }

        private void sbMatClear_Click(object sender, EventArgs e)
        {
            if (FUNCT.msgQuiz("Confirm clear data in field ?") == true)
            {
                ClearMaterial();
            }
        }

        private void picMain_Click(object sender, EventArgs e)
        {
            if (txtPictureFile_Main.Text.Trim() != "")
            {
                Form fc = Application.OpenForms["ShowImage"];
                if (fc != null)
                    fc.Close();

                ShowImage frmIMG = new ShowImage(txtPictureFile_Main.Text.Trim());
                frmIMG.Show();
            }
        }

        private void bbiPrintPDF_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            int[] selectedRowHandles = gridView1.GetSelectedRows();
            if (selectedRowHandles.Length > 0)
            {
                gridView1.FocusedRowHandle = selectedRowHandles[0];
                string SMPLNo = gridView1.GetRowCellDisplayText(selectedRowHandles[0], "SMPL No.");
                if (FUNCT.msgQuiz("Confirm print sample request (pdf file) : " + SMPLNo + " ?") == true)
                {
                    layoutControlItem120.Text = "Print pdf file processing ..";
                    layoutControlItem120.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                    pbcEXPORT.Properties.Step = 1;
                    pbcEXPORT.Properties.PercentView = true;
                    pbcEXPORT.Properties.Maximum = 12;
                    pbcEXPORT.Properties.Minimum = 0;

                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT TOP(1) SRQ.OIDSMPL, SRQ.SMPLNo, SRQ.SMPLRevise, CPN.EngName AS Company, BCH.Name AS Branch, SRQ.ContactName, SRQ.RequestDate, ");
                    sbSQL.Append("       SRQ.Season, CUS.Name AS Customer, GC.CategoryName AS Category, SRQ.ModelName, SRQ.SMPLPatternNo, SRQ.SMPLItem, SRQ.StateArrangements, ");
                    sbSQL.Append("       CASE WHEN SRQ.PatternSizeZone = 0 THEN 'Japan' ELSE CASE WHEN SRQ.PatternSizeZone = 1 THEN 'Europe' ELSE CASE WHEN SRQ.PatternSizeZone = 2 THEN 'US' ELSE '' END END END AS SizeZone, ");
                    sbSQL.Append("       SRQ.PictureFile, SRQ.DeliveryRequest, UF.UseFor, CASE WHEN SRQ.SpecificationSize = 0 THEN 'Necessary' ELSE 'Unnecessary' END AS SpecificationSize, ");
                    sbSQL.Append("       (SELECT TOP (1) UN.UnitName AS Unit FROM SMPLQuantityRequired AS SQR INNER JOIN Unit AS UN ON SQR.OIDUnit = UN.OIDUNIT WHERE (SQR.OIDSMPL = SRQ.OIDSMPL)) AS Unit, ");
                    sbSQL.Append("       (SELECT CountColor + 'color ' + CountSize + 'size' AS TTCS FROM(SELECT TOP(1)(SELECT CONVERT(varchar, COUNT(OIDCOLOR)) AS Color FROM(SELECT OIDCOLOR FROM SMPLQuantityRequired AS B WHERE(OIDSMPL = A.OIDSMPL) GROUP BY OIDCOLOR) AS CColor) AS CountColor, (SELECT CONVERT(varchar, COUNT(OIDSIZE)) AS Size FROM(SELECT OIDSIZE FROM SMPLQuantityRequired AS C WHERE(OIDSMPL = A.OIDSMPL) GROUP BY OIDSIZE) AS CSize) AS CountSize FROM SMPLQuantityRequired AS A WHERE(OIDSMPL = SRQ.OIDSMPL)) AS A1) AS TTCS,   ");
                    //sbSQL.Append("       (SELECT TOP (1) Quantity FROM SMPLQuantityRequired WHERE (OIDSMPL = SRQ.OIDSMPL)) AS Pcs, ");
                    sbSQL.Append("       ISNULL((SELECT Quantity + '-' AS 'data()' FROM(SELECT CONVERT(VARCHAR, Quantity) AS Quantity FROM SMPLQuantityRequired WHERE OIDSMPL = SRQ.OIDSMPL) AS XA FOR XML PATH('')), '')  AS Pcs, ");
                    sbSQL.Append("       (SELECT SUM(Quantity) AS TTPcs FROM SMPLQuantityRequired WHERE (OIDSMPL = SRQ.OIDSMPL)) AS TTPcs, SRQ.ReferenceNo, SRQ.Situation, U.FullName ");
                    sbSQL.Append("FROM   SMPLRequest AS SRQ INNER JOIN ");
                    sbSQL.Append("        Branchs AS BCH ON SRQ.OIDBranch = BCH.OIDBranch INNER JOIN ");
                    sbSQL.Append("        Company AS CPN ON BCH.OIDCOMPANY = CPN.OIDCOMPANY INNER JOIN ");
                    sbSQL.Append("        SMPLUseFor AS UF ON SRQ.UseFor = UF.OIDUF LEFT OUTER JOIN ");
                    sbSQL.Append("        Customer AS CUS ON SRQ.OIDCUST = CUS.OIDCUST LEFT OUTER JOIN ");
                    sbSQL.Append("        GarmentCategory AS GC ON SRQ.OIDCATEGORY = GC.OIDGCATEGORY LEFT OUTER JOIN ");
                    sbSQL.Append("        Users AS U ON SRQ.UpdatedBy = U.OIDUSER ");
                    sbSQL.Append("WHERE (SRQ.SMPLNo = N'" + SMPLNo + "') ");
                    string[] SMPL = DBC.DBQuery(sbSQL.ToString()).getMultipleValue();
                    if (SMPL.Length > 0)
                    {
                        //****** BEGIN EXPORT *******

                        String sFilePath = System.IO.Path.Combine(new ObjSet.Folder(@"C:\MDS\Export\").GetPath() + SMPLNo + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");
                        if (File.Exists(sFilePath)) { File.Delete(sFilePath); }
                        bool chkExcel = false;
                        Microsoft.Office.Interop.Excel.Application objApp = new Microsoft.Office.Interop.Excel.Application();
                        Microsoft.Office.Interop.Excel.Worksheet objSheet = new Microsoft.Office.Interop.Excel.Worksheet();
                        Microsoft.Office.Interop.Excel.Workbook objWorkBook = null;
                        //object missing = System.Reflection.Missing.Value;

                        pbcEXPORT.PerformStep();
                        pbcEXPORT.Update();

                        try
                        {
                            int blankCol = 0;
                            objWorkBook = objApp.Workbooks.Add(Type.Missing);
                            objWorkBook = objApp.Workbooks.Open(this.reportPath + "SMPL.xlsx");

                            //objSheet = (Microsoft.Office.Interop.Excel.Worksheet)objWorkBook.ActiveSheet;
                            objSheet = (Microsoft.Office.Interop.Excel.Worksheet)objWorkBook.Sheets[1];
                            objSheet.Name = SMPLNo;

                            objSheet.Cells[2, 1] = SMPL[23].ToUpper().Trim(); //SMPL[1].Replace("-" + SMPL[2], "");
                            if (SMPL[2] != "0")
                                objSheet.Cells[2, 2] = "Revised-" + SMPL[2];

                            objSheet.Cells[2, 11] = SMPL[25];
                            objSheet.Cells[4, 2] = SMPL[3] + " " + SMPL[4];
                            objSheet.Cells[5, 2] = SMPL[5];
                            objSheet.Cells[1, 11] = SMPL[6] == "" ? "" : Convert.ToDateTime(SMPL[6]).ToString("dd/MM/yyyy");
                            objSheet.Cells[7, 2] = SMPL[7] + " " + SMPL[8] + " " + SMPL[9];
                            objSheet.Cells[9, 2] = "- " + SMPL[10] + " -";
                            objSheet.Cells[10, 2] = SMPL[11];
                            objSheet.Cells[11, 2] = SMPL[12];
                            objSheet.Cells[9, 11] = SMPL[13];
                            objSheet.Cells[10, 11] = "ใช้แพทเทิร์น " + SMPL[14];
                            objSheet.Cells[5, 11] = SMPL[17] == "" ? "SMPL" : SMPL[17].ToUpper().Trim();

                            if (SMPL[24] != "")
                            {
                                objSheet.Cells[6, 5] = SMPL[24];
                                objSheet.Range[objSheet.Cells[6, 5], objSheet.Cells[6, 6]].Interior.Color = Color.Yellow;

                            }

                            pbcEXPORT.PerformStep();
                            pbcEXPORT.Update();

                            int LastCol = 10;
                            //********* Set Column ************
                            int CountCol = DBC.DBQuery("SELECT COUNT(OIDSMPLDT) AS CountCol FROM SMPLQuantityRequired WHERE (OIDSMPL = '" + SMPL[0] + "')").getInt();
                            //CountCol = 8;
                            if (CountCol > 6)
                            {
                                for (int ci = 0; ci < CountCol - 6; ci++)
                                {
                                    objSheet.Columns[7].Insert();
                                    LastCol++;
                                }
                            }

                            pbcEXPORT.PerformStep();
                            pbcEXPORT.Update();

                            //Set Column Size
                            float LWCol = 0;
                            float NWCol = 0;

                            if (CountCol < 6)
                            {
                                LWCol = (float)((double)60 / (double)CountCol);
                                NWCol = (float)((double)17.38 / (double)(6 - CountCol));
                            }
                            else
                            {
                                LWCol = (float)((double)77.38 / (double)CountCol);
                                NWCol = 0;
                            }

                            int loopCountCol = 0;
                            for (int sc = 4; sc < LastCol; sc++)
                            {
                                if (loopCountCol < CountCol)
                                {
                                    objSheet.Columns[sc].ColumnWidth = LWCol;
                                }
                                else
                                {
                                    objSheet.Columns[sc].ColumnWidth = NWCol;
                                }
                                loopCountCol++;
                            }
                            //******* End Set Column **********

                            Microsoft.Office.Interop.Excel.Range oRange;
                            float Left = 0;
                            float Top = 0;

                            if (SMPL[15] != "")
                            {
                                oRange = (Microsoft.Office.Interop.Excel.Range)objSheet.Cells[38, 3];
                                Left = (float)((double)oRange.Left) + 1;
                                Top = (float)((double)oRange.Top) + 1;
                                string PathImgFile = this.imgPath + SMPL[15];
                                Bitmap original = new Bitmap(PathImgFile);
                                float scaleHeight = 200;
                                float scaleWidth = (scaleHeight * original.Width) / original.Height;
                                objSheet.Shapes.AddPicture(System.IO.Path.Combine(PathImgFile), Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, Left, Top, (int)scaleWidth, (int)scaleHeight);
                                original.Dispose();
                            }
                            objSheet.Cells[28, LastCol + 1] = "※ TAG";
                            objSheet.Cells[30, 2] = SMPL[16] == "" ? "" : Convert.ToDateTime(SMPL[16]).ToString("dd-MMM-yy");
                            objSheet.Cells[30, LastCol + 1] = SMPL[16] == "" ? "" : "ต้องการตัวอย่าง " + Convert.ToDateTime(SMPL[16]).ToString("MMMM-dd-yyyy");
                            objSheet.Cells[30, LastCol + 1].Font.Size = 14;
                            objSheet.Cells[31, 2] = SMPL[17];
                            objSheet.Cells[34, 2] = SMPL[18];

                            objSheet.Cells[29, LastCol] = SMPL[19];
                            objSheet.Cells[28, 2] = SMPL[20];
                            string xPCS = SMPL[21].Replace(" ", "");
                            xPCS = xPCS.Length > 0 ? xPCS.Substring(0, xPCS.Length - 1) : "";
                            objSheet.Cells[28, 5] = xPCS + SMPL[19].ToLower();
                            objSheet.Cells[29, LastCol - 1] = SMPL[22];

                            pbcEXPORT.PerformStep();
                            pbcEXPORT.Update();

                            sbSQL.Clear();
                            sbSQL.Append("SELECT SQR.OIDSIZE, PS.SizeName ");
                            sbSQL.Append("FROM   SMPLQuantityRequired AS SQR INNER JOIN ");
                            sbSQL.Append("       ProductSize AS PS ON SQR.OIDSIZE = PS.OIDSIZE ");
                            sbSQL.Append("WHERE (SQR.OIDSMPL = '" + SMPL[0] + "') ");
                            sbSQL.Append("ORDER BY SQR.OIDSIZE, SQR.OIDCOLOR ");
                            DataTable dtRQ = DBC.DBQuery(sbSQL.ToString()).getDataTable();
                            if (dtRQ.Rows.Count > 0)
                            {
                                int runCell = 4;
                                string chkSizeID = "";
                                foreach (DataRow drRQ in dtRQ.Rows)
                                {
                                    string SizeID = drRQ["OIDSIZE"].ToString();
                                    string SizeName = drRQ["SizeName"].ToString();
                                    if (chkSizeID != SizeID)
                                    {
                                        objSheet.Cells[13, runCell] = SizeName;
                                        chkSizeID = SizeID;
                                    }
                                    else //Merge Cell
                                    {
                                        objSheet.Range[objSheet.Cells[13, runCell - 1], objSheet.Cells[13, runCell]].Merge();
                                        objSheet.Cells[13, runCell] = SizeName;
                                    }
                                    runCell++;
                                }

                                if (runCell <= LastCol)
                                {
                                    objSheet.Range[objSheet.Cells[13, runCell], objSheet.Cells[13, LastCol]].Merge();
                                }
                            }

                            pbcEXPORT.PerformStep();
                            pbcEXPORT.Update();

                            sbSQL.Clear();
                            sbSQL.Append("SELECT SQR.OIDCOLOR, PC.ColorName ");
                            sbSQL.Append("FROM   SMPLQuantityRequired AS SQR INNER JOIN ");
                            sbSQL.Append("       ProductColor AS PC ON SQR.OIDCOLOR = PC.OIDCOLOR ");
                            sbSQL.Append("WHERE (SQR.OIDSMPL = '" + SMPL[0] + "') ");
                            sbSQL.Append("ORDER BY SQR.OIDSIZE, SQR.OIDCOLOR ");
                            DataTable dtRQ2 = DBC.DBQuery(sbSQL.ToString()).getDataTable();
                            if (dtRQ2.Rows.Count > 0)
                            {
                                int runCell = 4;
                                string chkColorID = "";
                                foreach (DataRow drRQ in dtRQ2.Rows)
                                {
                                    string ColorID = drRQ["OIDCOLOR"].ToString();
                                    string ColorName = drRQ["ColorName"].ToString();
                                    if (chkColorID != ColorID)
                                    {
                                        objSheet.Cells[14, runCell] = ColorName;
                                        chkColorID = ColorID;
                                    }
                                    else //Merge Cell
                                    {
                                        objSheet.Range[objSheet.Cells[14, runCell - 1], objSheet.Cells[14, runCell]].Merge();
                                        objSheet.Cells[14, runCell] = ColorName;
                                    }
                                    runCell++;
                                }

                                blankCol = runCell;
                                if (runCell <= LastCol)
                                {
                                    objSheet.Range[objSheet.Cells[14, runCell], objSheet.Cells[14, LastCol]].Merge();
                                }
                            }

                            pbcEXPORT.PerformStep();
                            pbcEXPORT.Update();

                            sbSQL.Clear();
                            sbSQL.Append("SELECT Quantity ");
                            sbSQL.Append("FROM   SMPLQuantityRequired ");
                            sbSQL.Append("WHERE (OIDSMPL = '" + SMPL[0] + "') ");
                            sbSQL.Append("ORDER BY OIDSIZE, OIDCOLOR ");
                            DataTable dtRQ3 = DBC.DBQuery(sbSQL.ToString()).getDataTable();
                            if (dtRQ3.Rows.Count > 0)
                            {
                                int runCell = 4;
                                foreach (DataRow drRQ in dtRQ3.Rows)
                                {
                                    string Quantity = drRQ["Quantity"].ToString();
                                    objSheet.Cells[16, runCell] = Quantity;
                                    runCell++;
                                }

                                if (runCell <= LastCol)
                                {
                                    objSheet.Range[objSheet.Cells[16, runCell], objSheet.Cells[16, LastCol]].Merge();
                                    //Hanger
                                    objSheet.Range[objSheet.Cells[17, runCell], objSheet.Cells[17, LastCol]].Merge();
                                }
                            }

                            pbcEXPORT.PerformStep();
                            pbcEXPORT.Update();

                            int BeginFB = 14;
                            int LastRow = BeginFB;

                            int UseRow = BeginFB;

                            int BeginMT = BeginFB;
                            int BeginFBComp = BeginFB + 4;
                            int BeginMTComp = BeginFB + 7;

                            sbSQL.Clear();
                            sbSQL.Append("SELECT RQ.OIDSMPLDT, PT.OIDGParts, PT.GarmentParts, ");
                            sbSQL.Append("       ISNULL((SELECT ColorName + ',' AS 'data()' FROM(SELECT DISTINCT Z.ColorName FROM SMPLRequestFabric AS A INNER JOIN SMPLRequestFabricParts AS YA ON A.OIDSMPLFB = YA.OIDSMPLFB INNER JOIN ProductColor AS Z ON A.OIDCOLOR = Z.OIDCOLOR WHERE A.OIDSMPLDT = RQ.OIDSMPLDT  AND YA.OIDGParts = PT.OIDGParts) AS XA FOR XML PATH('')), '')  AS FBCodeColor, ");
                            sbSQL.Append("       ISNULL((SELECT VendFBCode + ',' AS 'data()' FROM(SELECT DISTINCT B.VendFBCode FROM SMPLRequestFabric AS B INNER JOIN SMPLRequestFabricParts AS YB ON B.OIDSMPLFB = YB.OIDSMPLFB WHERE B.OIDSMPLDT = RQ.OIDSMPLDT  AND YB.OIDGParts = PT.OIDGParts) AS XB FOR XML PATH('')), '')  AS VendorFBCode, ");
                            sbSQL.Append("       ISNULL((SELECT SMPLotNo + ',' AS 'data()' FROM(SELECT DISTINCT SMPLotNo FROM SMPLRequestFabric AS C INNER JOIN SMPLRequestFabricParts AS YC ON C.OIDSMPLFB = YC.OIDSMPLFB WHERE C.OIDSMPLDT = RQ.OIDSMPLDT  AND YC.OIDGParts = PT.OIDGParts) AS XC FOR XML PATH('')), '')  AS FBLotNo, ");
                            sbSQL.Append("       ISNULL((SELECT Name + ',' AS 'data()' FROM(SELECT DISTINCT REPLACE(REPLACE(VD.Name, ' CO.,LTD.', ''), ' CO.,LTD', '') AS Name FROM SMPLRequestFabric AS D INNER JOIN SMPLRequestFabricParts AS YD ON D.OIDSMPLFB = YD.OIDSMPLFB INNER JOIN Vendor AS VD ON D.OIDVEND = VD.OIDVEND WHERE D.OIDSMPLDT = RQ.OIDSMPLDT  AND YD.OIDGParts = PT.OIDGParts) AS XD FOR XML PATH('')), '')  AS Vendor, ");
                            sbSQL.Append("       ISNULL((SELECT CASE WHEN ISNULL(PathFile, '') = '' THEN '' ELSE PathFile + ',' END AS 'data()' FROM(SELECT DISTINCT E.PathFile FROM SMPLRequestFabric AS E INNER JOIN SMPLRequestFabricParts AS YE ON E.OIDSMPLFB = YE.OIDSMPLFB WHERE E.OIDSMPLDT = RQ.OIDSMPLDT  AND YE.OIDGParts = PT.OIDGParts) AS XE FOR XML PATH('')),'')  AS FBFile, ");
                            sbSQL.Append("       ISNULL((SELECT CASE WHEN ISNULL(Composition, '') = '' THEN '' ELSE Composition + ',' END AS 'data()' FROM(SELECT DISTINCT F.Composition FROM SMPLRequestFabric AS F INNER JOIN SMPLRequestFabricParts AS YF ON F.OIDSMPLFB = YF.OIDSMPLFB WHERE F.OIDSMPLDT = RQ.OIDSMPLDT  AND YF.OIDGParts = PT.OIDGParts) AS XF FOR XML PATH('')), '')  AS Composition, ");
                            sbSQL.Append("       ISNULL((SELECT Remark + ',' AS 'data()' FROM(SELECT DISTINCT Remark FROM SMPLRequestFabric AS G INNER JOIN SMPLRequestFabricParts AS YG ON G.OIDSMPLFB = YG.OIDSMPLFB WHERE G.OIDSMPLDT = RQ.OIDSMPLDT  AND YG.OIDGParts = PT.OIDGParts) AS XG FOR XML PATH('')), '')  AS FBRemark  ");
                            sbSQL.Append("FROM   SMPLQuantityRequired AS RQ CROSS JOIN ");
                            sbSQL.Append("       (SELECT DISTINCT FP.OIDGParts, GP.GarmentParts ");
                            sbSQL.Append("        FROM   SMPLRequestFabricParts AS FP INNER JOIN ");
                            sbSQL.Append("               GarmentParts AS GP ON FP.OIDGParts = GP.OIDGParts ");
                            sbSQL.Append("        WHERE  (FP.OIDSMPLDT IN ");
                            sbSQL.Append("                  (SELECT OIDSMPLDT ");
                            sbSQL.Append("                   FROM   SMPLQuantityRequired AS A ");
                            sbSQL.Append("                   WHERE  (OIDSMPL = '" + SMPL[0] + "')))) AS PT ");
                            sbSQL.Append("WHERE (RQ.OIDSMPL = '" + SMPL[0] + "') ");
                            sbSQL.Append("ORDER BY PT.OIDGParts, RQ.OIDSIZE, RQ.OIDCOLOR ");

                            DataTable dtFB = DBC.DBQuery(sbSQL.ToString()).getDataTable();
                            if (dtFB.Rows.Count > 0)
                            {
                                DataTable chkDtFB = dtFB;
                                string Supplier = "";
                                string Composition = "";

                                string ChkComposition = "";
                                string ChkSupplier = "";

                                string chkParts = "";
                                int runCell = 4;
                                int runLoop = 0;

                                DataTable dtComposition = new DataTable();
                                dtComposition.Columns.Add("Composition", typeof(String));
                                dtComposition.Columns.Add("Supplier", typeof(String));

                                bool chkFBCode = false;
                                bool chkFBLot = false;
                                bool chkFBRemark = false;
                                foreach (DataRow drFB in dtFB.Rows)
                                {
                                    string OIDGParts = drFB["OIDGParts"].ToString();
                                    string GarmentParts = drFB["GarmentParts"].ToString();

                                    if (chkParts != OIDGParts)
                                    {
                                        UseRow++;
                                        if (chkFBCode == true)
                                            UseRow++;
                                        if (chkFBLot == true)
                                            UseRow++;
                                        if (chkFBRemark == true)
                                            UseRow++;

                                        if (UseRow > 15)
                                        {
                                            objSheet.Rows[UseRow].Insert();
                                            objSheet.Range[objSheet.Cells[UseRow, 2], objSheet.Cells[UseRow, 3]].Merge();
                                            LastRow = UseRow;
                                        }
                                        objSheet.Cells[UseRow, 2] = GarmentParts.ToUpper().Trim();
                                        objSheet.Cells[UseRow, 2].Font.Size = 14;
                                        objSheet.Cells[UseRow, 2].Font.Color = Color.Black;
                                        if (blankCol <= LastCol)
                                        {
                                            objSheet.Cells[UseRow, blankCol] = "สีผ้า " + GarmentParts;
                                        }
                                        chkParts = OIDGParts;

                                        //** FB CODE*************************************
                                        chkFBCode = false;
                                        chkFBLot = false;
                                        chkFBRemark = false;
                                        foreach (DataRow drChkFB in chkDtFB.Rows)
                                        {
                                            string OIDGP = drChkFB["OIDGParts"].ToString();
                                            string FBColor = drChkFB["VendorFBCode"].ToString().Trim();
                                            FBColor = FBColor.Length > 0 ? FBColor.Substring(0, FBColor.Length - 1) : "";
                                            if (OIDGParts == OIDGP && FBColor != "")
                                            {
                                                chkFBCode = true;
                                                break;
                                            }
                                        }

                                        if (chkFBCode == true)
                                        {
                                            objSheet.Rows[UseRow + 1].Insert();
                                            objSheet.Range[objSheet.Cells[UseRow + 1, 2], objSheet.Cells[UseRow + 1, 3]].Merge();
                                            objSheet.Cells[UseRow + 1, 2] = "FABRIC ITEM";
                                            if (blankCol <= LastCol)
                                            {
                                                objSheet.Cells[UseRow + 1, blankCol] = "รหัสผ้า";
                                            }
                                            LastRow = UseRow + 1;
                                        }

                                        //** FB LOT*************************************
                                        foreach (DataRow drChkFB in chkDtFB.Rows)
                                        {
                                            string OIDGP = drChkFB["OIDGParts"].ToString();
                                            string FBLot = drChkFB["FBLotNo"].ToString().Trim();
                                            FBLot = FBLot.Length > 0 ? FBLot.Substring(0, FBLot.Length - 1) : "";
                                            if (OIDGParts == OIDGP && FBLot != "")
                                            {
                                                chkFBLot = true;
                                                break;
                                            }
                                        }

                                        if (chkFBLot == true)
                                        {
                                            if (chkFBCode == true)
                                            {
                                                objSheet.Rows[UseRow + 2].Insert();
                                                objSheet.Range[objSheet.Cells[UseRow + 2, 2], objSheet.Cells[UseRow + 2, 3]].Merge();
                                                objSheet.Cells[UseRow + 2, 2] = "FABRIC LOT";
                                                if (blankCol <= LastCol)
                                                {
                                                    objSheet.Cells[UseRow + 2, blankCol] = "ล็อตผ้า";
                                                }
                                                LastRow = UseRow + 2;
                                            }
                                            else
                                            {
                                                objSheet.Rows[UseRow + 1].Insert();
                                                objSheet.Range[objSheet.Cells[UseRow + 1, 2], objSheet.Cells[UseRow + 1, 3]].Merge();
                                                objSheet.Cells[UseRow + 1, 2] = "FABRIC LOT";
                                                if (blankCol <= LastCol)
                                                {
                                                    objSheet.Cells[UseRow + 1, blankCol] = "ล็อตผ้า";
                                                }
                                                LastRow = UseRow + 1;
                                            }
                                        }

                                        //** FB Remark*************************************
                                        foreach (DataRow drChkFB in chkDtFB.Rows)
                                        {
                                            string OIDGP = drChkFB["OIDGParts"].ToString();
                                            string FBxRemark = drChkFB["FBRemark"].ToString().Trim();
                                            FBxRemark = FBxRemark.Length > 0 ? FBxRemark.Substring(0, FBxRemark.Length - 1) : "";
                                            if (OIDGParts == OIDGP && FBxRemark != "")
                                            {
                                                chkFBRemark = true;
                                                break;
                                            }
                                        }

                                        if (chkFBRemark == true)
                                        {
                                            if (chkFBCode == true)
                                            {
                                                if (chkFBLot == true)
                                                {
                                                    objSheet.Rows[UseRow + 3].Insert();
                                                    objSheet.Range[objSheet.Cells[UseRow + 3, 2], objSheet.Cells[UseRow + 3, 3]].Merge();
                                                    objSheet.Cells[UseRow + 3, 2] = "FABRIC REMARK";
                                                    if (blankCol <= LastCol)
                                                    {
                                                        objSheet.Cells[UseRow + 3, blankCol] = "หมายเหตุ";
                                                    }
                                                    LastRow = UseRow + 3;
                                                }
                                                else
                                                {
                                                    objSheet.Rows[UseRow + 2].Insert();
                                                    objSheet.Range[objSheet.Cells[UseRow + 2, 2], objSheet.Cells[UseRow + 2, 3]].Merge();
                                                    objSheet.Cells[UseRow + 2, 2] = "FABRIC REMARK";
                                                    if (blankCol <= LastCol)
                                                    {
                                                        objSheet.Cells[UseRow + 2, blankCol] = "หมายเหตุ";
                                                    }
                                                    LastRow = UseRow + 2;
                                                }
                                            }
                                            else if (chkFBLot == true)
                                            {
                                                objSheet.Rows[UseRow + 2].Insert();
                                                objSheet.Range[objSheet.Cells[UseRow + 2, 2], objSheet.Cells[UseRow + 2, 3]].Merge();
                                                objSheet.Cells[UseRow + 2, 2] = "FABRIC REMARK";
                                                if (blankCol <= LastCol)
                                                {
                                                    objSheet.Cells[UseRow + 2, blankCol] = "หมายเหตุ";
                                                }
                                                LastRow = UseRow + 2;
                                            }
                                            else
                                            {
                                                objSheet.Rows[UseRow + 1].Insert();
                                                objSheet.Range[objSheet.Cells[UseRow + 1, 2], objSheet.Cells[UseRow + 1, 3]].Merge();
                                                objSheet.Cells[UseRow + 1, 2] = "FABRIC REMARK";
                                                if (blankCol <= LastCol)
                                                {
                                                    objSheet.Cells[UseRow + 1, blankCol] = "หมายเหตุ";
                                                }
                                                LastRow = UseRow + 1;
                                            }
                                        }

                                        runCell = 4;
                                    }

                                    if (GarmentParts.ToUpper().Trim() == "BODY")
                                    {
                                        string FBFile = drFB["FBFile"].ToString().Trim();
                                        FBFile = FBFile.Length > 0 ? FBFile.Substring(0, FBFile.Length - 1) : "";
                                        FBFile = FBFile.Trim().Replace(" ", "");
                                        if (FBFile.Length > 0)
                                        {
                                            if (FBFile.Substring(0, 1) == ",")
                                                FBFile = FBFile.Substring(1);
                                        }

                                        if (FBFile.Length > 0)
                                        {
                                            if (FBFile.Substring(FBFile.Length - 1, 1) == ",")
                                                FBFile = FBFile.Substring(0, FBFile.Length - 1);
                                        }

                                        //objSheet.Cells[12, runCell] = FBFile;
                                        if (FBFile != "")
                                        {
                                            if (FBFile.IndexOf(',') > -1) //มากกว่า 1 รูป
                                            {
                                                string[] imgFile = FBFile.Split(',');
                                                int TTImg = imgFile.Length;

                                                for (int i = 0; i < imgFile.Length; i++)
                                                {
                                                    string FabricFile = "";
                                                    if (imgFile[i] != "")
                                                    {
                                                        if (imgFile[i].IndexOf('/') > -1)
                                                            FabricFile = imgFile[i];
                                                        else
                                                            FabricFile = this.imgPath + imgFile[i];
                                                    }

                                                    if (FabricFile != "")
                                                    {
                                                        oRange = (Microsoft.Office.Interop.Excel.Range)objSheet.Cells[12, runCell];
                                                        if (i == 0)
                                                            Left = (float)((double)oRange.Left) + 1;
                                                        else
                                                            Left = (float)(((double)oRange.Left + 1 + ((double)(oRange.Width / TTImg) * i)));
                                                        Top = (float)((double)oRange.Top) + 1;
                                                        objSheet.Shapes.AddPicture(System.IO.Path.Combine(FabricFile), Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, Left, Top, (int)(oRange.Width - 2) / TTImg, oRange.Height - 2);
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                string FabricFile = "";
                                                if (FBFile != "")
                                                {
                                                    if (FBFile.IndexOf('/') > -1)
                                                        FabricFile = FBFile;
                                                    else
                                                        FabricFile = this.imgPath + FBFile;
                                                }

                                                if (FabricFile != "")
                                                {
                                                    oRange = (Microsoft.Office.Interop.Excel.Range)objSheet.Cells[12, runCell];
                                                    Left = (float)((double)oRange.Left) + 1;
                                                    Top = (float)((double)oRange.Top) + 1;
                                                    objSheet.Shapes.AddPicture(System.IO.Path.Combine(FabricFile), Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, Left, Top, oRange.Width - 2, oRange.Height - 2);
                                                }
                                            }
                                        }
                                    }

                                    string FBCodeColor = drFB["FBCodeColor"].ToString().Trim();
                                    FBCodeColor = FBCodeColor.Length > 0 ? FBCodeColor.Substring(0, FBCodeColor.Length - 1) : "";

                                    string VendorFBCode = drFB["VendorFBCode"].ToString().Trim();
                                    VendorFBCode = VendorFBCode.Length > 0 ? VendorFBCode.Substring(0, VendorFBCode.Length - 1) : "";

                                    string FBLotNo = drFB["FBLotNo"].ToString().Trim();
                                    FBLotNo = FBLotNo.Length > 0 ? FBLotNo.Substring(0, FBLotNo.Length - 1) : "";

                                    string FBRemark = drFB["FBRemark"].ToString().Trim();
                                    FBRemark = FBRemark.Length > 0 ? FBRemark.Substring(0, FBRemark.Length - 1) : "";

                                    Supplier = drFB["Vendor"].ToString().Trim();
                                    Supplier = Supplier.Length > 0 ? Supplier.Substring(0, Supplier.Length - 1) : "";

                                    Composition = drFB["Composition"].ToString().Trim();
                                    Composition = Composition.Length > 0 ? Composition.Substring(0, Composition.Length - 1) : "";

                                    if (ChkComposition != Composition)
                                    {
                                        string xSupplier = "";
                                        if (Supplier != ChkSupplier)
                                        {
                                            xSupplier = Supplier;
                                            ChkSupplier = Supplier;
                                        }
                                        dtComposition.Rows.Add(Composition, xSupplier);
                                        ChkComposition = Composition;
                                    }

                                    objSheet.Cells[UseRow, runCell] = FBCodeColor;

                                    if (chkFBCode == true)
                                    {
                                        objSheet.Cells[UseRow + 1, runCell] = VendorFBCode;
                                    }

                                    if (chkFBLot == true)
                                    {
                                        if (chkFBCode == true)
                                        {
                                            objSheet.Cells[UseRow + 2, runCell] = FBLotNo;
                                        }
                                        else
                                        {
                                            objSheet.Cells[UseRow + 1, runCell] = FBLotNo;
                                        }
                                    }

                                    if (chkFBRemark == true)
                                    {
                                        if (chkFBCode == true)
                                        {
                                            if (chkFBLot == true)
                                            {
                                                objSheet.Cells[UseRow + 3, runCell] = FBRemark;
                                            }
                                            else
                                            {
                                                objSheet.Cells[UseRow + 2, runCell] = FBRemark;
                                            }
                                        }
                                        else if (chkFBLot == true)
                                        {
                                            objSheet.Cells[UseRow + 2, runCell] = FBRemark;
                                        }
                                        else
                                        {
                                            objSheet.Cells[UseRow + 1, runCell] = FBRemark;
                                        }
                                    }


                                    runCell++;
                                    runLoop++;
                                }

                                for (int xi = BeginFB; xi <= LastRow; xi++)
                                {
                                    if (blankCol <= LastCol)
                                    {
                                        objSheet.Range[objSheet.Cells[xi, blankCol], objSheet.Cells[xi, LastCol]].Merge();
                                        objSheet.Range[objSheet.Cells[xi, blankCol], objSheet.Cells[xi, LastCol]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                                        objSheet.Cells[xi, blankCol].Font.Size = 13;

                                        if (objSheet.Cells[xi, 2].Value.ToString() == "FABRIC ITEM" || objSheet.Cells[xi, 2].Value.ToString() == "FABRIC LOT" || objSheet.Cells[xi, 2].Value.ToString() == "FABRIC REMARK")
                                        {
                                            objSheet.Range[objSheet.Cells[xi, 2], objSheet.Cells[xi, LastCol]].Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDot;
                                            objSheet.Range[objSheet.Cells[xi, 2], objSheet.Cells[xi, LastCol]].Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                                            if (objSheet.Cells[xi, 2].Value.ToString() == "FABRIC REMARK")
                                            {
                                                objSheet.Range[objSheet.Cells[xi, 2], objSheet.Cells[xi, LastCol]].Font.Size = 11;
                                                objSheet.Range[objSheet.Cells[xi, 2], objSheet.Cells[xi, LastCol]].Font.Color = Color.Red;
                                                objSheet.Range[objSheet.Cells[xi, 2], objSheet.Cells[xi, LastCol]].Rows.AutoFit();
                                            }
                                        }
                                        else
                                        {
                                            if (xi == BeginFB + 1)
                                            {
                                                objSheet.Range[objSheet.Cells[xi, 2], objSheet.Cells[xi, LastCol + 3]].Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                                                objSheet.Range[objSheet.Cells[xi, 2], objSheet.Cells[xi, LastCol + 3]].Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick;
                                            }
                                            else
                                            {
                                                objSheet.Range[objSheet.Cells[xi, 2], objSheet.Cells[xi, LastCol]].Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                                                objSheet.Range[objSheet.Cells[xi, 2], objSheet.Cells[xi, LastCol]].Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                                            }
                                        }

                                    }
                                }

                                BeginMT = LastRow;



                                //**** Fabric Composition ****
                                LastRow += 3;
                                BeginFBComp = LastRow;
                                if (dtComposition.Rows.Count > 0)
                                {
                                    int chkRow = 0;
                                    foreach (DataRow drComp in dtComposition.Rows)
                                    {
                                        if (chkRow > 2)
                                        {
                                            objSheet.Rows[BeginFBComp].Insert();
                                            objSheet.Range[objSheet.Cells[BeginFBComp - 1, 1], objSheet.Cells[BeginFBComp, 1]].Merge();
                                            objSheet.Range[objSheet.Cells[BeginFBComp, 2], objSheet.Cells[BeginFBComp, LastCol + 3]].Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                        }
                                        objSheet.Cells[BeginFBComp, 2] = drComp[0].ToString();
                                        objSheet.Cells[BeginFBComp, LastCol + 1] = drComp[1].ToString();
                                        LastRow++;
                                        BeginFBComp++;
                                        chkRow++;
                                    }

                                }

                            }

                            pbcEXPORT.PerformStep();
                            pbcEXPORT.Update();

                            //**** Material ****
                            sbSQL.Clear();
                            sbSQL.Append("SELECT    RQ.OIDSMPLDT, MT.OIDITEM, MT.Code, MT.Description, ");
                            sbSQL.Append("          ISNULL((SELECT VendMTCode + ',' AS 'data()' FROM(SELECT DISTINCT B.VendMTCode FROM SMPLRequestMaterial AS B INNER JOIN Items AS YB ON B.OIDITEM = YB.OIDITEM WHERE B.OIDSMPLDT = RQ.OIDSMPLDT  AND YB.OIDITEM = MT.OIDITEM) AS XB FOR XML PATH('')), '')  AS VendorMTCode, ");
                            sbSQL.Append("          ISNULL((SELECT Name + ',' AS 'data()' FROM(SELECT DISTINCT REPLACE(REPLACE(VD.Name, ' CO.,LTD.', ''), ' CO.,LTD', '') AS Name FROM SMPLRequestMaterial AS D INNER JOIN Items AS YD ON D.OIDITEM = YD.OIDITEM INNER JOIN Vendor AS VD ON D.OIDVEND = VD.OIDVEND WHERE D.OIDSMPLDT = RQ.OIDSMPLDT  AND YD.OIDITEM = MT.OIDITEM) AS XD FOR XML PATH('')), '')  AS Vendor, ");
                            sbSQL.Append("          ISNULL((SELECT CASE WHEN ISNULL(Composition, '') = '' THEN '' ELSE Composition + ',' END AS 'data()' FROM(SELECT DISTINCT F.Composition FROM SMPLRequestMaterial AS F INNER JOIN Items AS YF ON F.OIDITEM = YF.OIDITEM WHERE F.OIDSMPLDT = RQ.OIDSMPLDT  AND YF.OIDITEM = MT.OIDITEM) AS XF FOR XML PATH('')), '')  AS Composition, ");
                            sbSQL.Append("          ISNULL((SELECT CASE WHEN ISNULL(Situation, '') = '' THEN '' ELSE Situation + ',' END AS 'data()' FROM(SELECT DISTINCT G.Situation FROM SMPLRequestMaterial AS G INNER JOIN Items AS YG ON G.OIDITEM = YG.OIDITEM WHERE G.OIDSMPLDT = RQ.OIDSMPLDT  AND YG.OIDITEM = MT.OIDITEM) AS XG FOR XML PATH('')), '')  AS Situation, ");
                            sbSQL.Append("          ISNULL((SELECT CASE WHEN ISNULL(Comment, '') = '' THEN '' ELSE Comment + ',' END AS 'data()' FROM(SELECT DISTINCT H.Comment FROM SMPLRequestMaterial AS H INNER JOIN Items AS YH ON H.OIDITEM = YH.OIDITEM WHERE H.OIDSMPLDT = RQ.OIDSMPLDT  AND YH.OIDITEM = MT.OIDITEM) AS XH FOR XML PATH('')), '')  AS Comment, ");
                            sbSQL.Append("          ISNULL((SELECT CASE WHEN ISNULL(Remark, '') = '' THEN '' ELSE Remark + ',' END AS 'data()' FROM(SELECT DISTINCT I.Remark FROM SMPLRequestMaterial AS I INNER JOIN Items AS YI ON I.OIDITEM = YI.OIDITEM WHERE I.OIDSMPLDT = RQ.OIDSMPLDT  AND YI.OIDITEM = MT.OIDITEM) AS XI FOR XML PATH('')), '')  AS Remark ");
                            sbSQL.Append("FROM      SMPLQuantityRequired AS RQ CROSS JOIN ");
                            sbSQL.Append("          (SELECT DISTINCT IT.OIDITEM, IT.Code, IT.Description ");
                            sbSQL.Append("           FROM   SMPLRequestMaterial AS SMT INNER JOIN ");
                            sbSQL.Append("                  Items AS IT ON SMT.OIDITEM = IT.OIDITEM ");
                            sbSQL.Append("           WHERE  (SMT.OIDSMPLDT IN ");
                            sbSQL.Append("                             (SELECT OIDSMPLDT ");
                            sbSQL.Append("                              FROM   SMPLQuantityRequired AS A ");
                            sbSQL.Append("                              WHERE  (OIDSMPL = '" + SMPL[0] + "')))) AS MT ");
                            sbSQL.Append("WHERE (RQ.OIDSMPL = '" + SMPL[0] + "') ");
                            sbSQL.Append("ORDER BY MT.OIDITEM, RQ.OIDSIZE, RQ.OIDCOLOR ");
                            DataTable dtMT = DBC.DBQuery(sbSQL.ToString()).getDataTable();
                            if (dtMT.Rows.Count > 0)
                            {
                                LastRow = BeginMT;
                                UseRow = BeginMT;
                                string chkITEM = "";

                                string MTSupplier = "";
                                string MTComposition = "";

                                string ChkMTComposition = "";
                                string ChkMTSupplier = "";

                                DataTable dtMTComposition = new DataTable();
                                dtMTComposition.Columns.Add("Composition", typeof(String));
                                dtMTComposition.Columns.Add("Supplier", typeof(String));

                                int runCell = 4;
                                int runLoop = 0;
                                foreach (DataRow drMT in dtMT.Rows)
                                {
                                    string OIDITEM = drMT["OIDITEM"].ToString();
                                    string Code = drMT["Code"].ToString();
                                    string Description = drMT["Description"].ToString();

                                    if (chkITEM != OIDITEM)
                                    {
                                        UseRow++;
                                        if (UseRow > 15)
                                        {
                                            objSheet.Rows[UseRow].Insert();
                                            objSheet.Range[objSheet.Cells[UseRow, 2], objSheet.Cells[UseRow, 3]].Merge();
                                            LastRow = UseRow;
                                        }
                                        objSheet.Cells[UseRow, 2] = Description.ToUpper().Trim();
                                        objSheet.Cells[UseRow, 2].Font.Size = 14;
                                        objSheet.Cells[UseRow, 2].Font.Color = Color.Black;
                                        //if (blankCol <= LastCol)
                                        //    objSheet.Cells[UseRow, blankCol] = "สีผ้า " + GarmentParts;
                                        chkITEM = OIDITEM;

                                        runCell = 4;
                                    }

                                    string VendorMTCode = drMT["VendorMTCode"].ToString().Trim();
                                    VendorMTCode = VendorMTCode.Length > 0 ? VendorMTCode.Substring(0, VendorMTCode.Length - 1) : "";
                                    objSheet.Cells[UseRow, runCell] = VendorMTCode;

                                    string Situation = drMT["Situation"].ToString().Trim();
                                    Situation = Situation.Length > 0 ? Situation.Substring(0, Situation.Length - 1) : "";

                                    //if (Situation != "")
                                    //{
                                    //    if (blankCol <= LastCol)
                                    //    {
                                    //        objSheet.Cells[UseRow, blankCol] = Situation;
                                    //    }
                                    //}

                                    string Comment = drMT["Comment"].ToString().Trim();
                                    Comment = Comment.Length > 0 ? Comment.Substring(0, Comment.Length - 1) : "";
                                    string Remark = drMT["Remark"].ToString().Trim();
                                    Remark = Remark.Length > 0 ? Remark.Substring(0, Remark.Length - 1) : "";
                                    string Recommend = "";
                                    if (Situation != "")
                                        Recommend += Situation;
                                    if (Comment != "")
                                    {
                                        if (Recommend != "")
                                            Recommend += " / ";
                                        Recommend += Comment;
                                    }
                                    if (Remark != "")
                                    {
                                        if (Recommend != "")
                                            Recommend += " / ";
                                        Recommend += Remark;
                                    }
                                    objSheet.Range[objSheet.Cells[UseRow, LastCol + 1], objSheet.Cells[UseRow, LastCol + 3]].Merge();
                                    objSheet.Range[objSheet.Cells[UseRow, LastCol + 1], objSheet.Cells[UseRow, LastCol + 3]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                                    objSheet.Cells[UseRow, LastCol + 1] = Recommend;

                                    MTSupplier = drMT["Vendor"].ToString().Trim();
                                    MTSupplier = MTSupplier.Length > 0 ? MTSupplier.Substring(0, MTSupplier.Length - 1) : "";

                                    MTComposition = drMT["Composition"].ToString().Trim();
                                    MTComposition = MTComposition.Length > 0 ? MTComposition.Substring(0, MTComposition.Length - 1) : "";

                                    if (ChkMTComposition != MTComposition)
                                    {
                                        string xSupplier = "";
                                        if (MTSupplier != ChkMTSupplier)
                                        {
                                            xSupplier = MTSupplier;
                                            ChkMTSupplier = MTSupplier;
                                        }
                                        dtMTComposition.Rows.Add(MTComposition, xSupplier);
                                        ChkMTComposition = MTComposition;
                                    }

                                    runCell++;
                                    runLoop++;
                                }


                                objSheet.Range[objSheet.Cells[BeginMT + 1, 2], objSheet.Cells[BeginMT + 1, LastCol + 3]].Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                                objSheet.Range[objSheet.Cells[BeginMT + 1, 2], objSheet.Cells[BeginMT + 1, LastCol + 3]].Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick;

                                for (int xi = BeginMT + 1; xi <= LastRow; xi++)
                                {
                                    if (blankCol <= LastCol)
                                    {
                                        objSheet.Range[objSheet.Cells[xi, blankCol], objSheet.Cells[xi, LastCol]].Merge();
                                        objSheet.Range[objSheet.Cells[xi, blankCol], objSheet.Cells[xi, LastCol]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                                        objSheet.Cells[xi, blankCol].Font.Size = 13;

                                        //objSheet.Range[objSheet.Cells[xi, 2], objSheet.Cells[xi, 10]].Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                                        //objSheet.Range[objSheet.Cells[xi, 2], objSheet.Cells[xi, 10]].Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick;

                                    }
                                }

                                //**** Material Composition ****
                                LastRow += 6;
                                BeginMTComp = LastRow;
                                if (dtMTComposition.Rows.Count > 0)
                                {
                                    int chkRow = 0;
                                    foreach (DataRow drComp in dtMTComposition.Rows)
                                    {
                                        if (chkRow > 2)
                                        {
                                            objSheet.Rows[BeginMTComp].Insert();
                                            objSheet.Range[objSheet.Cells[BeginMTComp - 1, 1], objSheet.Cells[BeginMTComp, 1]].Merge();
                                            objSheet.Range[objSheet.Cells[BeginMTComp, 2], objSheet.Cells[BeginMTComp, LastCol + 3]].Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                        }
                                        objSheet.Cells[BeginMTComp, 2] = drComp[0].ToString();
                                        objSheet.Cells[BeginMTComp, LastCol + 1] = drComp[1].ToString();
                                        BeginMTComp++;
                                        LastRow++;
                                        chkRow++;
                                    }
                                    LastRow--;
                                }
                            }

                            pbcEXPORT.PerformStep();
                            pbcEXPORT.Update();

                            sbSQL.Clear();
                            sbSQL.Append("SELECT DISTINCT DT.OIDITEM, DT.Code, DT.Description, DT.MaterialType, PS.SizeName + ' (' + PC.ColorName + ')' AS SizeColor, PC.OIDCOLOR, PS.OIDSIZE, ");
                            sbSQL.Append("       ISNULL((SELECT GarmentParts + ', ' AS 'data()' FROM(SELECT xGP.OIDGParts, xGP.GarmentParts FROM SMPLRequestFabricParts AS xSFBP INNER JOIN GarmentParts AS xGP ON xSFBP.OIDGParts = xGP.OIDGParts AND xSFBP.OIDSMPLFB = SFB.OIDSMPLFB) AS XB FOR XML PATH('')), '')  AS FabricParts ");
                            sbSQL.Append("FROM   SMPLQuantityRequired AS SQR INNER JOIN ");
                            sbSQL.Append("       ProductColor AS PC ON SQR.OIDCOLOR = PC.OIDCOLOR INNER JOIN ");
                            sbSQL.Append("       ProductSize AS PS ON SQR.OIDSIZE = PS.OIDSIZE INNER JOIN ");
                            sbSQL.Append("       SMPLRequestFabric AS SFB ON SQR.OIDSMPLDT = SFB.OIDSMPLDT INNER JOIN ");
                            sbSQL.Append("       (SELECT DISTINCT SRFB.OIDITEM, ITM.Code, ITM.Description, ITM.MaterialType ");
                            sbSQL.Append("        FROM   SMPLRequestFabric AS SRFB INNER JOIN ");
                            sbSQL.Append("               Items AS ITM ON SRFB.OIDITEM = ITM.OIDITEM ");
                            sbSQL.Append("        WHERE  (SRFB.OIDSMPLDT IN ");
                            sbSQL.Append("                  (SELECT OIDSMPLDT ");
                            sbSQL.Append("                   FROM   SMPLQuantityRequired ");
                            sbSQL.Append("                   WHERE  (OIDSMPL = '" + SMPL[0] + "')))) AS DT ON (SQR.OIDSMPL = '" + SMPL[0] + "') AND (SFB.OIDITEM = DT.OIDITEM) ");
                            sbSQL.Append("ORDER BY DT.OIDITEM, PC.OIDCOLOR, PS.OIDSIZE ");
                            DataTable dtITEM = DBC.DBQuery(sbSQL.ToString()).getDataTable();
                            if (dtITEM.Rows.Count > 0)
                            {
                                string chkITEM = "";
                                StringBuilder sbITEM = new StringBuilder();
                                int runLoop = 0;
                                foreach (DataRow drITEM in dtITEM.Rows)
                                {
                                    string ID = drITEM["OIDITEM"].ToString();
                                    string Code = drITEM["Code"].ToString();
                                    string Description = drITEM["Description"].ToString();
                                    string MaterialType = drITEM["MaterialType"].ToString();
                                    string SizeColor = drITEM["SizeColor"].ToString();
                                    string FabricParts = drITEM["FabricParts"].ToString().Trim();
                                    FabricParts = FabricParts.Length > 0 ? FabricParts.Substring(0, FabricParts.Length - 1) : "";

                                    if (chkITEM != ID)
                                    {
                                        chkITEM = ID;
                                        if (runLoop > 0)
                                            sbITEM.Append("\n");

                                        if (MaterialType != "8")
                                            sbITEM.Append(Code + " : " + Description + "\n");
                                        else
                                            sbITEM.Append(Description + "\n");

                                    }
                                    sbITEM.Append(" ※ " + SizeColor + " -> " + FabricParts + "\n");

                                    runLoop++;
                                }

                                if (sbITEM.Length > 0)
                                {
                                    objSheet.Cells[11, LastCol + 1] = "FABRIC CODE";
                                    objSheet.Cells[11, LastCol + 1].Font.Size = 11;
                                    objSheet.Range[objSheet.Cells[12, LastCol + 1], objSheet.Cells[BeginMT, LastCol + 3]].Merge();
                                    objSheet.Range[objSheet.Cells[12 + 1, LastCol + 1], objSheet.Cells[BeginMT, LastCol + 3]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                                    objSheet.Range[objSheet.Cells[12 + 1, LastCol + 1], objSheet.Cells[BeginMT, LastCol + 3]].verticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
                                    objSheet.Cells[12, LastCol + 1] = sbITEM.ToString();
                                }
                            }


                            objWorkBook.SaveAs(sFilePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            objApp.Workbooks.Close();
                            chkExcel = true;

                        }
                        catch (Exception)
                        {
                            //Error Alert
                            chkExcel = false;
                        }
                        finally
                        {
                            objApp.Quit();
                            objWorkBook = null;
                            objApp = null;
                        }

                        pbcEXPORT.PerformStep();
                        pbcEXPORT.Update();


                        if (chkExcel == true)
                        {
                            //System.Diagnostics.Process.Start(sFilePath);
                            DevExpress.XtraSpreadsheet.SpreadsheetControl spsForecast = new DevExpress.XtraSpreadsheet.SpreadsheetControl();
                            IWorkbook workbook = spsForecast.Document;
                            string ext = "";
                            using (FileStream stream = new FileStream(sFilePath, FileMode.Open))
                            {
                                // workbook.CalculateFull();
                                ext = Path.GetExtension(sFilePath);
                                if (ext == ".xlsx")
                                    workbook.LoadDocument(stream, DocumentFormat.Xlsx);
                                else if (ext == ".xls")
                                    workbook.LoadDocument(stream, DocumentFormat.Xls);
                                else if (ext == ".csv")
                                    workbook.LoadDocument(stream, DocumentFormat.Csv);

                            }

                            string pdfPathFile = sFilePath.Replace(ext, ".pdf");

                            using (FileStream pdfFileStream = new FileStream(pdfPathFile, FileMode.Create))
                            {
                                workbook.ExportToPdf(pdfFileStream);
                                pbcEXPORT.PerformStep();
                                pbcEXPORT.Update();

                                System.Diagnostics.Process.Start(pdfPathFile);
                                pbcEXPORT.PerformStep();
                                pbcEXPORT.Update();
                            }

                        }
                        //****** END EXPORT *******
                    }
                    else
                    {
                        FUNCT.msgError("ไม่พบข้อมูลเอกสาร Sample Request: " + SMPLNo);
                    }

                    layoutControlItem120.Text = "Status ..";
                    layoutControlItem120.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                }

            }
        }



        //Open File Image
        public void openFile_Image(XtraOpenFileDialog xopen, TextEdit txt, PictureEdit pic)
        {
            //string fileName = string.Empty;
            xopen.Filter = "Image files | *.jpg; *.jpeg; *.jpe; *.jfif; *.png";
            if (xopen.ShowDialog() == DialogResult.OK)
            {
                string filename = xopen.FileName;
                txt.Text = filename;
                pic.Image = Image.FromFile(filename);
                pic.Properties.SizeMode = DevExpress.XtraEditors.Controls.PictureSizeMode.Zoom;
            }
            //return fileName;
        }

        //Upload Image
        public string uploadImg(TextEdit txt, string newFilenames)
        {
            string imgName = txt.Text.ToString().Trim().Replace("'", "''");
            string newFileName = "null";
            if (imgName != "")
            {
                try
                {
                    string path = imgPath;
                    string filename = imgName;
                    string extension = Path.GetExtension(filename);
                    Random generator = new Random();
                    string r = generator.Next(0, 999999).ToString("D4");
                    newFileName = newFilenames + "-" + DateTime.Now.ToString("yyyyMMdd") + "-" + r + extension;
                    File.Copy(filename, path + Path.GetFileName(newFileName));
                    //MessageBox.Show("Upload Files is Successfull.", "Upload Status");
                }
                catch (Exception)
                {
                    FUNCT.msgWarning("Uplaod ไม่ได้ เนื่องจากมีไฟล์นี้ใน Directory ปัจจุบันแล้ว!");
                }
            }
            //if (newFileName != "null")
            //{
            //    newFileName = "N'"+ newFileName + "'";
            //}
            return newFileName;
        }

        public string uploadImg(string txt, string newFilenames)
        {
            string imgName = txt.Trim().Replace("'", "''");
            string newFileName = "null";
            if (imgName != "")
            {
                try
                {
                    string path = imgPath;
                    string filename = imgName;
                    string extension = Path.GetExtension(filename);
                    Random generator = new Random();
                    string r = generator.Next(0, 999999).ToString("D4");
                    newFileName = newFilenames + "-" + DateTime.Now.ToString("yyyyMMdd") + "-" + r + extension;
                    File.Copy(filename, path + Path.GetFileName(newFileName));
                    //MessageBox.Show("Upload Files is Successfull.", "Upload Status");
                }
                catch (Exception)
                {
                    FUNCT.msgWarning("Uplaod ไม่ได้ เนื่องจากมีไฟล์นี้ใน Directory ปัจจุบันแล้ว!");
                }
            }
            //if (newFileName != "null")
            //{
            //    newFileName = "N'" + newFileName + "'";
            //}
            return newFileName;
        }

        private void sbRefreshCS_Click(object sender, EventArgs e)
        {
            LoadSizeColor();
        }
    }
}