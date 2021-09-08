using System;
using System.Windows.Forms;
using DevExpress.LookAndFeel;
using DevExpress.XtraGrid.Views.Grid;
using System.Data.SqlClient;
using DevExpress.Utils;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using System.Data;
using System.Collections;
using DevExpress.XtraGrid;
using System.ComponentModel;
using DevExpress.XtraGrid.Columns;
using System.Collections.Generic;
using System.Linq;
using TheepClass;
using DBConnect;
using System.Text;
using System.Drawing;
using DevExpress.XtraEditors;
using DevExpress.XtraGrid.Views.Base;
using System.IO;
using DevExpress.Spreadsheet;

namespace MDS.Development
{
    public partial class DEV02 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private const string COMPANY_CODE = "PTC";
        // Global Var
        goClass.dbConn db = new goClass.dbConn();
        goClass.ctool ct = new goClass.ctool();
        classHardQuery hq = new classHardQuery();
        SqlConnection mainConn = new goClass.dbConn().MDS();

        public LogIn UserLogin { get; set; }
        
        public string ConnectionString { get; set; }
        string CONNECT_STRING = "";
        DatabaseConnect DBC;

        string reportPath = @"\\192.168.101.3\Software_tuw\PTC-MDS\Report\";

        int chkReadWrite = 0;
        private Functionality.Function FUNC = new Functionality.Function();
        public DEV02()
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
            rgDocActive.EditValue = 1;
            rgDocUser.EditValue = 0;

            lblUser.Text = "Login : " + UserLogin.FullName;
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT TOP (1) ReadWriteStatus FROM FunctionAccess WHERE (OIDUser = '" + UserLogin.OIDUser + "') AND(FunctionNo = 'DEV02') ");
            chkReadWrite = DBC.DBQuery(sbSQL).getInt();

            //MessageBox.Show(chkReadWrite.ToString());
            if (chkReadWrite == 0)
            {
                ribbonPageGroup1.Visible = false;
                rpgManage.Visible = false;

                sbClear.Enabled = false;
            }

            sbSQL.Clear();
            sbSQL.Append("SELECT FullName, OIDUSER FROM Users ORDER BY OIDUSER ");
            new ObjDE.setGridLookUpEdit(glueCreateBy, sbSQL, "FullName", "OIDUSER").getData();
            new ObjDE.setGridLookUpEdit(glueUpdateBy, sbSQL, "FullName", "OIDUSER").getData();

            tabMARKING.SelectedTabPageIndex = 0;
            LoadData();   //Default Load Form
            NewData();    //Clear Default Data

        }

        private void LoadUserSMPL()
        {
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT    ID, UserAccount, Branch, Department ");
            sbSQL.Append("FROM      (SELECT 0 AS ID, N'All Users' AS UserAccount, N'' AS Branch, N'' AS Department ");
            sbSQL.Append("           UNION ALL ");
            sbSQL.Append("           SELECT DISTINCT US.OIDUSER AS ID, US.FullName AS UserAccount, BN.Name AS Branch, DP.Name AS Department ");
            sbSQL.Append("           FROM   SMPLRequest AS SRQ INNER JOIN ");
            sbSQL.Append("                  Users AS US ON SRQ.CreatedBy = US.OIDUSER LEFT OUTER JOIN ");
            sbSQL.Append("                  Branchs AS BN ON US.OIDBranch = BN.OIDBranch LEFT OUTER JOIN ");
            sbSQL.Append("                  Departments AS DP ON US.OIDDEPT = DP.OIDDEPT ");
            sbSQL.Append("           WHERE  (SRQ.Status = '0') ");
            sbSQL.Append("           AND    (SRQ.SMPLStatus = '1') ");
            sbSQL.Append("           AND    (SRQ.OIDSMPL IN ");
            sbSQL.Append("                        (SELECT DISTINCT SQR.OIDSMPL ");
            sbSQL.Append("                         FROM   SMPLQuantityRequired AS SQR INNER JOIN ");
            sbSQL.Append("                                SMPLRequestFabric AS SFB ON SQR.OIDSMPLDT = SFB.OIDSMPLDT INNER JOIN ");
            sbSQL.Append("                                SMPLRequestFabricParts AS SFBP ON SFB.OIDSMPLFB = SFBP.OIDSMPLFB AND SFB.OIDSMPLDT = SFBP.OIDSMPLDT))) AS USDATA ");
            sbSQL.Append("ORDER BY Branch, Department, UserAccount ");
            new ObjDE.setGridLookUpEdit(glueUSER, sbSQL, "UserAccount", "ID").getData();
            glueUSER.Properties.View.PopulateColumns(glueUSER.Properties.DataSource);
            glueUSER.Properties.View.Columns["ID"].Visible = false;

            DataTable dtUSER = glueUSER.Properties.DataSource as DataTable;
            if (dtUSER != null)
            {
                bool chkHAS = false;
                foreach (DataRow drUSER in dtUSER.Rows)
                {
                    string UID = drUSER["ID"].ToString();
                    if (UID == UserLogin.OIDUser.ToString())
                    {
                        chkHAS = true;
                        glueUSER.EditValue = UserLogin.OIDUser;
                        break;
                    }
                }
                if(chkHAS == false)
                    glueUSER.EditValue = 0;
            }
            else
            {
                glueUSER.EditValue = 0;
            }
        }

        private void LoadSMPL()
        {
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT DISTINCT  ");
            sbSQL.Append("       SRQ.OIDSMPL AS ID, SRQ.SMPLNo AS [SMPL No.], CASE WHEN SRQ.SMPLRevise > 0 THEN CONVERT(VARCHAR, SRQ.SMPLRevise) ELSE '' END AS Revise, CONVERT(VARCHAR(10), SRQ.RequestDate, 103) AS RequestDate, SRQ.ContactName AS RequestBy, CONVERT(VARCHAR(10), SRQ.DeliveryRequest, 103) AS DeliveryRequest, SRQ.ReferenceNo, SRQ.Season, SRQ.SMPLItem, SRQ.ModelName, (CASE SRQ.PatternSizeZone WHEN 0 THEN 'Japan' WHEN 1 THEN 'Europe' WHEN 2 THEN 'US' END) AS PatternSizeZone, SRQ.SMPLPatternNo AS PatternNo, SUF.UseFor, GC.CategoryName, PS.StyleName,  ");
            sbSQL.Append("       SUBSTRING((SELECT GarmentParts + ', ' AS 'data()' FROM(SELECT DISTINCT GP.OIDGParts, GP.GarmentParts FROM SMPLQuantityRequired AS SQR INNER JOIN SMPLRequestFabric AS SFB ON SQR.OIDSMPLDT = SFB.OIDSMPLDT AND SQR.OIDSMPL = SRQ.OIDSMPL  INNER JOIN SMPLRequestFabricParts AS SFBP ON SFB.OIDSMPLFB = SFBP.OIDSMPLFB AND SFB.OIDSMPLDT = SFBP.OIDSMPLDT INNER JOIN GarmentParts AS GP ON SFBP.OIDGParts = GP.OIDGParts) AS FParts FOR XML PATH('')), 1, LEN((SELECT GarmentParts + ', ' AS 'data()' FROM(SELECT DISTINCT GP.OIDGParts, GP.GarmentParts FROM SMPLQuantityRequired AS SQR INNER JOIN SMPLRequestFabric AS SFB ON SQR.OIDSMPLDT = SFB.OIDSMPLDT AND SQR.OIDSMPL = SRQ.OIDSMPL  INNER JOIN SMPLRequestFabricParts AS SFBP ON SFB.OIDSMPLFB = SFBP.OIDSMPLFB AND SFB.OIDSMPLDT = SFBP.OIDSMPLDT INNER JOIN GarmentParts AS GP ON SFBP.OIDGParts = GP.OIDGParts) AS FParts FOR XML PATH(''))) -1) AS FabricParts, ");
            sbSQL.Append("       CUS.Name AS Customer, BN.Name AS Branch, DP.Name AS Department, (SELECT TOP(1) FullName FROM Users WHERE OIDUSER = SRQ.CreatedBy) AS CreatedBy, SRQ.CreatedDate, (SELECT TOP(1) FullName FROM Users WHERE OIDUSER = SRQ.UpdatedBy) AS UpdatedBy, SRQ.UpdatedDate ");
            sbSQL.Append("FROM   SMPLRequest AS SRQ INNER JOIN ");
            sbSQL.Append("       SMPLUseFor AS SUF ON SRQ.UseFor = SUF.OIDUF LEFT OUTER JOIN ");
            sbSQL.Append("       GarmentCategory AS GC ON SRQ.OIDCATEGORY = GC.OIDGCATEGORY LEFT OUTER JOIN ");
            sbSQL.Append("       ProductStyle AS PS ON SRQ.OIDSTYLE = PS.OIDSTYLE LEFT OUTER JOIN ");
            sbSQL.Append("       Customer AS CUS ON SRQ.OIDCUST = CUS.OIDCUST LEFT OUTER JOIN ");
            sbSQL.Append("       Branchs AS BN ON SRQ.OIDBranch = BN.OIDBranch LEFT OUTER JOIN ");
            sbSQL.Append("       Departments AS DP ON SRQ.OIDDEPT = DP.OIDDEPT ");
            sbSQL.Append("WHERE (SRQ.Status = '0')  ");
            sbSQL.Append("AND (SRQ.SMPLStatus = '1') ");
            sbSQL.Append("AND (SRQ.OIDSMPL IN(SELECT DISTINCT SQR.OIDSMPL FROM SMPLQuantityRequired AS SQR INNER JOIN SMPLRequestFabric AS SFB ON SQR.OIDSMPLDT = SFB.OIDSMPLDT INNER JOIN SMPLRequestFabricParts AS SFBP ON SFB.OIDSMPLFB = SFBP.OIDSMPLFB AND SFB.OIDSMPLDT = SFBP.OIDSMPLDT)) ");
            if (glueUSER.EditValue.ToString() != "0")
            {
                sbSQL.Append("AND (SRQ.CreatedBy = '" + glueUSER.EditValue.ToString() + "') ");
            }
            sbSQL.Append("ORDER BY SRQ.CreatedDate, SRQ.OIDSMPL ");
            new ObjDE.setGridControl(gcSMPL, gvSMPL, sbSQL).getData(false, false, false, true);
            gvSMPL.Columns["ID"].Visible = false;

            gvSMPL.Columns["Revise"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gvSMPL.Columns["RequestDate"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gvSMPL.Columns["DeliveryRequest"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gvSMPL.Columns["Season"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gvSMPL.Columns["StyleName"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gvSMPL.Columns["PatternSizeZone"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gvSMPL.Columns["Branch"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gvSMPL.Columns["Department"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gvSMPL.Columns["CreatedDate"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gvSMPL.Columns["UpdatedDate"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;

            gvSMPL.Columns[0].Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
            gvSMPL.Columns[1].Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
            gvSMPL.Columns[2].Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
            gvSMPL.Columns[3].Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
        }

        private void LoadData()
        {
            getGrid_MARK(gcMARK, gvMARK, UserLogin.OIDUser, Convert.ToInt32(rgDocActive.EditValue.ToString()), Convert.ToInt32(rgDocUser.EditValue.ToString()));

            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT '0' AS ID, 'for Cost' AS RequestType ");
            sbSQL.Append("UNION ALL ");
            sbSQL.Append("SELECT '1' AS ID, 'for Production' AS RequestType ");
            sbSQL.Append("UNION ALL ");
            sbSQL.Append("SELECT '2' AS ID, 'Cutting Request' AS RequestType ");
            new ObjDE.setGridLookUpEdit(glueMarkingRequestType, sbSQL, "RequestType", "ID").getData();
            glueMarkingRequestType.Properties.View.PopulateColumns(glueMarkingRequestType.Properties.DataSource);
            glueMarkingRequestType.Properties.View.Columns["ID"].Visible = false;

            sbSQL.Clear();
            sbSQL.Append("SELECT '0' AS ID, 'ALL' AS Factory ");
            sbSQL.Append("UNION ALL ");
            sbSQL.Append("SELECT '1' AS ID, 'SAI-4' AS Factory ");
            sbSQL.Append("UNION ALL ");
            sbSQL.Append("SELECT '2' AS ID, 'RAMA2' AS Factory ");
            sbSQL.Append("UNION ALL ");
            sbSQL.Append("SELECT '3' AS ID, 'PTC' AS Factory ");
            sbSQL.Append("UNION ALL ");
            sbSQL.Append("SELECT '4' AS ID, 'YEH' AS Factory ");
            new ObjDE.setGridLookUpEdit(glueCuttingFac, sbSQL, "Factory", "ID").getData();
            glueCuttingFac.Properties.View.PopulateColumns(glueCuttingFac.Properties.DataSource);
            glueCuttingFac.Properties.View.Columns["ID"].Visible = false;

            new ObjDE.setGridLookUpEdit(glueSewingFac, sbSQL, "Factory", "ID").getData();
            glueSewingFac.Properties.View.PopulateColumns(glueSewingFac.Properties.DataSource);
            glueSewingFac.Properties.View.Columns["ID"].Visible = false;

            LoadUserSMPL();
            LoadSMPL();
            
            sbSQL.Clear();
            //sbSQL.Append("SELECT Name AS Branch, OIDBranch AS ID FROM Branchs WHERE OIDCOMPANY = (SELECT TOP(1) OIDCOMPANY FROM Company WHERE Code=N'" + COMPANY_CODE + "') ORDER BY OIDBranch");
            sbSQL.Append("SELECT Name AS Branch, OIDBranch AS ID FROM Branchs WHERE OIDCOMPANY = '" + UserLogin.OIDCompany + "' ORDER BY OIDBranch");
            new ObjDE.setSearchLookUpEdit(slueBranch, sbSQL, "Branch", "ID").getData();
            slueBranch.Properties.View.PopulateColumns(slueBranch.Properties.DataSource);
            slueBranch.Properties.View.Columns["ID"].Visible = false;

            sbSQL.Clear();
            sbSQL.Append("SELECT '0' AS ID, 'Japan' AS PatternSizeZone ");
            sbSQL.Append("UNION ALL ");
            sbSQL.Append("SELECT '1' AS ID, 'EU' AS PatternSizeZone ");
            sbSQL.Append("UNION ALL ");
            sbSQL.Append("SELECT '2' AS ID, 'US' AS PatternSizeZone ");
            new ObjDE.setSearchLookUpEdit(sluePatternSizeZone, sbSQL, "PatternSizeZone", "ID").getData();
            sluePatternSizeZone.Properties.View.PopulateColumns(sluePatternSizeZone.Properties.DataSource);
            sluePatternSizeZone.Properties.View.Columns["ID"].Visible = false;

            sbSQL.Clear();
            sbSQL.Append("SELECT '0' AS ID, 'Standard' AS DetailsType ");
            sbSQL.Append("UNION ALL ");
            sbSQL.Append("SELECT '1' AS ID, 'Positive' AS DetailsType ");
            sbSQL.Append("UNION ALL ");
            sbSQL.Append("SELECT '2' AS ID, 'Negative' AS DetailsType ");
            SearchLookUpEdit slueDetailsType = new SearchLookUpEdit();
            new ObjDE.setSearchLookUpEdit(slueDetailsType, sbSQL, "ColorName", "ID").getData();
            slueDetailsType.Properties.View.PopulateColumns(slueDetailsType.Properties.DataSource);
            slueDetailsType.Properties.View.Columns["ID"].Visible = false;

            repDetailType.DataSource = slueDetailsType.Properties.DataSource;
            repDetailType.DisplayMember = "DetailsType";
            repDetailType.ValueMember = "ID";
            repDetailType.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
            repDetailType.View.PopulateColumns(repDetailType.DataSource);
            repDetailType.View.Columns["ID"].Visible = false;

        }

        private void NewData()
        {
            ClearSMPL();


            //Marking Tab
            lblMarkID.Text = "";
            txeMarkingNo.Text = "";
            dtDocumentDate.EditValue = DateTime.Now;
            glueMarkingRequestType.EditValue = 0;
            glueCuttingFac.EditValue = 0;
            glueSewingFac.EditValue = 0;

            slueBranch.EditValue = "";
            txeSampleRequestNo.Text = "";
            dteRequestDate.EditValue = DateTime.Now;
            //rdoSpecofSize.SelectedIndex = 0;

            txeSeason.Text = "";
            txeCustomer.Text = "";
            txeRequestBy.Text = "";
            dtDeliveryRequest.EditValue = DateTime.Now;
            //rdoUseFor.SelectedIndex = 0;
            mmRemark.Text = "";

            txeItemNo.Text = "";
            txtModelName.Text = "";
            txeCategory.Text = "";
            txeStyleName.Text = "";
            txeSaleSection.Text = "";


            glueCreateBy.EditValue = UserLogin.OIDUser;
            txtCreateDate.EditValue = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            glueUpdateBy.EditValue = UserLogin.OIDUser;
            txtUpdateDate.EditValue = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");


            //Marking Detail Tab
            gcListofFabric.DataSource = null;

            txePatternNo.Text = "";
            //rdoPatternSizeZone.SelectedIndex = 0;
            txeVendFBCode.Text = "";
            txeColor.Text = "";
            txeSampleLotNo.Text = "";
            txeFBType.Text = "";
            gcMDT.DataSource = null;

            txtTotal_Standard.Text = "";
            txtUsable_Standard.Text = "";
            txtWeight_Standard.Text = "";
            gcSTD.DataSource = null;

            txtTotal_Positive.Text = "";
            txtUsable_Positive.Text = "";
            txtWeight_Positive.Text = "";
            gcPOS.DataSource = null;

            txtTotal_Negative.Text = "";
            txtUsable_Negative.Text = "";
            txtWeight_Negative.Text = "";
            gcNEG.DataSource = null;
        }

        public void newMarking()
        {
            //ct.showInfoMessage("new Marking");
            txeSampleRequestNo.EditValue = null;
        }

        public void newMarkingDetail()
        {
            ct.showInfoMessage("new MarkingDetail");
        }

        private void bbiNew_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            LoadData();
            NewData();

            tabMARKING.SelectedTabPage = lcgMark;
        }


        private void bbiSave_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (glueMarkingRequestType.Text.Trim() == "")
            {
                FUNC.msgWarning("Please select marking request type.");
                glueMarkingRequestType.Focus();
            }
            else if (glueCuttingFac.Text.Trim() == "")
            {
                FUNC.msgWarning("Please select cutting factory.");
                glueCuttingFac.Focus();
            }
            else if (glueSewingFac.Text.Trim() == "")
            {
                FUNC.msgWarning("Please select sewing factory.");
                glueSewingFac.Focus();
            }
            else if (lblID.Text == "")
            {
                FUNC.msgWarning("Please select sample request for create marking.");
                gcSMPL.Focus();
            }
            else
            {
                string ACTION = "";
                string msgACTION = "Save";
                if (lblStatus.Text.Trim() == "NEW")
                {
                    ACTION = "NEW";
                    msgACTION = "Save new";
                }
                else if (lblStatus.Text.Trim() == "UPDATE")
                {
                    ACTION = "UPDATE";
                    msgACTION = "Update";
                }

                if (FUNC.msgQuiz("Confirm " + msgACTION.ToLower() + " marking ?") == true)
                {
                    string CreatedBy = UserLogin.OIDUser.ToString() != "" ? UserLogin.OIDUser.ToString() : "0";
                    string CreatedDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    string UpdatedBy = UserLogin.OIDUser.ToString() != "" ? UserLogin.OIDUser.ToString() : "0";
                    string UpdatedDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                    string DocumentDate = dtDocumentDate.Text.Trim() != "" ? "'" + Convert.ToDateTime(dtDocumentDate.EditValue.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL";
                    string MarkingRequestType = glueMarkingRequestType.Text.Trim() != "" ? "'" + glueMarkingRequestType.EditValue.ToString() + "'" : "NULL";
                    string Branch = slueBranch.Text.Trim() != "" ? "'" + slueBranch.EditValue.ToString() + "'" : "NULL";
                    string CuttingFac = glueCuttingFac.Text.Trim() != "" ? "'" + glueCuttingFac.EditValue.ToString() + "'" : "NULL";
                    string SewingFac = glueSewingFac.Text.Trim() != "" ? "'" + glueSewingFac.EditValue.ToString() + "'" : "NULL";

                    string newMarking = "";
                    StringBuilder sbSQL = new StringBuilder();
                    if (ACTION == "NEW")
                    {
                        string DEPCode = this.DBC.DBQuery("SELECT TOP(1) Code FROM Departments WHERE (Name = N'" + txeSaleSection.Text + "') AND (OIDBRANCH = '" + slueBranch.EditValue.ToString() + "') ").getString();

                        StringBuilder sbGEN = new StringBuilder();
                        sbGEN.Append("SELECT TOP (1) FORMAT(CAST(SUBSTRING(MarkingNo, LEN(MarkingNo) - 4, 4) + 1 AS Int), '0000') AS genD4 ");
                        sbGEN.Append("FROM Marking ");
                        sbGEN.Append("WHERE (MarkingNo LIKE N'MK" + txeSeason.Text + DEPCode + "%') AND (LEN(MarkingNo) > 9) ");
                        sbGEN.Append("ORDER BY genD4 DESC ");
                        string strRUN = this.DBC.DBQuery(sbGEN).getString();
                        if (strRUN == "")
                            strRUN = "0001";
                        newMarking = "MK" + txeSeason.Text + DEPCode + strRUN;

                        sbSQL.Append("INSERT INTO Marking(OIDSMPL, Status, MarkingNo, DocumentDate, MarkingType, Branch, CuttingFactory, SewingFactory, Remark, CreatedBy, CreatedDate, UpdatedBy, UpdatedDate) ");
                        sbSQL.Append(" VALUES('" + lblID.Text + "', '1', N'" + newMarking + "', " + DocumentDate + ", " + MarkingRequestType + ", " + Branch + ", " + CuttingFac + ", " + SewingFac + ", N'" + mmRemark.Text + "', '" + CreatedBy + "', '" + CreatedDate + "', '" + UpdatedBy + "', '" + UpdatedDate + "')  ");

                        sbSQL.Append("UPDATE SMPLRequest SET Status = '1' WHERE (OIDSMPL = '" + lblID.Text + "')  ");

                        DataTable dtMDT = (DataTable)gcMDT.DataSource;
                        if (dtMDT != null)
                        {
                            int runRow = 0;
                            foreach (DataRow drMDT in dtMDT.Rows)
                            {
                                string ID = drMDT["ID"].ToString(); 
                                string OIDITEM = drMDT["OIDITEM"].ToString();                       OIDITEM = OIDITEM.Trim() == "" ? "NULL" : "'" + OIDITEM.Trim() + "'";
                                string ItemCode = drMDT["ItemCode"].ToString(); 
                                string ItemDescription = drMDT["ItemDescription"].ToString(); 
                                string SMPLPatternNo = drMDT["SMPLPatternNo"].ToString(); 
                                string FBPartsID = drMDT["FBPartsID"].ToString(); 
                                string FBParts = drMDT["FBParts"].ToString(); 
                                string VendorFBCode = drMDT["VendorFBCode"].ToString(); 
                                string SampleLotNo = drMDT["SampleLotNo"].ToString(); 
                                string DetailsType = drMDT["DetailsType"].ToString(); 
                                string TotalWidth = drMDT["TotalWidth"].ToString();                 TotalWidth = TotalWidth.Trim() == "" ? "0" : TotalWidth;
                                string UsableWidth = drMDT["UsableWidth"].ToString();               UsableWidth = UsableWidth.Trim() == "" ? "0" : UsableWidth;
                                string WeightGM2 = drMDT["WeightGM2"].ToString();                   WeightGM2 = WeightGM2.Trim() == "" ? "0" : WeightGM2;
                                string OIDSIZE = drMDT["OIDSIZE"].ToString();                       OIDSIZE = OIDSIZE.Trim() == "" ? "NULL" : "'" + OIDSIZE.Trim() + "'";
                                string SizeName = drMDT["SizeName"].ToString(); 
                                string ActualLengthCm = drMDT["ActualLengthCm"].ToString();         ActualLengthCm = ActualLengthCm.Trim() == "" ? "0" : ActualLengthCm;
                                string QtyPcs = drMDT["QtyPcs"].ToString();                         QtyPcs = QtyPcs.Trim() == "" ? "0" : QtyPcs;
                                string LengthBodyCm = drMDT["LengthBodyCm"].ToString();             LengthBodyCm = LengthBodyCm.Trim() == "" ? "0" : LengthBodyCm;
                                string LengthBodyM = drMDT["LengthBodyM"].ToString();               LengthBodyM = LengthBodyM.Trim() == "" ? "0" : LengthBodyM;
                                string LengthBodyInc = drMDT["LengthBodyInc"].ToString();           LengthBodyInc = LengthBodyInc.Trim() == "" ? "0" : LengthBodyInc;
                                string LengthBodyYrd = drMDT["LengthBodyYrd"].ToString();           LengthBodyYrd = LengthBodyYrd.Trim() == "" ? "0" : LengthBodyYrd;
                                string WeightMg = drMDT["WeightMg"].ToString();                     WeightMg = WeightMg.Trim() == "" ? "0" : WeightMg;
                                string WeightPcs = drMDT["WeightPcs"].ToString();                   WeightPcs = WeightPcs.Trim() == "" ? "0" : WeightPcs;
                                string OIDSMPLDT = drMDT["OIDSMPLDT"].ToString();

                                string PatternSizeZone = drMDT["PatternSizeZone"].ToString();       PatternSizeZone = PatternSizeZone.Trim() == "" ? "NULL" : "'" + PatternSizeZone.Trim() + "'";

                                sbSQL.Append("INSERT INTO MarkingDetails(OIDMARK, OIDITEM, OIDSIZE, OIDSIZEZONE, OIDSMPLDTStuff, GPartsStuff, DetailsType, Details, TotalWidthSTD, UsableWidth, GM2, PracticalLengthCM, QuantityPCS, LengthPer1CM, LengthPer1M, LengthPer1INCH, LengthPer1YARD, WeightG, WeightKG) ");
                                sbSQL.Append(" SELECT TOP(1) OIDMARK, " + OIDITEM + " AS OIDITEM, " + OIDSIZE + " AS OIDSIZE, " + PatternSizeZone + " AS OIDSIZEZONE, ");
                                sbSQL.Append("          N'" + OIDSMPLDT + "' AS OIDSMPLDTStuff, N'" + FBPartsID + "' AS GPartsStuff, '" + DetailsType + "' AS DetailsType, N'' AS Details, ");
                                sbSQL.Append("          '" + TotalWidth + "' AS TotalWidthSTD, '" + UsableWidth + "' AS UsableWidth, '" + WeightGM2 + "' AS GM2, '" + ActualLengthCm + "' AS PracticalLengthCM, ");
                                sbSQL.Append("          '" + QtyPcs + "' AS QuantityPCS, '" + LengthBodyCm + "' AS LengthPer1CM, '" + LengthBodyM + "' AS LengthPer1M, '" + LengthBodyInc + "' AS LengthPer1INCH, ");
                                sbSQL.Append("          '" + LengthBodyYrd + "' AS LengthPer1YARD, '" + WeightMg + "' AS WeightG, '" + WeightPcs + "' AS WeightKG  ");
                                sbSQL.Append(" FROM Marking ");
                                sbSQL.Append(" WHERE (MarkingNo = N'" + newMarking + "')  ");

                                runRow++;
                            }
                        }
                    }
                    else if (ACTION == "UPDATE")
                    {
                        newMarking = txeMarkingNo.Text.Trim();

                        sbSQL.Append("UPDATE Marking SET ");
                        sbSQL.Append("  DocumentDate = " + DocumentDate + ", ");
                        sbSQL.Append("  MarkingType = " + MarkingRequestType + ", ");
                        sbSQL.Append("  Branch = " + Branch + ", ");
                        sbSQL.Append("  CuttingFactory = " + CuttingFac + ", ");
                        sbSQL.Append("  SewingFactory = " + SewingFac + ", ");
                        sbSQL.Append("  Remark = N'" + mmRemark.Text + "', ");
                        sbSQL.Append("  UpdatedBy = '" + UpdatedBy + "', ");
                        sbSQL.Append("  UpdatedDate = '" + UpdatedDate + "' ");
                        sbSQL.Append("WHERE (OIDMARK = '" + lblMarkID.Text + "')  ");

                        sbSQL.Append("UPDATE SMPLRequest SET Status = '1' WHERE (OIDSMPL = '" + lblID.Text + "')  ");

                        DataTable dtMDT = (DataTable)gcMDT.DataSource;
                        if (dtMDT != null)
                        {
                            string ITEM_SIZE = "";
                            int runRow = 0;
                            foreach (DataRow drMDT in dtMDT.Rows)
                            {
                                string ITEM = drMDT["OIDITEM"].ToString();      ITEM = ITEM.Trim();
                                string SIZE = drMDT["OIDSIZE"].ToString();      SIZE = SIZE.Trim();
                                if (ITEM != "" && SIZE != "")
                                {
                                    if (ITEM_SIZE != "")
                                        ITEM_SIZE += ", ";
                                    ITEM_SIZE += "'" + ITEM + "-" + SIZE + "'";

                                    string ID = drMDT["ID"].ToString();
                                    string OIDITEM = drMDT["OIDITEM"].ToString();                           OIDITEM = OIDITEM.Trim() == "" ? "NULL" : "'" + OIDITEM.Trim() + "'";
                                    string ItemCode = drMDT["ItemCode"].ToString();
                                    string ItemDescription = drMDT["ItemDescription"].ToString();
                                    string SMPLPatternNo = drMDT["SMPLPatternNo"].ToString();
                                    string FBPartsID = drMDT["FBPartsID"].ToString();
                                    string FBParts = drMDT["FBParts"].ToString();
                                    string VendorFBCode = drMDT["VendorFBCode"].ToString();
                                    string SampleLotNo = drMDT["SampleLotNo"].ToString();
                                    string DetailsType = drMDT["DetailsType"].ToString();
                                    string TotalWidth = drMDT["TotalWidth"].ToString();                     TotalWidth = TotalWidth.Trim() == "" ? "0" : TotalWidth;
                                    string UsableWidth = drMDT["UsableWidth"].ToString();                   UsableWidth = UsableWidth.Trim() == "" ? "0" : UsableWidth;
                                    string WeightGM2 = drMDT["WeightGM2"].ToString();                       WeightGM2 = WeightGM2.Trim() == "" ? "0" : WeightGM2;
                                    string OIDSIZE = drMDT["OIDSIZE"].ToString();                           OIDSIZE = OIDSIZE.Trim() == "" ? "NULL" : "'" + OIDSIZE.Trim() + "'";
                                    string SizeName = drMDT["SizeName"].ToString();
                                    string ActualLengthCm = drMDT["ActualLengthCm"].ToString();             ActualLengthCm = ActualLengthCm.Trim() == "" ? "0" : ActualLengthCm;
                                    string QtyPcs = drMDT["QtyPcs"].ToString();                             QtyPcs = QtyPcs.Trim() == "" ? "0" : QtyPcs;
                                    string LengthBodyCm = drMDT["LengthBodyCm"].ToString();                 LengthBodyCm = LengthBodyCm.Trim() == "" ? "0" : LengthBodyCm;
                                    string LengthBodyM = drMDT["LengthBodyM"].ToString();                   LengthBodyM = LengthBodyM.Trim() == "" ? "0" : LengthBodyM;
                                    string LengthBodyInc = drMDT["LengthBodyInc"].ToString();               LengthBodyInc = LengthBodyInc.Trim() == "" ? "0" : LengthBodyInc;
                                    string LengthBodyYrd = drMDT["LengthBodyYrd"].ToString();               LengthBodyYrd = LengthBodyYrd.Trim() == "" ? "0" : LengthBodyYrd;
                                    string WeightMg = drMDT["WeightMg"].ToString();                         WeightMg = WeightMg.Trim() == "" ? "0" : WeightMg;
                                    string WeightPcs = drMDT["WeightPcs"].ToString();                       WeightPcs = WeightPcs.Trim() == "" ? "0" : WeightPcs;
                                    string OIDSMPLDT = drMDT["OIDSMPLDT"].ToString();

                                    string PatternSizeZone = drMDT["PatternSizeZone"].ToString();           PatternSizeZone = PatternSizeZone.Trim() == "" ? "NULL" : "'" + PatternSizeZone.Trim() + "'";

                                    sbSQL.Append("IF NOT EXISTS(SELECT OIDMARKDT FROM MarkingDetails WHERE OIDMARK = '" + lblMarkID.Text + "' AND OIDITEM = '" + ITEM + "' AND OIDSIZE = '" + SIZE + "' AND DetailsType = '" + DetailsType + "') ");
                                    sbSQL.Append(" BEGIN ");
                                    sbSQL.Append("  INSERT INTO MarkingDetails(OIDMARK, OIDITEM, OIDSIZE, OIDSIZEZONE, OIDSMPLDTStuff, GPartsStuff, DetailsType, Details, TotalWidthSTD, UsableWidth, GM2, PracticalLengthCM, QuantityPCS, LengthPer1CM, LengthPer1M, LengthPer1INCH, LengthPer1YARD, WeightG, WeightKG) ");
                                    sbSQL.Append("  VALUES('" + lblMarkID.Text + "', " + OIDITEM + ", " + OIDSIZE + ", " + PatternSizeZone + ", N'" + OIDSMPLDT + "', N'" + FBPartsID + "', '" + DetailsType + "', N'', '" + TotalWidth + "', '" + UsableWidth + "', '" + WeightGM2 + "', '" + ActualLengthCm + "', '" + QtyPcs + "', '" + LengthBodyCm + "', '" + LengthBodyM + "', '" + LengthBodyInc + "', '" + LengthBodyYrd + "', '" + WeightMg + "', '" + WeightPcs + "')  ");
                                    sbSQL.Append(" END ");
                                    sbSQL.Append("ELSE ");
                                    sbSQL.Append(" BEGIN ");
                                    sbSQL.Append("  UPDATE MarkingDetails SET  ");
                                    sbSQL.Append("    OIDSIZEZONE=" + PatternSizeZone + ", OIDSMPLDTStuff=N'" + OIDSMPLDT + "', GPartsStuff=N'" + FBPartsID + "', ");
                                    sbSQL.Append("    DetailsType='" + DetailsType + "', Details=N'', TotalWidthSTD='" + TotalWidth + "', ");
                                    sbSQL.Append("    UsableWidth='" + UsableWidth + "', GM2='" + WeightGM2 + "', PracticalLengthCM='" + ActualLengthCm + "', ");
                                    sbSQL.Append("    QuantityPCS='" + QtyPcs + "', LengthPer1CM='" + LengthBodyCm + "', LengthPer1M='" + LengthBodyM + "', ");
                                    sbSQL.Append("    LengthPer1INCH='" + LengthBodyInc + "', LengthPer1YARD='" + LengthBodyYrd + "', WeightG='" + WeightMg + "', WeightKG='" + WeightPcs + "'  ");
                                    sbSQL.Append("  WHERE (OIDMARK = '" + lblMarkID.Text + "') AND (OIDITEM = '" + ITEM + "') AND (OIDSIZE = '" + SIZE + "') AND (DetailsType = '" + DetailsType + "') ");
                                    sbSQL.Append(" END ");
                                    //MessageBox.Show(sbSQL.ToString());
                                }
                                runRow++;
                            }

                            if (ITEM_SIZE != "")
                            {
                                sbSQL.Append("DELETE FROM MarkingDetails ");
                                sbSQL.Append("WHERE (OIDMARK = '" + lblMarkID.Text + "') AND ((CONVERT(VARCHAR, OIDITEM) + '-' + CONVERT(VARCHAR, OIDSIZE)) NOT IN (" + ITEM_SIZE + "))  ");
                            }

                        }
                    }

                    if (sbSQL.Length > 0)
                    {
                        bool chkSave = this.DBC.DBQuery(sbSQL).runSQL();

                        if (chkSave == false)
                        {
                            //textBox1.Text = sbSQL.ToString();
                            FUNC.msgERROR("Found problem on save.\nพบปัญหาในการบันทึกข้อมูล");
                        }
                        else
                        {
                            if (FUNC.msgQuiz(msgACTION + " marking complete. Do you want to load this marking ?") == true)
                            {
                                //Load marking after save.
                                tabMARKING.SelectedTabPage = lcgMark;
                                LoadMarkingDocument(newMarking, "UPDATE");
                            }
                            else
                            {
                                //Clear All Document
                                tabMARKING.SelectedTabPage = lcgList;
                                LoadNewData();
                            }
                        }
                    }
                }
            }
        }

        private void bbiExcel_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            //string pathFile = new ObjSet.Folder(@"C:\MDS\Export\").GetPath() + "PaymentTermList_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
            //gvPTerm.ExportToXlsx(pathFile);
            //System.Diagnostics.Process.Start(pathFile);
        }

        private void bbiPrintPreview_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            //gcPTerm.ShowPrintPreview();
        }

        private void bbiPrint_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            int[] selectedRowHandles = gvMARK.GetSelectedRows();
            if (selectedRowHandles.Length > 0)
            {
                gvMARK.FocusedRowHandle = selectedRowHandles[0];
                string MARKID = gvMARK.GetRowCellDisplayText(selectedRowHandles[0], "ID");
                string MARKNo = gvMARK.GetRowCellDisplayText(selectedRowHandles[0], "Marking No.");
                if (FUNC.msgQuiz("Confirm print marking (excel file)  : " + MARKNo + " ?") == true)
                {
                    layoutControlItem120.Text = "Print excel file processing ..";
                    layoutControlItem120.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                    pbcEXPORT.Properties.Step = 1;
                    pbcEXPORT.Properties.PercentView = true;
                    pbcEXPORT.Properties.Maximum = 11;
                    pbcEXPORT.Properties.Minimum = 0;

                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT    TOP (1) SMPL.SMPLItem, SMPL.ModelName, SMPL.SMPLPatternNo, US.UserName, US.FullName, CUS.Name AS Customer, SMPL.Season ");
                    sbSQL.Append("FROM      Marking AS MK INNER JOIN ");
                    sbSQL.Append("          SMPLRequest AS SMPL ON MK.OIDSMPL = SMPL.OIDSMPL AND MK.OIDMARK = '" + MARKID + "' LEFT OUTER JOIN ");
                    sbSQL.Append("          Users AS US ON MK.CreatedBy = US.OIDUSER LEFT OUTER JOIN ");
                    sbSQL.Append("          Customer AS CUS ON SMPL.OIDCUST = CUS.OIDCUST ");

                    string[] MARK = this.DBC.DBQuery(sbSQL).getMultipleValue();
                    if (MARK.Length > 0)
                    {
                        //****** BEGIN EXPORT *******

                        String sFilePath = System.IO.Path.Combine(new ObjSet.Folder(@"C:\MDS\Export\").GetPath() + MARKNo + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");
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
                            objWorkBook = objApp.Workbooks.Open(this.reportPath + "MARK.xlsx");

                            int LastRow = 9;
                            //** Standard ***
                            LastRow = 9;
                            objSheet = (Microsoft.Office.Interop.Excel.Worksheet)objWorkBook.Sheets[1];

                            objSheet.Cells[2, 1] = MARK[5].ToUpper().Trim(); 
                            objSheet.Cells[4, 4] = MARK[0];
                            objSheet.Cells[4, 9] = MARK[1];
                            objSheet.Cells[4, 18] = MARK[2];
                            objSheet.Cells[6, 14] = MARK[6];
                            objSheet.Cells[5, 21] = MARK[3].ToUpper().Trim();

                            pbcEXPORT.PerformStep();
                            pbcEXPORT.Update();

                            sbSQL.Clear();
                            sbSQL.Append("SELECT PS.SizeName AS Size, ");
                            sbSQL.Append("       (SELECT GarmentParts + ', ' AS 'data()' FROM(SELECT DISTINCT GP.OIDGParts, GP.GarmentParts FROM SMPLQuantityRequired AS SQR INNER JOIN SMPLRequestFabric AS FB ON SQR.OIDSMPLDT = FB.OIDSMPLDT AND SQR.OIDSMPL = MK.OIDSMPL INNER JOIN SMPLRequestFabricParts AS SFBP ON FB.OIDSMPLFB = SFBP.OIDSMPLFB AND FB.OIDSMPLDT = SFBP.OIDSMPLDT AND FB.OIDITEM = MKD.OIDITEM INNER JOIN GarmentParts AS GP ON SFBP.OIDGParts = GP.OIDGParts AND GP.GarmentParts IS NOT NULL) AS FBParts FOR XML PATH('')) AS FBParts, ");
                            sbSQL.Append("       (SELECT VendFBCode + ', ' AS 'data()' FROM(SELECT DISTINCT VendFBCode FROM SMPLRequestFabric WHERE OIDSMPLDT IN(SELECT OIDSMPLDT FROM SMPLQuantityRequired  WHERE OIDSMPL = MK.OIDSMPL) AND OIDITEM = MKD.OIDITEM AND OIDSMPLDT IN(SELECT value FROM MarkingDetails CROSS APPLY string_split(OIDSMPLDTStuff, ',')  WHERE OIDMARKDT = MKD.OIDMARKDT) AND VendFBCode IS NOT NULL) AS VendFBCode FOR XML PATH('')) AS VendorFBCode, MKD.TotalWidthSTD, MKD.UsableWidth, MKD.GM2, ");
                            sbSQL.Append("       MKD.PracticalLengthCM, MKD.QuantityPCS, MKD.LengthPer1CM, MKD.LengthPer1M, MKD.LengthPer1INCH, MKD.LengthPer1YARD, MKD.WeightG, MKD.WeightKG ");
                            sbSQL.Append("FROM   MarkingDetails AS MKD INNER JOIN ");
                            sbSQL.Append("       Marking AS MK ON MKD.OIDMARK = MK.OIDMARK INNER JOIN ");
                            sbSQL.Append("       ProductSize AS PS ON MKD.OIDSIZE = PS.OIDSIZE ");
                            sbSQL.Append("WHERE (MKD.OIDMARK = '" + MARKID + "') AND(MKD.DetailsType = 0) ");
                            sbSQL.Append("ORDER BY MKD.OIDITEM, MKD.OIDSIZE ");
                            DataTable dtSTD = this.DBC.DBQuery(sbSQL).getDataTable();
                            if (dtSTD != null)
                            {
                                int totalRow = dtSTD.Rows.Count;
                                int diffRow = totalRow > 8 ? totalRow - 8 : 0;

                                if (diffRow > 0)
                                {
                                    //ลบแถว
                                    for (int i = 0; i < diffRow; i++)
                                    {
                                        objSheet.Rows[17].Delete();
                                    }

                                    //แทรกแถว + Merge cell
                                    for (int i = 0; i < diffRow; i++)
                                    {
                                        objSheet.Rows[16].Insert();
                                        objSheet.Range[objSheet.Cells[16, 3], objSheet.Cells[16, 4]].Merge();
                                        objSheet.Range[objSheet.Cells[16, 5], objSheet.Cells[16, 7]].Merge();
                                        objSheet.Range[objSheet.Cells[16, 8], objSheet.Cells[16, 9]].Merge();
                                        objSheet.Range[objSheet.Cells[16, 11], objSheet.Cells[16, 12]].Merge();
                                        objSheet.Range[objSheet.Cells[16, 13], objSheet.Cells[16, 15]].Merge();
                                        objSheet.Range[objSheet.Cells[16, 17], objSheet.Cells[16, 18]].Merge();
                                    }
                                }

                                string chkSize = "";
                                string chkFBParts = "";
                                string chkVendorFBCode = "";
                                string chkTotalWidthSTD = "";
                                string chkUsableWidth = "";
                                string chkGM2 = "";

                                int runRow = 0;
                                foreach (DataRow drMARK in dtSTD.Rows)
                                {
                                    string Size = drMARK["Size"].ToString().ToUpper().Trim();
                                    if (chkSize == Size)
                                        objSheet.Range[objSheet.Cells[LastRow - 1, 2], objSheet.Cells[LastRow, 2]].Merge();
                                    objSheet.Cells[LastRow, 2] = Size;

                                    string FBParts = drMARK["FBParts"].ToString().Trim();
                                    FBParts = FBParts.IndexOf(',') > -1 ? FBParts.ToUpper().Trim().Substring(0, FBParts.ToUpper().Trim().Length - 1) : FBParts.ToUpper().Trim();
                                    if (chkFBParts == FBParts)
                                        objSheet.Range[objSheet.Cells[LastRow - 1, 3], objSheet.Cells[LastRow, 3]].Merge();
                                    objSheet.Cells[LastRow, 3] = FBParts;

                                    string VendorFBCode = drMARK["VendorFBCode"].ToString().Trim();
                                    VendorFBCode = VendorFBCode.IndexOf(',') > -1 ? VendorFBCode.ToUpper().Trim().Substring(0, VendorFBCode.ToUpper().Trim().Length - 1) : VendorFBCode.ToUpper().Trim();
                                    if (chkVendorFBCode == VendorFBCode)
                                        objSheet.Range[objSheet.Cells[LastRow - 1, 5], objSheet.Cells[LastRow, 5]].Merge();
                                    objSheet.Cells[LastRow, 5] = VendorFBCode;

                                    string TotalWidthSTD = drMARK["TotalWidthSTD"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["TotalWidthSTD"].ToString()).ToString("###0.####");
                                    if (chkTotalWidthSTD == TotalWidthSTD)
                                        objSheet.Range[objSheet.Cells[LastRow - 1, 8], objSheet.Cells[LastRow, 8]].Merge();
                                    objSheet.Cells[LastRow, 8] = TotalWidthSTD;

                                    string UsableWidth = drMARK["UsableWidth"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["UsableWidth"].ToString()).ToString("###0.####");
                                    if (chkUsableWidth == UsableWidth)
                                        objSheet.Range[objSheet.Cells[LastRow - 1, 10], objSheet.Cells[LastRow, 10]].Merge();
                                    objSheet.Cells[LastRow, 10] = UsableWidth;

                                    string GM2 = drMARK["GM2"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["GM2"].ToString()).ToString("###0.####");
                                    if (chkGM2 == GM2)
                                        objSheet.Range[objSheet.Cells[LastRow - 1, 11], objSheet.Cells[LastRow, 11]].Merge();
                                    objSheet.Cells[LastRow, 11] = GM2;

                                    string PracticalLengthCM = drMARK["PracticalLengthCM"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["PracticalLengthCM"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 13] = PracticalLengthCM;

                                    string QuantityPCS = drMARK["QuantityPCS"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["QuantityPCS"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 16] = QuantityPCS;

                                    string LengthPer1CM = drMARK["LengthPer1CM"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["LengthPer1CM"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 17] = LengthPer1CM;

                                    string LengthPer1M = drMARK["LengthPer1M"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["LengthPer1M"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 19] = drMARK["LengthPer1M"].ToString().ToUpper().Trim();

                                    string LengthPer1INCH = drMARK["LengthPer1INCH"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["LengthPer1INCH"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 20] = LengthPer1INCH;

                                    string LengthPer1YARD = drMARK["LengthPer1YARD"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["LengthPer1YARD"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 21] = LengthPer1YARD;

                                    string WeightG = drMARK["WeightG"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["WeightG"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 22] = WeightG;

                                    string WeightKG = drMARK["WeightKG"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["WeightKG"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 23] = WeightKG;

                                    if (chkSize != Size || chkFBParts != FBParts || chkVendorFBCode != VendorFBCode || chkTotalWidthSTD != TotalWidthSTD || chkUsableWidth != UsableWidth || chkGM2 != GM2)
                                    {
                                        chkSize = Size;
                                        chkFBParts = FBParts;
                                        chkVendorFBCode = VendorFBCode;
                                        chkTotalWidthSTD = TotalWidthSTD;
                                        chkUsableWidth = UsableWidth;
                                        chkGM2 = GM2;
                                    }

                                    LastRow++; 
                                    runRow++;
                                }
                            }

                            pbcEXPORT.PerformStep();
                            pbcEXPORT.Update();

                            //** Positive ***
                            LastRow = 9;
                            objSheet = (Microsoft.Office.Interop.Excel.Worksheet)objWorkBook.Sheets[2];

                            objSheet.Cells[2, 1] = MARK[5].ToUpper().Trim();
                            objSheet.Cells[4, 4] = MARK[0];
                            objSheet.Cells[4, 9] = MARK[1];
                            objSheet.Cells[4, 18] = MARK[2];
                            objSheet.Cells[6, 14] = MARK[6];
                            objSheet.Cells[5, 21] = MARK[3].ToUpper().Trim();

                            pbcEXPORT.PerformStep();
                            pbcEXPORT.Update();

                            sbSQL.Clear();
                            sbSQL.Append("SELECT PS.SizeName AS Size, ");
                            sbSQL.Append("       (SELECT GarmentParts + ', ' AS 'data()' FROM(SELECT DISTINCT GP.OIDGParts, GP.GarmentParts FROM SMPLQuantityRequired AS SQR INNER JOIN SMPLRequestFabric AS FB ON SQR.OIDSMPLDT = FB.OIDSMPLDT AND SQR.OIDSMPL = MK.OIDSMPL INNER JOIN SMPLRequestFabricParts AS SFBP ON FB.OIDSMPLFB = SFBP.OIDSMPLFB AND FB.OIDSMPLDT = SFBP.OIDSMPLDT AND FB.OIDITEM = MKD.OIDITEM INNER JOIN GarmentParts AS GP ON SFBP.OIDGParts = GP.OIDGParts AND GP.GarmentParts IS NOT NULL) AS FBParts FOR XML PATH('')) AS FBParts, ");
                            sbSQL.Append("       (SELECT VendFBCode + ', ' AS 'data()' FROM(SELECT DISTINCT VendFBCode FROM SMPLRequestFabric WHERE OIDSMPLDT IN(SELECT OIDSMPLDT FROM SMPLQuantityRequired  WHERE OIDSMPL = MK.OIDSMPL) AND OIDITEM = MKD.OIDITEM AND OIDSMPLDT IN(SELECT value FROM MarkingDetails CROSS APPLY string_split(OIDSMPLDTStuff, ',')  WHERE OIDMARKDT = MKD.OIDMARKDT) AND VendFBCode IS NOT NULL) AS VendFBCode FOR XML PATH('')) AS VendorFBCode, MKD.TotalWidthSTD, MKD.UsableWidth, MKD.GM2, ");
                            sbSQL.Append("       MKD.PracticalLengthCM, MKD.QuantityPCS, MKD.LengthPer1CM, MKD.LengthPer1M, MKD.LengthPer1INCH, MKD.LengthPer1YARD, MKD.WeightG, MKD.WeightKG ");
                            sbSQL.Append("FROM   MarkingDetails AS MKD INNER JOIN ");
                            sbSQL.Append("       Marking AS MK ON MKD.OIDMARK = MK.OIDMARK INNER JOIN ");
                            sbSQL.Append("       ProductSize AS PS ON MKD.OIDSIZE = PS.OIDSIZE ");
                            sbSQL.Append("WHERE (MKD.OIDMARK = '" + MARKID + "') AND(MKD.DetailsType = 1) ");
                            sbSQL.Append("ORDER BY MKD.OIDITEM, MKD.OIDSIZE ");
                            DataTable dtPOS = this.DBC.DBQuery(sbSQL).getDataTable();
                            if (dtPOS != null)
                            {
                                int totalRow = dtPOS.Rows.Count;
                                int diffRow = totalRow > 8 ? totalRow - 8 : 0;

                                if (diffRow > 0)
                                {
                                    //ลบแถว
                                    for (int i = 0; i < diffRow; i++)
                                    {
                                        objSheet.Rows[17].Delete();
                                    }

                                    //แทรกแถว + Merge cell
                                    for (int i = 0; i < diffRow; i++)
                                    {
                                        objSheet.Rows[16].Insert();
                                        objSheet.Range[objSheet.Cells[16, 3], objSheet.Cells[16, 4]].Merge();
                                        objSheet.Range[objSheet.Cells[16, 5], objSheet.Cells[16, 7]].Merge();
                                        objSheet.Range[objSheet.Cells[16, 8], objSheet.Cells[16, 9]].Merge();
                                        objSheet.Range[objSheet.Cells[16, 11], objSheet.Cells[16, 12]].Merge();
                                        objSheet.Range[objSheet.Cells[16, 13], objSheet.Cells[16, 15]].Merge();
                                        objSheet.Range[objSheet.Cells[16, 17], objSheet.Cells[16, 18]].Merge();
                                    }
                                }

                                string chkSize = "";
                                string chkFBParts = "";
                                string chkVendorFBCode = "";
                                string chkTotalWidthSTD = "";
                                string chkUsableWidth = "";
                                string chkGM2 = "";

                                int runRow = 0;
                                foreach (DataRow drMARK in dtPOS.Rows)
                                {
                                    string Size = drMARK["Size"].ToString().ToUpper().Trim();
                                    if (chkSize == Size)
                                        objSheet.Range[objSheet.Cells[LastRow - 1, 2], objSheet.Cells[LastRow, 2]].Merge();
                                    objSheet.Cells[LastRow, 2] = Size;

                                    string FBParts = drMARK["FBParts"].ToString().Trim();
                                    FBParts = FBParts.IndexOf(',') > -1 ? FBParts.ToUpper().Trim().Substring(0, FBParts.ToUpper().Trim().Length - 1) : FBParts.ToUpper().Trim();
                                    if (chkFBParts == FBParts)
                                        objSheet.Range[objSheet.Cells[LastRow - 1, 3], objSheet.Cells[LastRow, 3]].Merge();
                                    objSheet.Cells[LastRow, 3] = FBParts;

                                    string VendorFBCode = drMARK["VendorFBCode"].ToString().Trim();
                                    VendorFBCode = VendorFBCode.IndexOf(',') > -1 ? VendorFBCode.ToUpper().Trim().Substring(0, VendorFBCode.ToUpper().Trim().Length - 1) : VendorFBCode.ToUpper().Trim();
                                    if (chkVendorFBCode == VendorFBCode)
                                        objSheet.Range[objSheet.Cells[LastRow - 1, 5], objSheet.Cells[LastRow, 5]].Merge();
                                    objSheet.Cells[LastRow, 5] = VendorFBCode;

                                    string TotalWidthSTD = drMARK["TotalWidthSTD"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["TotalWidthSTD"].ToString()).ToString("###0.####");
                                    if (chkTotalWidthSTD == TotalWidthSTD)
                                        objSheet.Range[objSheet.Cells[LastRow - 1, 8], objSheet.Cells[LastRow, 8]].Merge();
                                    objSheet.Cells[LastRow, 8] = TotalWidthSTD;

                                    string UsableWidth = drMARK["UsableWidth"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["UsableWidth"].ToString()).ToString("###0.####");
                                    if (chkUsableWidth == UsableWidth)
                                        objSheet.Range[objSheet.Cells[LastRow - 1, 10], objSheet.Cells[LastRow, 10]].Merge();
                                    objSheet.Cells[LastRow, 10] = UsableWidth;

                                    string GM2 = drMARK["GM2"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["GM2"].ToString()).ToString("###0.####");
                                    if (chkGM2 == GM2)
                                        objSheet.Range[objSheet.Cells[LastRow - 1, 11], objSheet.Cells[LastRow, 11]].Merge();
                                    objSheet.Cells[LastRow, 11] = GM2;

                                    string PracticalLengthCM = drMARK["PracticalLengthCM"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["PracticalLengthCM"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 13] = PracticalLengthCM;

                                    string QuantityPCS = drMARK["QuantityPCS"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["QuantityPCS"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 16] = QuantityPCS;

                                    string LengthPer1CM = drMARK["LengthPer1CM"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["LengthPer1CM"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 17] = LengthPer1CM;

                                    string LengthPer1M = drMARK["LengthPer1M"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["LengthPer1M"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 19] = drMARK["LengthPer1M"].ToString().ToUpper().Trim();

                                    string LengthPer1INCH = drMARK["LengthPer1INCH"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["LengthPer1INCH"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 20] = LengthPer1INCH;

                                    string LengthPer1YARD = drMARK["LengthPer1YARD"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["LengthPer1YARD"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 21] = LengthPer1YARD;

                                    string WeightG = drMARK["WeightG"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["WeightG"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 22] = WeightG;

                                    string WeightKG = drMARK["WeightKG"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["WeightKG"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 23] = WeightKG;

                                    if (chkSize != Size || chkFBParts != FBParts || chkVendorFBCode != VendorFBCode || chkTotalWidthSTD != TotalWidthSTD || chkUsableWidth != UsableWidth || chkGM2 != GM2)
                                    {
                                        chkSize = Size;
                                        chkFBParts = FBParts;
                                        chkVendorFBCode = VendorFBCode;
                                        chkTotalWidthSTD = TotalWidthSTD;
                                        chkUsableWidth = UsableWidth;
                                        chkGM2 = GM2;
                                    }

                                    LastRow++; 
                                    runRow++;
                                }
                            }

                            pbcEXPORT.PerformStep();
                            pbcEXPORT.Update();

                            //** Nagative ***
                            LastRow = 9;
                            objSheet = (Microsoft.Office.Interop.Excel.Worksheet)objWorkBook.Sheets[3];

                            objSheet.Cells[2, 1] = MARK[5].ToUpper().Trim();
                            objSheet.Cells[4, 4] = MARK[0];
                            objSheet.Cells[4, 9] = MARK[1];
                            objSheet.Cells[4, 18] = MARK[2];
                            objSheet.Cells[6, 14] = MARK[6];
                            objSheet.Cells[5, 21] = MARK[3].ToUpper().Trim();

                            pbcEXPORT.PerformStep();
                            pbcEXPORT.Update();

                            sbSQL.Clear();
                            sbSQL.Append("SELECT PS.SizeName AS Size, ");
                            sbSQL.Append("       (SELECT GarmentParts + ', ' AS 'data()' FROM(SELECT DISTINCT GP.OIDGParts, GP.GarmentParts FROM SMPLQuantityRequired AS SQR INNER JOIN SMPLRequestFabric AS FB ON SQR.OIDSMPLDT = FB.OIDSMPLDT AND SQR.OIDSMPL = MK.OIDSMPL INNER JOIN SMPLRequestFabricParts AS SFBP ON FB.OIDSMPLFB = SFBP.OIDSMPLFB AND FB.OIDSMPLDT = SFBP.OIDSMPLDT AND FB.OIDITEM = MKD.OIDITEM INNER JOIN GarmentParts AS GP ON SFBP.OIDGParts = GP.OIDGParts AND GP.GarmentParts IS NOT NULL) AS FBParts FOR XML PATH('')) AS FBParts, ");
                            sbSQL.Append("       (SELECT VendFBCode + ', ' AS 'data()' FROM(SELECT DISTINCT VendFBCode FROM SMPLRequestFabric WHERE OIDSMPLDT IN(SELECT OIDSMPLDT FROM SMPLQuantityRequired  WHERE OIDSMPL = MK.OIDSMPL) AND OIDITEM = MKD.OIDITEM AND OIDSMPLDT IN(SELECT value FROM MarkingDetails CROSS APPLY string_split(OIDSMPLDTStuff, ',')  WHERE OIDMARKDT = MKD.OIDMARKDT) AND VendFBCode IS NOT NULL) AS VendFBCode FOR XML PATH('')) AS VendorFBCode, MKD.TotalWidthSTD, MKD.UsableWidth, MKD.GM2, ");
                            sbSQL.Append("       MKD.PracticalLengthCM, MKD.QuantityPCS, MKD.LengthPer1CM, MKD.LengthPer1M, MKD.LengthPer1INCH, MKD.LengthPer1YARD, MKD.WeightG, MKD.WeightKG ");
                            sbSQL.Append("FROM   MarkingDetails AS MKD INNER JOIN ");
                            sbSQL.Append("       Marking AS MK ON MKD.OIDMARK = MK.OIDMARK INNER JOIN ");
                            sbSQL.Append("       ProductSize AS PS ON MKD.OIDSIZE = PS.OIDSIZE ");
                            sbSQL.Append("WHERE (MKD.OIDMARK = '" + MARKID + "') AND(MKD.DetailsType = 2) ");
                            sbSQL.Append("ORDER BY MKD.OIDITEM, MKD.OIDSIZE ");
                            DataTable dtNEG = this.DBC.DBQuery(sbSQL).getDataTable();
                            if (dtNEG != null)
                            {
                                int totalRow = dtNEG.Rows.Count;
                                int diffRow = totalRow > 8 ? totalRow - 8 : 0;

                                if (diffRow > 0)
                                {
                                    //ลบแถว
                                    for (int i = 0; i < diffRow; i++)
                                    {
                                        objSheet.Rows[17].Delete();
                                    }

                                    //แทรกแถว + Merge cell
                                    for (int i = 0; i < diffRow; i++)
                                    {
                                        objSheet.Rows[16].Insert();
                                        objSheet.Range[objSheet.Cells[16, 3], objSheet.Cells[16, 4]].Merge();
                                        objSheet.Range[objSheet.Cells[16, 5], objSheet.Cells[16, 7]].Merge();
                                        objSheet.Range[objSheet.Cells[16, 8], objSheet.Cells[16, 9]].Merge();
                                        objSheet.Range[objSheet.Cells[16, 11], objSheet.Cells[16, 12]].Merge();
                                        objSheet.Range[objSheet.Cells[16, 13], objSheet.Cells[16, 15]].Merge();
                                        objSheet.Range[objSheet.Cells[16, 17], objSheet.Cells[16, 18]].Merge();
                                    }
                                }

                                string chkSize = "";
                                string chkFBParts = "";
                                string chkVendorFBCode = "";
                                string chkTotalWidthSTD = "";
                                string chkUsableWidth = "";
                                string chkGM2 = "";

                                int runRow = 0;
                                foreach (DataRow drMARK in dtNEG.Rows)
                                {
                                    string Size = drMARK["Size"].ToString().ToUpper().Trim();
                                    if (chkSize == Size)
                                        objSheet.Range[objSheet.Cells[LastRow - 1, 2], objSheet.Cells[LastRow, 2]].Merge();
                                    objSheet.Cells[LastRow, 2] = Size;

                                    string FBParts = drMARK["FBParts"].ToString().Trim();
                                    FBParts = FBParts.IndexOf(',') > -1 ? FBParts.ToUpper().Trim().Substring(0, FBParts.ToUpper().Trim().Length - 1) : FBParts.ToUpper().Trim();
                                    if (chkFBParts == FBParts)
                                        objSheet.Range[objSheet.Cells[LastRow - 1, 3], objSheet.Cells[LastRow, 3]].Merge();
                                    objSheet.Cells[LastRow, 3] = FBParts;

                                    string VendorFBCode = drMARK["VendorFBCode"].ToString().Trim();
                                    VendorFBCode = VendorFBCode.IndexOf(',') > -1 ? VendorFBCode.ToUpper().Trim().Substring(0, VendorFBCode.ToUpper().Trim().Length - 1) : VendorFBCode.ToUpper().Trim();
                                    if (chkVendorFBCode == VendorFBCode)
                                        objSheet.Range[objSheet.Cells[LastRow - 1, 5], objSheet.Cells[LastRow, 5]].Merge();
                                    objSheet.Cells[LastRow, 5] = VendorFBCode;

                                    string TotalWidthSTD = drMARK["TotalWidthSTD"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["TotalWidthSTD"].ToString()).ToString("###0.####");
                                    if (chkTotalWidthSTD == TotalWidthSTD)
                                        objSheet.Range[objSheet.Cells[LastRow - 1, 8], objSheet.Cells[LastRow, 8]].Merge();
                                    objSheet.Cells[LastRow, 8] = TotalWidthSTD;

                                    string UsableWidth = drMARK["UsableWidth"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["UsableWidth"].ToString()).ToString("###0.####");
                                    if (chkUsableWidth == UsableWidth)
                                        objSheet.Range[objSheet.Cells[LastRow - 1, 10], objSheet.Cells[LastRow, 10]].Merge();
                                    objSheet.Cells[LastRow, 10] = UsableWidth;

                                    string GM2 = drMARK["GM2"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["GM2"].ToString()).ToString("###0.####");
                                    if (chkGM2 == GM2)
                                        objSheet.Range[objSheet.Cells[LastRow - 1, 11], objSheet.Cells[LastRow, 11]].Merge();
                                    objSheet.Cells[LastRow, 11] = GM2;

                                    string PracticalLengthCM = drMARK["PracticalLengthCM"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["PracticalLengthCM"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 13] = PracticalLengthCM;

                                    string QuantityPCS = drMARK["QuantityPCS"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["QuantityPCS"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 16] = QuantityPCS;

                                    string LengthPer1CM = drMARK["LengthPer1CM"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["LengthPer1CM"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 17] = LengthPer1CM;

                                    string LengthPer1M = drMARK["LengthPer1M"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["LengthPer1M"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 19] = drMARK["LengthPer1M"].ToString().ToUpper().Trim();

                                    string LengthPer1INCH = drMARK["LengthPer1INCH"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["LengthPer1INCH"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 20] = LengthPer1INCH;

                                    string LengthPer1YARD = drMARK["LengthPer1YARD"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["LengthPer1YARD"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 21] = LengthPer1YARD;

                                    string WeightG = drMARK["WeightG"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["WeightG"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 22] = WeightG;

                                    string WeightKG = drMARK["WeightKG"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["WeightKG"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 23] = WeightKG;

                                    if (chkSize != Size || chkFBParts != FBParts || chkVendorFBCode != VendorFBCode || chkTotalWidthSTD != TotalWidthSTD || chkUsableWidth != UsableWidth || chkGM2 != GM2)
                                    {
                                        chkSize = Size;
                                        chkFBParts = FBParts;
                                        chkVendorFBCode = VendorFBCode;
                                        chkTotalWidthSTD = TotalWidthSTD;
                                        chkUsableWidth = UsableWidth;
                                        chkGM2 = GM2;
                                    }

                                    LastRow++;
                                    runRow++;
                                }
                            }

                            pbcEXPORT.PerformStep();
                            pbcEXPORT.Update();


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
                        FUNC.msgError("ไม่พบข้อมูลเอกสาร Marking: " + MARKNo);
                    }
                    layoutControlItem120.Text = "Status ..";
                    layoutControlItem120.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                }

            }
        }


        private void slueRequestNo_EditValueChanged(object sender, EventArgs e)
        {
           
        }

        private void tabMARKING_SelectedPageChanged(object sender, DevExpress.XtraLayout.LayoutTabPageChangedEventArgs e)
        {
            if (tabMARKING.SelectedTabPage == lcgList) //LIST
            {
                bbiNew.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
                bbiEdit.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
                bbiSave.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
                bbiDelete.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
                bbiRefresh.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;

                ribbonPageGroup2.Visible = true;

                if (chkReadWrite == 1)
                    rpgManage.Visible = true;

                if (gvMARK.SelectedRowsCount != 0)
                {
                    GridView gv = gvMARK;
                    int[] selectedRowHandles = gvMARK.GetSelectedRows();
                    if (selectedRowHandles.Length > 0)
                    {
                        bbiPrint.Enabled = true;
                        bbiPrintPDF.Enabled = true;
                        //bbiCLONE.Enabled = true;

                        string OIDUSER = gv.GetFocusedRowCellValue("ByCreated").ToString();
                        if (UserLogin.OIDUser.ToString() == OIDUSER)
                        {
                            bbiUPDATE.Enabled = true;
                            //bbiREVISE.Enabled = true;
                            //bbiDELBILL.Enabled = true;

                            string MARKStatus = gv.GetFocusedRowCellValue("Status").ToString();
                            if (MARKStatus == "0")
                                bbiDELBILL.Enabled = false;
                            else
                                bbiDELBILL.Enabled = true;
                        }
                        else
                        {
                            bbiUPDATE.Enabled = false;
                            //bbiREVISE.Enabled = false;
                            bbiDELBILL.Enabled = false;
                        }
                    }
                    else
                    {
                        bbiPrint.Enabled = false;
                        bbiPrintPDF.Enabled = false;
                        bbiUPDATE.Enabled = false;
                        //bbiREVISE.Enabled = false;
                        //bbiCLONE.Enabled = false;
                        bbiDELBILL.Enabled = false;
                    }
                }
                else
                {
                    bbiPrint.Enabled = false;
                    bbiPrintPDF.Enabled = false;
                    bbiUPDATE.Enabled = false;
                    //bbiREVISE.Enabled = false;
                    //bbiCLONE.Enabled = false;
                    bbiDELBILL.Enabled = false;
                }
            }
            else if (tabMARKING.SelectedTabPage == lcgMark) //MARKING
            {
                bbiNew.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
                if (lblStatus.Text.Trim() == "READ-ONLY")
                {
                    bbiSave.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;

                    if (txeMarkingNo.Text.Trim() != "")
                    {
                        if (chkReadWrite == 1)
                            rpgManage.Visible = true;

                        //bbiCLONE.Enabled = true;

                        string OIDUSER = glueCreateBy.EditValue.ToString();
                        if (UserLogin.OIDUser.ToString() == OIDUSER)
                        {
                            bbiUPDATE.Enabled = true;
                            //bbiREVISE.Enabled = true;
                            //bbiDELBILL.Enabled = true;

                            string SMPLStatus = this.DBC.DBQuery("SELECT TOP (1) Status FROM Marking WHERE (MarkingNo = N'" + txeMarkingNo.Text.Trim() + "')").getString();
                            if (SMPLStatus == "0" || SMPLStatus == "")
                                bbiDELBILL.Enabled = false;
                            else
                                bbiDELBILL.Enabled = true;
                        }
                        else
                        {
                            bbiUPDATE.Enabled = false;
                            //bbiREVISE.Enabled = false;
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
            else if (tabMARKING.SelectedTabPage == lcgMarkDetail) //MARKING DETAIL
            {
                bbiNew.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
                bbiEdit.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
                if (lblStatus.Text.Trim() == "READ-ONLY")
                {
                    bbiSave.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;

                    if (txeMarkingNo.Text.Trim() != "")
                    {
                        if (chkReadWrite == 1)
                            rpgManage.Visible = true;

                        //bbiCLONE.Enabled = true;

                        string OIDUSER = glueCreateBy.EditValue.ToString();
                        if (UserLogin.OIDUser.ToString() == OIDUSER)
                        {
                            bbiUPDATE.Enabled = true;
                            //bbiREVISE.Enabled = true;
                            //bbiDELBILL.Enabled = true;

                            string SMPLStatus = DBC.DBQuery("SELECT TOP (1) Status FROM Marking WHERE (MarkingNo = N'" + txeMarkingNo.Text.Trim() + "')").getString();
                            if (SMPLStatus == "0" || SMPLStatus == "")
                                bbiDELBILL.Enabled = false;
                            else
                                bbiDELBILL.Enabled = true;
                        }
                        else
                        {
                            bbiUPDATE.Enabled = false;
                            //bbiREVISE.Enabled = false;
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

        private void btnInsert_Click(object sender, EventArgs e)
        {
            DataTable dtP = (DataTable)gcFBPart.DataSource;
            if (dtP != null)
            {
                if (dtP.Rows.Count == 0)
                {
                    FUNC.msgWarning("Please select fabric data.");
                    return;
                }
            }

            gvSTD.CloseEditor();
            gvSTD.UpdateCurrentRow();

            gvPOS.CloseEditor();
            gvPOS.UpdateCurrentRow();

            gvNEG.CloseEditor();
            gvNEG.UpdateCurrentRow();

            DataTable dtMDT = (DataTable)gcMDT.DataSource;
            if (dtMDT == null)
            {
                dtMDT = new DataTable();
                dtMDT.Columns.Add("ID", typeof(String));
                dtMDT.Columns.Add("OIDITEM", typeof(Int32));
                dtMDT.Columns.Add("ItemCode", typeof(String));
                dtMDT.Columns.Add("ItemDescription", typeof(String));
                dtMDT.Columns.Add("SMPLPatternNo", typeof(String));
                dtMDT.Columns.Add("FBPartsID", typeof(String));
                dtMDT.Columns.Add("FBParts", typeof(String));
                dtMDT.Columns.Add("VendorFBCode", typeof(String));
                dtMDT.Columns.Add("SampleLotNo", typeof(String));
                dtMDT.Columns.Add("DetailsType", typeof(String));
                dtMDT.Columns.Add("TotalWidth", typeof(Double));
                dtMDT.Columns.Add("UsableWidth", typeof(Double));
                dtMDT.Columns.Add("WeightGM2", typeof(Double));
                dtMDT.Columns.Add("OIDSIZE", typeof(Int32));
                dtMDT.Columns.Add("SizeName", typeof(String));
                dtMDT.Columns.Add("ActualLengthCm", typeof(Double));
                dtMDT.Columns.Add("QtyPcs", typeof(Int32));
                dtMDT.Columns.Add("LengthBodyCm", typeof(Double));
                dtMDT.Columns.Add("LengthBodyM", typeof(Double));
                dtMDT.Columns.Add("LengthBodyInc", typeof(Double));
                dtMDT.Columns.Add("LengthBodyYrd", typeof(Double));
                dtMDT.Columns.Add("WeightMg", typeof(Double));
                dtMDT.Columns.Add("WeightPcs", typeof(Double));
                dtMDT.Columns.Add("OIDSMPLDT", typeof(String));
                dtMDT.Columns.Add("PatternSizeZone", typeof(String));
            }

            //STD
            DataTable dtSTD = (DataTable)gcSTD.DataSource;
            if (dtSTD != null)
            {
                int runLoop = 0;
                foreach (DataRow drSTD in dtSTD.Rows)
                {
                    string OIDITEM = drSTD["OIDITEM"].ToString();
                    string OIDSIZE = drSTD["OIDSIZE"].ToString();
                    string SizeName = drSTD["SizeName"].ToString();

                    string TotalWidth = txtTotal_Standard.Text.Trim() == "" ? "0" : txtTotal_Standard.Text.Trim();
                    string UsableWidth = txtUsable_Standard.Text.Trim() == "" ? "0" : txtUsable_Standard.Text.Trim();
                    string WeightGM2 = txtWeight_Standard.Text.Trim() == "" ? "0" : txtWeight_Standard.Text.Trim();

                    string ActualLengthCm = drSTD["ActualLengthCm"].ToString();     ActualLengthCm = ActualLengthCm == "" ? "0" : ActualLengthCm;
                    string QtyPcs = drSTD["QtyPcs"].ToString();                     QtyPcs = QtyPcs == "" ? "0" : QtyPcs;
                    string LengthBodyCm = drSTD["LengthBodyCm"].ToString();         LengthBodyCm = LengthBodyCm == "" ? "0" : LengthBodyCm;
                    string LengthBodyM = drSTD["LengthBodyM"].ToString();           LengthBodyM = LengthBodyM == "" ? "0" : LengthBodyM;
                    string LengthBodyInc = drSTD["LengthBodyInc"].ToString();       LengthBodyInc = LengthBodyInc == "" ? "0" : LengthBodyInc;
                    string LengthBodyYrd = drSTD["LengthBodyYrd"].ToString();       LengthBodyYrd = LengthBodyYrd == "" ? "0" : LengthBodyYrd;
                    string WeightMg = drSTD["WeightMg"].ToString();                 WeightMg = WeightMg == "" ? "0" : WeightMg;
                    string WeightPcs = drSTD["WeightPcs"].ToString();               WeightPcs = WeightPcs == "" ? "0" : WeightPcs;

                    //เช็คว่าใน gcMDT มีข้อมูลแถวนี้แล้วหรือไม่ ถ้ามี ให้อัพเดตข้อมูลแถวนั้น ถ้าไม่มีให้เพิ่มแถวใหม่ใน gcMDT
                    if (dtMDT != null)
                    {
                        int chkRow = -1;
                        int runChkLoop = 0;
                        foreach (DataRow drMDT in dtMDT.Rows)
                        {
                            string xOIDITEM = drMDT["OIDITEM"].ToString();
                            string xOIDSIZE = drMDT["OIDSIZE"].ToString();
                            string xDetailsType = drMDT["DetailsType"].ToString();
                            if (OIDITEM == xOIDITEM && OIDSIZE == xOIDSIZE && xDetailsType == "0")
                            {
                                chkRow = runChkLoop;
                                break;
                            }
                            runChkLoop++;
                        }

                        if (chkRow == -1) //new row
                        {
                            dtMDT.Rows.Add("", lblOIDITEM.Text, lblItemCode.Text, lblItemDescription.Text, txePatternNo.Text, lblFBPartID.Text, lblParts.Text, txeVendFBCode.Text, txeSampleLotNo.Text, "0", TotalWidth, UsableWidth, WeightGM2, OIDSIZE, SizeName, ActualLengthCm, QtyPcs, LengthBodyCm, LengthBodyM, LengthBodyInc, LengthBodyYrd, WeightMg, WeightPcs, lblOIDSMPLDT.Text, sluePatternSizeZone.EditValue.ToString());
                        }
                        else //update row
                        {
                            dtMDT.Rows[chkRow].SetField("TotalWidth", TotalWidth);
                            dtMDT.Rows[chkRow].SetField("UsableWidth", UsableWidth);
                            dtMDT.Rows[chkRow].SetField("WeightGM2", WeightGM2);
                            dtMDT.Rows[chkRow].SetField("ActualLengthCm", ActualLengthCm);
                            dtMDT.Rows[chkRow].SetField("QtyPcs", QtyPcs);
                            dtMDT.Rows[chkRow].SetField("LengthBodyCm", LengthBodyCm);
                            dtMDT.Rows[chkRow].SetField("LengthBodyM", LengthBodyM);
                            dtMDT.Rows[chkRow].SetField("LengthBodyInc", LengthBodyInc);
                            dtMDT.Rows[chkRow].SetField("LengthBodyYrd", LengthBodyYrd);
                            dtMDT.Rows[chkRow].SetField("WeightMg", WeightMg);
                            dtMDT.Rows[chkRow].SetField("WeightPcs", WeightPcs);
                        }
                    }
                    else //new row
                    {
                        dtMDT.Rows.Add("", lblOIDITEM.Text, lblItemCode.Text, lblItemDescription.Text, txePatternNo.Text, lblFBPartID.Text, lblParts.Text, txeVendFBCode.Text, txeSampleLotNo.Text, "0", TotalWidth, UsableWidth, WeightGM2, OIDSIZE, SizeName, ActualLengthCm, QtyPcs, LengthBodyCm, LengthBodyM, LengthBodyInc, LengthBodyYrd, WeightMg, WeightPcs, lblOIDSMPLDT.Text, sluePatternSizeZone.EditValue.ToString());
                    }
                    runLoop++;
                }
            }

            //POS
            DataTable dtPOS = (DataTable)gcPOS.DataSource;
            if (dtPOS != null)
            {
                int runLoop = 0;
                foreach (DataRow drPOS in dtPOS.Rows)
                {
                    string OIDITEM = drPOS["OIDITEM"].ToString();
                    string OIDSIZE = drPOS["OIDSIZE"].ToString();
                    string SizeName = drPOS["SizeName"].ToString();

                    string TotalWidth = txtTotal_Positive.Text.Trim() == "" ? "0" : txtTotal_Positive.Text.Trim();
                    string UsableWidth = txtUsable_Positive.Text.Trim() == "" ? "0" : txtUsable_Positive.Text.Trim();
                    string WeightGM2 = txtWeight_Positive.Text.Trim() == "" ? "0" : txtWeight_Positive.Text.Trim();

                    string ActualLengthCm = drPOS["ActualLengthCm"].ToString();     ActualLengthCm = ActualLengthCm == "" ? "0" : ActualLengthCm;
                    string QtyPcs = drPOS["QtyPcs"].ToString();                     QtyPcs = QtyPcs == "" ? "0" : QtyPcs;
                    string LengthBodyCm = drPOS["LengthBodyCm"].ToString();         LengthBodyCm = LengthBodyCm == "" ? "0" : LengthBodyCm;
                    string LengthBodyM = drPOS["LengthBodyM"].ToString();           LengthBodyM = LengthBodyM == "" ? "0" : LengthBodyM;
                    string LengthBodyInc = drPOS["LengthBodyInc"].ToString();       LengthBodyInc = LengthBodyInc == "" ? "0" : LengthBodyInc;
                    string LengthBodyYrd = drPOS["LengthBodyYrd"].ToString();       LengthBodyYrd = LengthBodyYrd == "" ? "0" : LengthBodyYrd;
                    string WeightMg = drPOS["WeightMg"].ToString();                 WeightMg = WeightMg == "" ? "0" : WeightMg;
                    string WeightPcs = drPOS["WeightPcs"].ToString();               WeightPcs = WeightPcs == "" ? "0" : WeightPcs;

                    //เช็คว่าใน gcMDT มีข้อมูลแถวนี้แล้วหรือไม่ ถ้ามี ให้อัพเดตข้อมูลแถวนั้น ถ้าไม่มีให้เพิ่มแถวใหม่ใน gcMDT
                    if (dtMDT != null)
                    {
                        int chkRow = -1;
                        int runChkLoop = 0;
                        foreach (DataRow drMDT in dtMDT.Rows)
                        {
                            string xOIDITEM = drMDT["OIDITEM"].ToString();
                            string xOIDSIZE = drMDT["OIDSIZE"].ToString();
                            string xDetailsType = drMDT["DetailsType"].ToString();
                            if (OIDITEM == xOIDITEM && OIDSIZE == xOIDSIZE && xDetailsType == "1")
                            {
                                chkRow = runChkLoop;
                                break;
                            }
                            runChkLoop++;
                        }

                        if (chkRow == -1) //new row
                        {
                            dtMDT.Rows.Add("", lblOIDITEM.Text, lblItemCode.Text, lblItemDescription.Text, txePatternNo.Text, lblFBPartID.Text, lblParts.Text, txeVendFBCode.Text, txeSampleLotNo.Text, "1", TotalWidth, UsableWidth, WeightGM2, OIDSIZE, SizeName, ActualLengthCm, QtyPcs, LengthBodyCm, LengthBodyM, LengthBodyInc, LengthBodyYrd, WeightMg, WeightPcs, lblOIDSMPLDT.Text, sluePatternSizeZone.EditValue.ToString());
                        }
                        else //update row
                        {
                            dtMDT.Rows[chkRow].SetField("TotalWidth", TotalWidth);
                            dtMDT.Rows[chkRow].SetField("UsableWidth", UsableWidth);
                            dtMDT.Rows[chkRow].SetField("WeightGM2", WeightGM2);
                            dtMDT.Rows[chkRow].SetField("ActualLengthCm", ActualLengthCm);
                            dtMDT.Rows[chkRow].SetField("QtyPcs", QtyPcs);
                            dtMDT.Rows[chkRow].SetField("LengthBodyCm", LengthBodyCm);
                            dtMDT.Rows[chkRow].SetField("LengthBodyM", LengthBodyM);
                            dtMDT.Rows[chkRow].SetField("LengthBodyInc", LengthBodyInc);
                            dtMDT.Rows[chkRow].SetField("LengthBodyYrd", LengthBodyYrd);
                            dtMDT.Rows[chkRow].SetField("WeightMg", WeightMg);
                            dtMDT.Rows[chkRow].SetField("WeightPcs", WeightPcs);
                        }
                    }
                    else //new row
                    {
                        dtMDT.Rows.Add("", lblOIDITEM.Text, lblItemCode.Text, lblItemDescription.Text, txePatternNo.Text, lblFBPartID.Text, lblParts.Text, txeVendFBCode.Text, txeSampleLotNo.Text, "1", TotalWidth, UsableWidth, WeightGM2, OIDSIZE, SizeName, ActualLengthCm, QtyPcs, LengthBodyCm, LengthBodyM, LengthBodyInc, LengthBodyYrd, WeightMg, WeightPcs, lblOIDSMPLDT.Text, sluePatternSizeZone.EditValue.ToString());
                    }
                    runLoop++;
                }
            }

            //NEG
            DataTable dtNEG = (DataTable)gcNEG.DataSource;
            if (dtNEG != null)
            {
                int runLoop = 0;
                foreach (DataRow drNEG in dtNEG.Rows)
                {
                    string OIDITEM = drNEG["OIDITEM"].ToString();
                    string OIDSIZE = drNEG["OIDSIZE"].ToString();
                    string SizeName = drNEG["SizeName"].ToString();

                    string TotalWidth = txtTotal_Negative.Text.Trim() == "" ? "0" : txtTotal_Negative.Text.Trim();
                    string UsableWidth = txtUsable_Negative.Text.Trim() == "" ? "0" : txtUsable_Negative.Text.Trim();
                    string WeightGM2 = txtWeight_Negative.Text.Trim() == "" ? "0" : txtWeight_Negative.Text.Trim();

                    string ActualLengthCm = drNEG["ActualLengthCm"].ToString();     ActualLengthCm = ActualLengthCm == "" ? "0" : ActualLengthCm;
                    string QtyPcs = drNEG["QtyPcs"].ToString();                     QtyPcs = QtyPcs == "" ? "0" : QtyPcs;
                    string LengthBodyCm = drNEG["LengthBodyCm"].ToString();         LengthBodyCm = LengthBodyCm == "" ? "0" : LengthBodyCm;
                    string LengthBodyM = drNEG["LengthBodyM"].ToString();           LengthBodyM = LengthBodyM == "" ? "0" : LengthBodyM;
                    string LengthBodyInc = drNEG["LengthBodyInc"].ToString();       LengthBodyInc = LengthBodyInc == "" ? "0" : LengthBodyInc;
                    string LengthBodyYrd = drNEG["LengthBodyYrd"].ToString();       LengthBodyYrd = LengthBodyYrd == "" ? "0" : LengthBodyYrd;
                    string WeightMg = drNEG["WeightMg"].ToString();                 WeightMg = WeightMg == "" ? "0" : WeightMg;
                    string WeightPcs = drNEG["WeightPcs"].ToString();               WeightPcs = WeightPcs == "" ? "0" : WeightPcs;

                    //เช็คว่าใน gcMDT มีข้อมูลแถวนี้แล้วหรือไม่ ถ้ามี ให้อัพเดตข้อมูลแถวนั้น ถ้าไม่มีให้เพิ่มแถวใหม่ใน gcMDT
                    if (dtMDT != null)
                    {
                        int chkRow = -1;
                        int runChkLoop = 0;
                        foreach (DataRow drMDT in dtMDT.Rows)
                        {
                            string xOIDITEM = drMDT["OIDITEM"].ToString();
                            string xOIDSIZE = drMDT["OIDSIZE"].ToString();
                            string xDetailsType = drMDT["DetailsType"].ToString();
                            if (OIDITEM == xOIDITEM && OIDSIZE == xOIDSIZE && xDetailsType == "2")
                            {
                                chkRow = runChkLoop;
                                break;
                            }
                            runChkLoop++;
                        }

                        if (chkRow == -1) //new row
                        {
                            dtMDT.Rows.Add("", lblOIDITEM.Text, lblItemCode.Text, lblItemDescription.Text, txePatternNo.Text, lblFBPartID.Text, lblParts.Text, txeVendFBCode.Text, txeSampleLotNo.Text, "2", TotalWidth, UsableWidth, WeightGM2, OIDSIZE, SizeName, ActualLengthCm, QtyPcs, LengthBodyCm, LengthBodyM, LengthBodyInc, LengthBodyYrd, WeightMg, WeightPcs, lblOIDSMPLDT.Text, sluePatternSizeZone.EditValue.ToString());
                        }
                        else //update row
                        {
                            dtMDT.Rows[chkRow].SetField("TotalWidth", TotalWidth);
                            dtMDT.Rows[chkRow].SetField("UsableWidth", UsableWidth);
                            dtMDT.Rows[chkRow].SetField("WeightGM2", WeightGM2);
                            dtMDT.Rows[chkRow].SetField("ActualLengthCm", ActualLengthCm);
                            dtMDT.Rows[chkRow].SetField("QtyPcs", QtyPcs);
                            dtMDT.Rows[chkRow].SetField("LengthBodyCm", LengthBodyCm);
                            dtMDT.Rows[chkRow].SetField("LengthBodyM", LengthBodyM);
                            dtMDT.Rows[chkRow].SetField("LengthBodyInc", LengthBodyInc);
                            dtMDT.Rows[chkRow].SetField("LengthBodyYrd", LengthBodyYrd);
                            dtMDT.Rows[chkRow].SetField("WeightMg", WeightMg);
                            dtMDT.Rows[chkRow].SetField("WeightPcs", WeightPcs);
                        }
                    }
                    else //new row
                    {
                        dtMDT.Rows.Add("", lblOIDITEM.Text, lblItemCode.Text, lblItemDescription.Text, txePatternNo.Text, lblFBPartID.Text, lblParts.Text, txeVendFBCode.Text, txeSampleLotNo.Text, "2", TotalWidth, UsableWidth, WeightGM2, OIDSIZE, SizeName, ActualLengthCm, QtyPcs, LengthBodyCm, LengthBodyM, LengthBodyInc, LengthBodyYrd, WeightMg, WeightPcs, lblOIDSMPLDT.Text, sluePatternSizeZone.EditValue.ToString());
                    }
                    runLoop++;
                }
            }


            gcMDT.DataSource = dtMDT;
            gcMDT.Update();

            ClearVFBCode();
        }

        public void getGrid_MARK(GridControl glName, DevExpress.XtraGrid.Views.Grid.GridView gvName, int OIDUser = 0, int showDoc = 1, int showUser = 0)
        {
            LoadUserSMPL();
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT    MK.OIDMARK AS ID, MK.MarkingNo AS [Marking No.], CONVERT(VARCHAR(10), MK.DocumentDate, 103) AS DocumentDate, MK.OIDSMPL, SMPL.SMPLNo AS [SMPL No.],  ");
            sbSQL.Append("          CASE WHEN MK.MarkingType = 0 THEN 'for Cost' ELSE CASE WHEN MK.MarkingType = 1 THEN 'for Production' ELSE CASE WHEN MK.MarkingType = 2 THEN 'Cutting Request' ELSE '' END END END AS MarkingType, ");
            sbSQL.Append("          BN.Name AS Branch, ");
            sbSQL.Append("          CASE WHEN MK.CuttingFactory = 0 THEN 'ALL' ELSE CASE WHEN MK.CuttingFactory = 1 THEN 'SAI-4' ELSE CASE WHEN MK.CuttingFactory = 2 THEN 'RAMA2' ELSE CASE WHEN MK.CuttingFactory = 3 THEN 'PTC' ELSE CASE WHEN MK.CuttingFactory = 4 THEN 'YEH' ELSE '' END END END END END AS CuttingFactory, ");
            sbSQL.Append("          CASE WHEN MK.SewingFactory = 0 THEN 'ALL' ELSE CASE WHEN MK.SewingFactory = 1 THEN 'SAI-4' ELSE CASE WHEN MK.SewingFactory = 2 THEN 'RAMA2' ELSE CASE WHEN MK.SewingFactory = 3 THEN 'PTC' ELSE CASE WHEN MK.SewingFactory = 4 THEN 'YEH' ELSE '' END END END END END AS SewingFactory, ");
            sbSQL.Append("          SMPL.Season, CUS.Name AS Customer, DP.Name AS[Sales Section], GC.CategoryName AS Category, PS.StyleName AS Style, ");
            sbSQL.Append("          SMPL.SMPLItem AS [SMPL Item], SMPL.SMPLPatternNo AS [SMPL Pattern No.], ");
            sbSQL.Append("          (CASE smpl.PatternSizeZone WHEN 0 THEN 'Japan' WHEN 1 THEN 'Europe' WHEN 2 THEN 'US' END) AS[Pattern Size Zone], SMPL.ReferenceNo, ");
            sbSQL.Append("          US.FullName AS CreatedBy, SMPL.CreatedDate, US2.FullName AS UpdatedBy, SMPL.UpdatedDate, MK.Status, MK.CreatedBy AS ByCreated ");
            sbSQL.Append("FROM      Marking AS MK INNER JOIN ");
            sbSQL.Append("          SMPLRequest AS SMPL ON MK.OIDSMPL = SMPL.OIDSMPL LEFT OUTER JOIN ");
            sbSQL.Append("          Branchs AS BN ON SMPL.OIDBranch = BN.OIDBranch LEFT OUTER JOIN ");
            sbSQL.Append("          Departments AS DP ON SMPL.OIDDEPT = DP.OIDDEPT LEFT OUTER JOIN ");
            sbSQL.Append("          Customer AS CUS ON SMPL.OIDCUST = CUS.OIDCUST LEFT OUTER JOIN ");
            sbSQL.Append("          GarmentCategory AS GC ON SMPL.OIDCATEGORY = GC.OIDGCATEGORY LEFT OUTER JOIN ");
            sbSQL.Append("          ProductStyle AS PS ON SMPL.OIDSTYLE = PS.OIDSTYLE LEFT OUTER JOIN ");
            sbSQL.Append("          Users AS US ON MK.CreatedBy = US.OIDUSER LEFT OUTER JOIN ");
            sbSQL.Append("          Users AS US2 ON MK.UpdatedBy = US2.OIDUSER ");
            sbSQL.Append("WHERE (MK.MarkingNo <> N'')  ");
            if (showDoc == 1)
                sbSQL.Append("AND   (MK.Status = 1)  ");
            if (showUser == 0)
                sbSQL.Append("AND   (MK.CreatedBy = '" + OIDUser + "') ");
            sbSQL.Append("ORDER BY MK.CreatedDate DESC ");
            new ObjDE.setGridControl(glName, gvName, sbSQL).getData(false, false, false, true);

            gvName.Columns["ID"].Visible = false;
            gvName.Columns["OIDSMPL"].Visible = false;
            gvName.Columns["Status"].Visible = false;
            gvName.Columns["ByCreated"].Visible = false;

            gvName.Columns["Marking No."].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gvName.Columns["DocumentDate"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gvName.Columns["SMPL No."].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gvName.Columns["Pattern Size Zone"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gvName.Columns["Season"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;

            gvName.Columns[0].Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
            gvName.Columns[1].Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
            gvName.Columns[2].Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
            gvName.Columns[3].Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
            gvName.Columns[4].Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;

        }

        private void HideSelectDoc()
        {
            bbiPrint.Enabled = false;
            bbiPrintPDF.Enabled = false;
            bbiUPDATE.Enabled = false;
            //bbiREVISE.Enabled = false;
            //bbiCLONE.Enabled = false;
            bbiDELBILL.Enabled = false;
        }

        private void bbiRefresh_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (rgDocActive.EditValue == null)
                rgDocActive.EditValue = 1;
            if (rgDocUser.EditValue == null)
                rgDocUser.EditValue = 0;

            tabMARKING.SelectedTabPage = lcgList;
            getGrid_MARK(gcMARK, gvMARK, UserLogin.OIDUser, Convert.ToInt32(rgDocActive.EditValue.ToString()), Convert.ToInt32(rgDocUser.EditValue.ToString()));
            HideSelectDoc();
        }
        
        private void gvListofFabric_DoubleClick(object sender, EventArgs e)
        {
            DXMouseEventArgs ea = e as DXMouseEventArgs;
            GridView view = sender as GridView;
            GridHitInfo info = view.CalcHitInfo(ea.Location);
            if (info.InRow || info.InRowCell)
            {
                ClearVFBCode();
                GridView gv = gvListofFabric;
                string OIDITEM = gv.GetFocusedRowCellValue("OIDITEM").ToString();
                string ItemCode = gv.GetFocusedRowCellValue("ItemCode").ToString();
                string ItemDescription = gv.GetFocusedRowCellValue("ItemDescription").ToString();
                lblITEMSEL.Text = ItemCode + " : " + ItemDescription;

                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT DISTINCT SFB.OIDITEM, SRQ.SMPLPatternNo, SRQ.PatternSizeZone, SFB.VendFBCode,  ");
                sbSQL.Append("       (SELECT ColorName + ', ' AS 'data()' FROM(SELECT DISTINCT APC.ColorName FROM SMPLRequestFabric AS AFB INNER JOIN ProductColor AS APC ON AFB.OIDCOLOR = APC.OIDCOLOR WHERE AFB.OIDSMPLDT IN(SELECT OIDSMPLDT FROM SMPLQuantityRequired WHERE OIDSMPL = SRQ.OIDSMPL) AND AFB.OIDITEM = SFB.OIDITEM AND AFB.OIDCOLOR IS NOT NULL) AS OIDCOLOR FOR XML PATH('')) AS Color, ");
                sbSQL.Append("       (SELECT SMPLotNo + ', ' AS 'data()' FROM(SELECT DISTINCT SMPLotNo FROM SMPLRequestFabric WHERE OIDSMPLDT IN(SELECT OIDSMPLDT FROM SMPLQuantityRequired WHERE OIDSMPL = SRQ.OIDSMPL) AND OIDITEM = SFB.OIDITEM AND SMPLotNo IS NOT NULL) AS FBLot FOR XML PATH('')) AS SMPLotNo, ");
                sbSQL.Append("       (SELECT FBType + ', ' AS 'data()' FROM(SELECT DISTINCT FBType FROM SMPLRequestFabric WHERE OIDSMPLDT IN(SELECT OIDSMPLDT FROM SMPLQuantityRequired WHERE OIDSMPL = SRQ.OIDSMPL) AND OIDITEM = SFB.OIDITEM AND FBType IS NOT NULL) AS FBType FOR XML PATH('')) AS FabricType, ");
                sbSQL.Append("       (SELECT OIDGParts + ', ' AS 'data()' FROM(SELECT DISTINCT CONVERT(VARCHAR, SFBP.OIDGParts) AS OIDGParts FROM SMPLQuantityRequired AS QR INNER JOIN SMPLRequestFabric AS FB ON QR.OIDSMPLDT = FB.OIDSMPLDT AND QR.OIDSMPL = SRQ.OIDSMPL INNER JOIN SMPLRequestFabricParts AS SFBP ON FB.OIDSMPLFB = SFBP.OIDSMPLFB AND FB.OIDSMPLDT = SFBP.OIDSMPLDT AND FB.OIDITEM = SFB.OIDITEM AND SFBP.OIDGParts IS NOT NULL) AS OIDGParts FOR XML PATH('')) AS OIDGParts, ");
                sbSQL.Append("       (SELECT GarmentParts + ', ' AS 'data()' FROM(SELECT DISTINCT GP.OIDGParts, GP.GarmentParts FROM SMPLQuantityRequired AS SQR INNER JOIN SMPLRequestFabric AS FB ON SQR.OIDSMPLDT = FB.OIDSMPLDT AND SQR.OIDSMPL = SRQ.OIDSMPL INNER JOIN SMPLRequestFabricParts AS SFBP ON FB.OIDSMPLFB = SFBP.OIDSMPLFB AND FB.OIDSMPLDT = SFBP.OIDSMPLDT AND FB.OIDITEM = SFB.OIDITEM INNER JOIN GarmentParts AS GP ON SFBP.OIDGParts = GP.OIDGParts AND GP.GarmentParts IS NOT NULL) AS FParts FOR XML PATH('')) AS FabricParts, ");
                sbSQL.Append("       (SELECT OIDSMPLDT + ', ' AS 'data()' FROM(SELECT DISTINCT CONVERT(VARCHAR, SQR.OIDSMPLDT) AS OIDSMPLDT FROM SMPLQuantityRequired AS SQR INNER JOIN SMPLRequestFabric AS FB ON SQR.OIDSMPLDT = FB.OIDSMPLDT AND SQR.OIDSMPL = SRQ.OIDSMPL AND FB.OIDITEM = SFB.OIDITEM AND SQR.OIDSMPLDT IS NOT NULL) AS OIDSMPLDT  FOR XML PATH('')) AS OIDSMPLDT  ");
                sbSQL.Append("FROM   SMPLRequestFabric AS SFB INNER JOIN ");
                sbSQL.Append("       SMPLQuantityRequired AS SQR ON SFB.OIDSMPLDT = SQR.OIDSMPLDT INNER JOIN ");
                sbSQL.Append("       SMPLRequest AS SRQ ON SQR.OIDSMPL = SRQ.OIDSMPL AND SRQ.OIDSMPL = '" + lblID.Text.Trim() + "' ");
                sbSQL.Append("WHERE  (SFB.OIDITEM = '" + OIDITEM + "') ");
                DataTable dtVFBCode = this.DBC.DBQuery(sbSQL).getDataTable();
                if (dtVFBCode != null)
                {
                    btnInsert.Enabled = true;
                    int runRow = 0;
                    foreach (DataRow drVFBCode in dtVFBCode.Rows)
                    {
                        lblOIDITEM.Text = OIDITEM;
                        lblItemCode.Text = ItemCode;
                        lblItemDescription.Text = ItemDescription;

                        string SMPLPatternNo = drVFBCode["SMPLPatternNo"].ToString().Trim();
                        SMPLPatternNo = SMPLPatternNo.IndexOf(',') > -1 ? SMPLPatternNo.Substring(0, SMPLPatternNo.Length - 1) : SMPLPatternNo;
                        txePatternNo.Text = SMPLPatternNo;

                        string PatternSizeZone = drVFBCode["PatternSizeZone"].ToString().Trim();
                        PatternSizeZone = PatternSizeZone.IndexOf(',') > -1 ? PatternSizeZone.Substring(0, PatternSizeZone.Length - 1) : PatternSizeZone;
                        sluePatternSizeZone.EditValue = PatternSizeZone;

                        string VendFBCode = drVFBCode["VendFBCode"].ToString().Trim();
                        VendFBCode = VendFBCode.IndexOf(',') > -1 ? VendFBCode.Substring(0, VendFBCode.Length - 1) : VendFBCode;
                        txeVendFBCode.Text = VendFBCode;

                        string Color = drVFBCode["Color"].ToString().Trim();
                        Color = Color.IndexOf(',') > -1 ? Color.Substring(0, Color.Length - 1) : Color;
                        txeColor.Text = Color;

                        string SMPLotNo = drVFBCode["SMPLotNo"].ToString().Trim();
                        SMPLotNo = SMPLotNo.IndexOf(',') > -1 ? SMPLotNo.Substring(0, SMPLotNo.Length - 1) : SMPLotNo;
                        txeSampleLotNo.Text = SMPLotNo;

                        string FabricType = drVFBCode["FabricType"].ToString().Trim();
                        FabricType = FabricType.IndexOf(',') > -1 ? FabricType.Substring(0, FabricType.Length - 1) : FabricType;
                        txeFBType.Text = FabricType;

                        string OIDGParts = drVFBCode["OIDGParts"].ToString().Trim();
                        OIDGParts = OIDGParts.IndexOf(',') > -1 ? OIDGParts.Substring(0, OIDGParts.Length - 1) : OIDGParts;
                        lblFBPartID.Text = OIDGParts.Replace(" ", "");

                        string FabricParts = drVFBCode["FabricParts"].ToString().Trim();
                        FabricParts = FabricParts.IndexOf(',') > -1 ? FabricParts.Substring(0, FabricParts.Length - 1) : FabricParts;
                        lblParts.Text = FabricParts;

                        string OIDSMPLDT = drVFBCode["OIDSMPLDT"].ToString().Trim();
                        OIDSMPLDT = OIDSMPLDT.IndexOf(',') > -1 ? OIDSMPLDT.Substring(0, OIDSMPLDT.Length - 1) : OIDSMPLDT;
                        lblOIDSMPLDT.Text = OIDSMPLDT.Replace(" ", "");

                        runRow++;
                    }

                    gcFBPart.DataSource = null;
                    sbSQL.Clear();
                    sbSQL.Append("SELECT DISTINCT SFP.OIDGParts AS FBPartID, GMP.GarmentParts AS FBPart ");
                    sbSQL.Append("FROM   SMPLRequestFabric AS SFB INNER JOIN ");
                    sbSQL.Append("       SMPLRequestFabricParts AS SFP ON SFB.OIDSMPLFB = SFP.OIDSMPLFB AND SFB.OIDSMPLDT = SFP.OIDSMPLDT INNER JOIN ");
                    sbSQL.Append("       GarmentParts AS GMP ON SFP.OIDGParts = GMP.OIDGParts INNER JOIN ");
                    sbSQL.Append("       SMPLQuantityRequired AS SQR ON SFB.OIDSMPLDT = SQR.OIDSMPLDT INNER JOIN ");
                    sbSQL.Append("       SMPLRequest AS SRQ ON SQR.OIDSMPL = SRQ.OIDSMPL AND SRQ.OIDSMPL = '" + lblID.Text.Trim() + "' ");
                    sbSQL.Append("WHERE (SFB.OIDITEM = '" + OIDITEM + "') ");
                    sbSQL.Append("ORDER BY FBPartID, FBPart ");
                    DataTable dtFBP = this.DBC.DBQuery(sbSQL).getDataTable();

                    DataTable dtPart = new DataTable();
                    dtPart.Columns.Add("SelPart", typeof(Boolean));
                    dtPart.Columns.Add("FBPartID", typeof(String));
                    dtPart.Columns.Add("FBPart", typeof(String));

                    if (dtFBP != null)
                    {
                        foreach (DataRow drFBP in dtFBP.Rows)
                        {
                            dtPart.Rows.Add(true, drFBP["FBPartID"].ToString(), drFBP["FBPart"].ToString());
                        }
                    }
                    gcFBPart.DataSource = dtPart;
                    gcFBPart.EndUpdate();
                    gcFBPart.ResumeLayout();
                    gvFBPart.ClearSelection();

                    txtTotal_Standard.Text = "0";
                    txtUsable_Standard.Text = "0";
                    txtWeight_Standard.Text = "0";
                    txtTotal_Positive.Text = "0";
                    txtUsable_Positive.Text = "0";
                    txtWeight_Positive.Text = "0";
                    txtTotal_Negative.Text = "0";
                    txtUsable_Negative.Text = "0";
                    txtWeight_Negative.Text = "0";

                    if (lblMarkID.Text != "")
                    {
                        sbSQL.Clear();
                        sbSQL.Append("SELECT DISTINCT DetailsType, FORMAT(TotalWidthSTD, '###0.#####') AS TotalWidthSTD, FORMAT(UsableWidth, '###0.#####') AS UsableWidth, FORMAT(GM2, '###0.#####') AS GM2 ");
                        sbSQL.Append("FROM   MarkingDetails ");
                        sbSQL.Append("WHERE (OIDMARK = '" + lblMarkID.Text + "') AND (OIDITEM = '" + OIDITEM + "') ");
                        DataTable dtTT = this.DBC.DBQuery(sbSQL).getDataTable();
                        if (dtTT != null)
                        {
                            foreach (DataRow drTT in dtTT.Rows)
                            {
                                string DetailsType = drTT["DetailsType"].ToString();
                                string TotalWidthSTD = drTT["TotalWidthSTD"].ToString();
                                string UsableWidth = drTT["UsableWidth"].ToString();
                                string GM2 = drTT["GM2"].ToString();
                                if (DetailsType == "0")
                                {
                                    txtTotal_Standard.Text = TotalWidthSTD;
                                    txtUsable_Standard.Text = UsableWidth;
                                    txtWeight_Standard.Text = GM2;
                                }
                                else if (DetailsType == "1")
                                {
                                    txtTotal_Positive.Text = TotalWidthSTD;
                                    txtUsable_Positive.Text = UsableWidth;
                                    txtWeight_Positive.Text = GM2;
                                }
                                else if (DetailsType == "2")
                                {
                                    txtTotal_Negative.Text = TotalWidthSTD;
                                    txtUsable_Negative.Text = UsableWidth;
                                    txtWeight_Negative.Text = GM2;
                                }
                            }
                        }
                    }

                    //STD
                    sbSQL.Clear();
                    sbSQL.Append("SELECT    '" + OIDITEM + "' AS OIDITEM, SQR.OIDSIZE, PS.SizeName, FORMAT(SUM(ISNULL(MKDT.PracticalLengthCM, 0)), '##0.####') AS ActualLengthCm, FORMAT(SUM(ISNULL(MKDT.QuantityPCS, SQR.Quantity)), '##0.####') AS QtyPcs, FORMAT(SUM(ISNULL(MKDT.LengthPer1CM, 0)), '##0.####') AS LengthBodyCm, ");
                    sbSQL.Append("          FORMAT(SUM(ISNULL(MKDT.LengthPer1M, 0)), '##0.####') AS LengthBodyM, FORMAT(SUM(ISNULL(MKDT.LengthPer1INCH, 0)), '##0.####') AS LengthBodyInc, FORMAT(SUM(ISNULL(MKDT.LengthPer1YARD, 0)), '##0.####') AS LengthBodyYrd, ");
                    sbSQL.Append("          FORMAT(SUM(ISNULL(MKDT.WeightG, 0)), '##0.####') AS WeightMg, FORMAT(SUM(ISNULL(MKDT.WeightKG, 0)), '##0.####') AS WeightPcs ");
                    sbSQL.Append("FROM      SMPLRequest AS SRQ INNER JOIN ");
                    sbSQL.Append("          (SELECT OIDSMPL, OIDSIZE, SUM(Quantity) AS Quantity FROM SMPLQuantityRequired GROUP BY OIDSMPL, OIDSIZE) AS SQR ON SRQ.OIDSMPL = SQR.OIDSMPL AND SRQ.OIDSMPL = '" + lblID.Text.Trim() + "' INNER JOIN ");
                    sbSQL.Append("          ProductSize AS PS ON PS.OIDSIZE = SQR.OIDSIZE LEFT OUTER JOIN ");
                    sbSQL.Append("          Marking AS MK ON SRQ.OIDSMPL = MK.OIDSMPL ");
                    if (lblMarkID.Text != "")
                        sbSQL.Append("      AND MK.OIDMARK = '" + lblMarkID.Text + "' ");
                    sbSQL.Append("          LEFT OUTER JOIN ");
                    sbSQL.Append("          MarkingDetails AS MKDT ON MK.OIDMARK = MKDT.OIDMARK AND PS.OIDSIZE = MKDT.OIDSIZE AND MKDT.OIDITEM = '" + OIDITEM + "' AND MKDT.DetailsType = 0 ");
                    sbSQL.Append("GROUP BY SQR.OIDSIZE, PS.SizeNo, PS.SizeName ");
                    sbSQL.Append("ORDER BY PS.SizeNo ");
                    new ObjDE.setGridControl(gcSTD, gvSTD, sbSQL).getData(false, false, true, false);

                    //POS
                    sbSQL.Clear();
                    sbSQL.Append("SELECT    '" + OIDITEM + "' AS OIDITEM, SQR.OIDSIZE, PS.SizeName, FORMAT(SUM(ISNULL(MKDT.PracticalLengthCM, 0)), '##0.####') AS ActualLengthCm, FORMAT(SUM(ISNULL(MKDT.QuantityPCS, SQR.Quantity)), '##0.####') AS QtyPcs, FORMAT(SUM(ISNULL(MKDT.LengthPer1CM, 0)), '##0.####') AS LengthBodyCm, ");
                    sbSQL.Append("          FORMAT(SUM(ISNULL(MKDT.LengthPer1M, 0)), '##0.####') AS LengthBodyM, FORMAT(SUM(ISNULL(MKDT.LengthPer1INCH, 0)), '##0.####') AS LengthBodyInc, FORMAT(SUM(ISNULL(MKDT.LengthPer1YARD, 0)), '##0.####') AS LengthBodyYrd, ");
                    sbSQL.Append("          FORMAT(SUM(ISNULL(MKDT.WeightG, 0)), '##0.####') AS WeightMg, FORMAT(SUM(ISNULL(MKDT.WeightKG, 0)), '##0.####') AS WeightPcs ");
                    sbSQL.Append("FROM      SMPLRequest AS SRQ INNER JOIN ");
                    sbSQL.Append("          (SELECT OIDSMPL, OIDSIZE, SUM(Quantity) AS Quantity FROM SMPLQuantityRequired GROUP BY OIDSMPL, OIDSIZE) AS SQR ON SRQ.OIDSMPL = SQR.OIDSMPL AND SRQ.OIDSMPL = '" + lblID.Text.Trim() + "' INNER JOIN ");
                    sbSQL.Append("          ProductSize AS PS ON PS.OIDSIZE = SQR.OIDSIZE LEFT OUTER JOIN ");
                    sbSQL.Append("          Marking AS MK ON SRQ.OIDSMPL = MK.OIDSMPL ");
                    if (lblMarkID.Text != "")
                        sbSQL.Append("      AND MK.OIDMARK = '" + lblMarkID.Text + "' ");
                    sbSQL.Append("          LEFT OUTER JOIN ");
                    sbSQL.Append("          MarkingDetails AS MKDT ON MK.OIDMARK = MKDT.OIDMARK AND PS.OIDSIZE = MKDT.OIDSIZE AND MKDT.OIDITEM = '" + OIDITEM + "' AND MKDT.DetailsType = 1 ");
                    sbSQL.Append("GROUP BY SQR.OIDSIZE, PS.SizeNo, PS.SizeName ");
                    sbSQL.Append("ORDER BY PS.SizeNo ");
                    new ObjDE.setGridControl(gcPOS, gvPOS, sbSQL).getData(false, false, true, false);

                    //NEG
                    sbSQL.Clear();
                    sbSQL.Append("SELECT    '" + OIDITEM + "' AS OIDITEM, SQR.OIDSIZE, PS.SizeName, FORMAT(SUM(ISNULL(MKDT.PracticalLengthCM, 0)), '##0.####') AS ActualLengthCm, FORMAT(SUM(ISNULL(MKDT.QuantityPCS, SQR.Quantity)), '##0.####') AS QtyPcs, FORMAT(SUM(ISNULL(MKDT.LengthPer1CM, 0)), '##0.####') AS LengthBodyCm, ");
                    sbSQL.Append("          FORMAT(SUM(ISNULL(MKDT.LengthPer1M, 0)), '##0.####') AS LengthBodyM, FORMAT(SUM(ISNULL(MKDT.LengthPer1INCH, 0)), '##0.####') AS LengthBodyInc, FORMAT(SUM(ISNULL(MKDT.LengthPer1YARD, 0)), '##0.####') AS LengthBodyYrd, ");
                    sbSQL.Append("          FORMAT(SUM(ISNULL(MKDT.WeightG, 0)), '##0.####') AS WeightMg, FORMAT(SUM(ISNULL(MKDT.WeightKG, 0)), '##0.####') AS WeightPcs ");
                    sbSQL.Append("FROM      SMPLRequest AS SRQ INNER JOIN ");
                    sbSQL.Append("          (SELECT OIDSMPL, OIDSIZE, SUM(Quantity) AS Quantity FROM SMPLQuantityRequired GROUP BY OIDSMPL, OIDSIZE) AS SQR ON SRQ.OIDSMPL = SQR.OIDSMPL AND SRQ.OIDSMPL = '" + lblID.Text.Trim() + "' INNER JOIN ");
                    sbSQL.Append("          ProductSize AS PS ON PS.OIDSIZE = SQR.OIDSIZE LEFT OUTER JOIN ");
                    sbSQL.Append("          Marking AS MK ON SRQ.OIDSMPL = MK.OIDSMPL ");
                    if (lblMarkID.Text != "")
                        sbSQL.Append("      AND MK.OIDMARK = '" + lblMarkID.Text + "' ");
                    sbSQL.Append("          LEFT OUTER JOIN ");
                    sbSQL.Append("          MarkingDetails AS MKDT ON MK.OIDMARK = MKDT.OIDMARK AND PS.OIDSIZE = MKDT.OIDSIZE AND MKDT.OIDITEM = '" + OIDITEM + "' AND MKDT.DetailsType = 2 ");
                    sbSQL.Append("GROUP BY SQR.OIDSIZE, PS.SizeNo, PS.SizeName ");
                    sbSQL.Append("ORDER BY PS.SizeNo ");
                    new ObjDE.setGridControl(gcNEG, gvNEG, sbSQL).getData(false, false, true, false);


                    //ดึงข้อมูลจากในตารางด้านล่างขึ้นมา กรณีป้อนข้อมูลใหม่แล้วแต่ยังไม่ได้บันทึก เพื่อให้แสดงข้อมูลเป็นปัจจุบัน
                    DataTable dtMDT = (DataTable)gcMDT.DataSource;
                    if (dtMDT != null)
                    {
                        //Standard
                        DataTable dtSTD = (DataTable)gcSTD.DataSource;
                        if (dtSTD != null)
                        {
                            int stdRow = 0;
                            foreach (DataRow drSTD in dtSTD.Rows)
                            {
                                string xDetailsType = "0";
                                string xOIDSIZE = drSTD["OIDSIZE"].ToString();
                                string xOIDITEM = drSTD["OIDITEM"].ToString();

                                int rRow = 0;
                                foreach (DataRow drMDT in dtMDT.Rows)
                                {
                                    string OID_ITEM = drMDT["OIDITEM"].ToString();
                                    string DetailsType = drMDT["DetailsType"].ToString();
                                    string TotalWidth = drMDT["TotalWidth"].ToString();
                                    string UsableWidth = drMDT["UsableWidth"].ToString();
                                    string WeightGM2 = drMDT["WeightGM2"].ToString();
                                    string OIDSIZE = drMDT["OIDSIZE"].ToString();
                                    string ActualLengthCm = drMDT["ActualLengthCm"].ToString();
                                    string QtyPcs = drMDT["QtyPcs"].ToString();
                                    string LengthBodyCm = drMDT["LengthBodyCm"].ToString();
                                    string LengthBodyM = drMDT["LengthBodyM"].ToString();
                                    string LengthBodyInc = drMDT["LengthBodyInc"].ToString();
                                    string LengthBodyYrd = drMDT["LengthBodyYrd"].ToString();
                                    string WeightMg = drMDT["WeightMg"].ToString();
                                    string WeightPcs = drMDT["WeightPcs"].ToString();

                                    if (xDetailsType == DetailsType && xOIDSIZE == OIDSIZE && xOIDITEM == OID_ITEM)
                                    {
                                        txtTotal_Standard.Text = TotalWidth;
                                        txtUsable_Standard.Text = UsableWidth;
                                        txtWeight_Standard.Text = WeightGM2;
                                        dtSTD.Rows[stdRow].SetField("ActualLengthCm", ActualLengthCm);
                                        dtSTD.Rows[stdRow].SetField("QtyPcs", QtyPcs);
                                        dtSTD.Rows[stdRow].SetField("LengthBodyCm", LengthBodyCm);
                                        dtSTD.Rows[stdRow].SetField("LengthBodyM", LengthBodyM);
                                        dtSTD.Rows[stdRow].SetField("LengthBodyInc", LengthBodyInc);
                                        dtSTD.Rows[stdRow].SetField("LengthBodyYrd", LengthBodyYrd);
                                        dtSTD.Rows[stdRow].SetField("WeightMg", WeightMg);
                                        dtSTD.Rows[stdRow].SetField("WeightPcs", WeightPcs);
                                        break;
                                    }

                                    rRow++;
                                }

                                stdRow++;
                            }

                            gcSTD.DataSource = dtSTD;
                            gcSTD.Update();
                        }

                        //Positive
                        DataTable dtPOS = (DataTable)gcPOS.DataSource;
                        if (dtPOS != null)
                        {
                            int POSRow = 0;
                            foreach (DataRow drPOS in dtPOS.Rows)
                            {
                                string xDetailsType = "1";
                                string xOIDSIZE = drPOS["OIDSIZE"].ToString();
                                string xOIDITEM = drPOS["OIDITEM"].ToString();

                                int rRow = 0;
                                foreach (DataRow drMDT in dtMDT.Rows)
                                {
                                    string OID_ITEM = drMDT["OIDITEM"].ToString();
                                    string DetailsType = drMDT["DetailsType"].ToString();
                                    string TotalWidth = drMDT["TotalWidth"].ToString();
                                    string UsableWidth = drMDT["UsableWidth"].ToString();
                                    string WeightGM2 = drMDT["WeightGM2"].ToString();
                                    string OIDSIZE = drMDT["OIDSIZE"].ToString();
                                    string ActualLengthCm = drMDT["ActualLengthCm"].ToString();
                                    string QtyPcs = drMDT["QtyPcs"].ToString();
                                    string LengthBodyCm = drMDT["LengthBodyCm"].ToString();
                                    string LengthBodyM = drMDT["LengthBodyM"].ToString();
                                    string LengthBodyInc = drMDT["LengthBodyInc"].ToString();
                                    string LengthBodyYrd = drMDT["LengthBodyYrd"].ToString();
                                    string WeightMg = drMDT["WeightMg"].ToString();
                                    string WeightPcs = drMDT["WeightPcs"].ToString();

                                    if (xDetailsType == DetailsType && xOIDSIZE == OIDSIZE && xOIDITEM == OID_ITEM)
                                    {
                                        txtTotal_Positive.Text = TotalWidth;
                                        txtUsable_Positive.Text = UsableWidth;
                                        txtWeight_Positive.Text = WeightGM2;
                                        dtPOS.Rows[POSRow].SetField("ActualLengthCm", ActualLengthCm);
                                        dtPOS.Rows[POSRow].SetField("QtyPcs", QtyPcs);
                                        dtPOS.Rows[POSRow].SetField("LengthBodyCm", LengthBodyCm);
                                        dtPOS.Rows[POSRow].SetField("LengthBodyM", LengthBodyM);
                                        dtPOS.Rows[POSRow].SetField("LengthBodyInc", LengthBodyInc);
                                        dtPOS.Rows[POSRow].SetField("LengthBodyYrd", LengthBodyYrd);
                                        dtPOS.Rows[POSRow].SetField("WeightMg", WeightMg);
                                        dtPOS.Rows[POSRow].SetField("WeightPcs", WeightPcs);
                                        break;
                                    }

                                    rRow++;
                                }

                                POSRow++;
                            }

                            gcPOS.DataSource = dtPOS;
                            gcPOS.Update();
                        }

                        //Negative
                        DataTable dtNEG = (DataTable)gcNEG.DataSource;
                        if (dtNEG != null)
                        {
                            int NEGRow = 0;
                            foreach (DataRow drNEG in dtNEG.Rows)
                            {
                                string xDetailsType = "2";
                                string xOIDSIZE = drNEG["OIDSIZE"].ToString();
                                string xOIDITEM = drNEG["OIDITEM"].ToString();

                                int rRow = 0;
                                foreach (DataRow drMDT in dtMDT.Rows)
                                {
                                    string OID_ITEM = drMDT["OIDITEM"].ToString();
                                    string DetailsType = drMDT["DetailsType"].ToString();
                                    string TotalWidth = drMDT["TotalWidth"].ToString();
                                    string UsableWidth = drMDT["UsableWidth"].ToString();
                                    string WeightGM2 = drMDT["WeightGM2"].ToString();
                                    string OIDSIZE = drMDT["OIDSIZE"].ToString();
                                    string ActualLengthCm = drMDT["ActualLengthCm"].ToString();
                                    string QtyPcs = drMDT["QtyPcs"].ToString();
                                    string LengthBodyCm = drMDT["LengthBodyCm"].ToString();
                                    string LengthBodyM = drMDT["LengthBodyM"].ToString();
                                    string LengthBodyInc = drMDT["LengthBodyInc"].ToString();
                                    string LengthBodyYrd = drMDT["LengthBodyYrd"].ToString();
                                    string WeightMg = drMDT["WeightMg"].ToString();
                                    string WeightPcs = drMDT["WeightPcs"].ToString();

                                    if (xDetailsType == DetailsType && xOIDSIZE == OIDSIZE && xOIDITEM == OID_ITEM)
                                    {
                                        txtTotal_Negative.Text = TotalWidth;
                                        txtUsable_Negative.Text = UsableWidth;
                                        txtWeight_Negative.Text = WeightGM2;
                                        dtNEG.Rows[NEGRow].SetField("ActualLengthCm", ActualLengthCm);
                                        dtNEG.Rows[NEGRow].SetField("QtyPcs", QtyPcs);
                                        dtNEG.Rows[NEGRow].SetField("LengthBodyCm", LengthBodyCm);
                                        dtNEG.Rows[NEGRow].SetField("LengthBodyM", LengthBodyM);
                                        dtNEG.Rows[NEGRow].SetField("LengthBodyInc", LengthBodyInc);
                                        dtNEG.Rows[NEGRow].SetField("LengthBodyYrd", LengthBodyYrd);
                                        dtNEG.Rows[NEGRow].SetField("WeightMg", WeightMg);
                                        dtNEG.Rows[NEGRow].SetField("WeightPcs", WeightPcs);
                                        break;
                                    }

                                    rRow++;
                                }

                                NEGRow++;
                            }

                            gcNEG.DataSource = dtNEG;
                            gcNEG.Update();
                        }
                    }
    
                }
            }
            
        }

        private void gvSTD_RowClick(object sender, RowClickEventArgs e)
        {
            //ct.showInfoMessage(gvSTD.FocusedRowHandle.ToString());
        }

        private void gvPOS_RowClick(object sender, RowClickEventArgs e)
        {
            //ct.showInfoMessage(gvPOS.FocusedRowHandle.ToString());
            gvSTD.FocusedRowHandle = gvPOS.FocusedRowHandle;
        }

        private void gvNEG_RowClick(object sender, RowClickEventArgs e)
        {
            //ct.showInfoMessage(gvNEG.FocusedRowHandle.ToString());
            gvSTD.FocusedRowHandle = gvNEG.FocusedRowHandle;
        }

        private void gvMARK_DoubleClick(object sender, EventArgs e)
        {
            DXMouseEventArgs ea = e as DXMouseEventArgs;
            GridView view = sender as GridView;
            GridHitInfo info = view.CalcHitInfo(ea.Location);
            if (info.InRow || info.InRowCell)
            {
                GridView gv = gvMARK;
                string MARKNo = gv.GetFocusedRowCellValue("Marking No.").ToString();
                LoadMarkingDocument(MARKNo, "READ-ONLY");

                //SetReadOnly();
                tabMARKING.SelectedTabPage = lcgMark;

            }
        }
        
        private void gvSTD_ValidatingEditor(object sender, DevExpress.XtraEditors.Controls.BaseContainerValidateEditorEventArgs e)
        {

        }

        private void gvMDT_DoubleClick(object sender, EventArgs e)
        {
            DXMouseEventArgs ea = e as DXMouseEventArgs;
            GridView view = sender as GridView;
            GridHitInfo info = view.CalcHitInfo(ea.Location);
            if (info.InRow || info.InRowCell)
            {
                GridView gv = gvMDT;
                ClearVFBCode();
                string OIDITEM = gv.GetFocusedRowCellValue("OIDITEM").ToString();
                string ItemCode = gv.GetFocusedRowCellValue("ItemCode").ToString();
                string ItemDescription = gv.GetFocusedRowCellValue("ItemDescription").ToString();

                lblITEMSEL.Text = ItemCode + " : " + ItemDescription;

                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT DISTINCT SFB.OIDITEM, SRQ.SMPLPatternNo, (CASE SRQ.PatternSizeZone WHEN 0 THEN 'Japan' WHEN 1 THEN 'Europe' WHEN 2 THEN 'US' END) AS PatternSizeZone, SFB.VendFBCode,  ");
                sbSQL.Append("       (SELECT ColorName + ', ' AS 'data()' FROM(SELECT DISTINCT APC.ColorName FROM SMPLRequestFabric AS AFB INNER JOIN ProductColor AS APC ON AFB.OIDCOLOR = APC.OIDCOLOR WHERE AFB.OIDSMPLDT IN(SELECT OIDSMPLDT FROM SMPLQuantityRequired WHERE OIDSMPL = SRQ.OIDSMPL) AND AFB.OIDITEM = SFB.OIDITEM AND AFB.OIDCOLOR IS NOT NULL) AS OIDCOLOR FOR XML PATH('')) AS Color, ");
                sbSQL.Append("       (SELECT SMPLotNo + ', ' AS 'data()' FROM(SELECT DISTINCT SMPLotNo FROM SMPLRequestFabric WHERE OIDSMPLDT IN(SELECT OIDSMPLDT FROM SMPLQuantityRequired WHERE OIDSMPL = SRQ.OIDSMPL) AND OIDITEM = SFB.OIDITEM AND SMPLotNo IS NOT NULL) AS FBLot FOR XML PATH('')) AS SMPLotNo, ");
                sbSQL.Append("       (SELECT FBType + ', ' AS 'data()' FROM(SELECT DISTINCT FBType FROM SMPLRequestFabric WHERE OIDSMPLDT IN(SELECT OIDSMPLDT FROM SMPLQuantityRequired WHERE OIDSMPL = SRQ.OIDSMPL) AND OIDITEM = SFB.OIDITEM AND FBType IS NOT NULL) AS FBType FOR XML PATH('')) AS FabricType, ");
                sbSQL.Append("       (SELECT OIDGParts + ', ' AS 'data()' FROM(SELECT DISTINCT CONVERT(VARCHAR, SFBP.OIDGParts) AS OIDGParts FROM SMPLQuantityRequired AS QR INNER JOIN SMPLRequestFabric AS FB ON QR.OIDSMPLDT = FB.OIDSMPLDT AND QR.OIDSMPL = SRQ.OIDSMPL INNER JOIN SMPLRequestFabricParts AS SFBP ON FB.OIDSMPLFB = SFBP.OIDSMPLFB AND FB.OIDSMPLDT = SFBP.OIDSMPLDT AND FB.OIDITEM = SFB.OIDITEM AND SFBP.OIDGParts IS NOT NULL) AS OIDGParts FOR XML PATH('')) AS OIDGParts, ");
                sbSQL.Append("       (SELECT GarmentParts + ', ' AS 'data()' FROM(SELECT DISTINCT GP.OIDGParts, GP.GarmentParts FROM SMPLQuantityRequired AS SQR INNER JOIN SMPLRequestFabric AS FB ON SQR.OIDSMPLDT = FB.OIDSMPLDT AND SQR.OIDSMPL = SRQ.OIDSMPL INNER JOIN SMPLRequestFabricParts AS SFBP ON FB.OIDSMPLFB = SFBP.OIDSMPLFB AND FB.OIDSMPLDT = SFBP.OIDSMPLDT AND FB.OIDITEM = SFB.OIDITEM INNER JOIN GarmentParts AS GP ON SFBP.OIDGParts = GP.OIDGParts AND GP.GarmentParts IS NOT NULL) AS FParts FOR XML PATH('')) AS FabricParts, ");
                sbSQL.Append("       (SELECT OIDSMPLDT + ', ' AS 'data()' FROM(SELECT DISTINCT CONVERT(VARCHAR, SQR.OIDSMPLDT) AS OIDSMPLDT FROM SMPLQuantityRequired AS SQR INNER JOIN SMPLRequestFabric AS FB ON SQR.OIDSMPLDT = FB.OIDSMPLDT AND SQR.OIDSMPL = SRQ.OIDSMPL AND FB.OIDITEM = SFB.OIDITEM AND SQR.OIDSMPLDT IS NOT NULL) AS OIDSMPLDT  FOR XML PATH('')) AS OIDSMPLDT  ");
                sbSQL.Append("FROM   SMPLRequestFabric AS SFB INNER JOIN ");
                sbSQL.Append("       SMPLQuantityRequired AS SQR ON SFB.OIDSMPLDT = SQR.OIDSMPLDT INNER JOIN ");
                sbSQL.Append("       SMPLRequest AS SRQ ON SQR.OIDSMPL = SRQ.OIDSMPL AND SRQ.OIDSMPL = '" + lblID.Text.Trim() + "' ");
                sbSQL.Append("WHERE  (SFB.OIDITEM = '" + OIDITEM + "') ");
                DataTable dtVFBCode = this.DBC.DBQuery(sbSQL).getDataTable();
                if (dtVFBCode != null)
                {
                    btnInsert.Enabled = true;
                    int runRow = 0;
                    foreach (DataRow drVFBCode in dtVFBCode.Rows)
                    {
                        lblOIDITEM.Text = OIDITEM;
                        lblItemCode.Text = ItemCode;
                        lblItemDescription.Text = ItemDescription;

                        string SMPLPatternNo = drVFBCode["SMPLPatternNo"].ToString().Trim();
                        SMPLPatternNo = SMPLPatternNo.IndexOf(',') > -1 ? SMPLPatternNo.Substring(0, SMPLPatternNo.Length - 1) : SMPLPatternNo;
                        txePatternNo.Text = SMPLPatternNo;

                        string PatternSizeZone = drVFBCode["PatternSizeZone"].ToString().Trim();
                        PatternSizeZone = PatternSizeZone.IndexOf(',') > -1 ? PatternSizeZone.Substring(0, PatternSizeZone.Length - 1) : PatternSizeZone;
                        sluePatternSizeZone.EditValue = PatternSizeZone;

                        string VendFBCode = drVFBCode["VendFBCode"].ToString().Trim();
                        VendFBCode = VendFBCode.IndexOf(',') > -1 ? VendFBCode.Substring(0, VendFBCode.Length - 1) : VendFBCode;
                        txeVendFBCode.Text = VendFBCode;

                        string Color = drVFBCode["Color"].ToString().Trim();
                        Color = Color.IndexOf(',') > -1 ? Color.Substring(0, Color.Length - 1) : Color;
                        txeColor.Text = Color;

                        string SMPLotNo = drVFBCode["SMPLotNo"].ToString().Trim();
                        SMPLotNo = SMPLotNo.IndexOf(',') > -1 ? SMPLotNo.Substring(0, SMPLotNo.Length - 1) : SMPLotNo;
                        txeSampleLotNo.Text = SMPLotNo;

                        string FabricType = drVFBCode["FabricType"].ToString().Trim();
                        FabricType = FabricType.IndexOf(',') > -1 ? FabricType.Substring(0, FabricType.Length - 1) : FabricType;
                        txeFBType.Text = FabricType;

                        string OIDGParts = drVFBCode["OIDGParts"].ToString().Trim();
                        OIDGParts = OIDGParts.IndexOf(',') > -1 ? OIDGParts.Substring(0, OIDGParts.Length - 1) : OIDGParts;
                        lblFBPartID.Text = OIDGParts.Replace(" ", "");

                        string FabricParts = drVFBCode["FabricParts"].ToString().Trim();
                        FabricParts = FabricParts.IndexOf(',') > -1 ? FabricParts.Substring(0, FabricParts.Length - 1) : FabricParts;
                        lblParts.Text = FabricParts;

                        string OIDSMPLDT = drVFBCode["OIDSMPLDT"].ToString().Trim();
                        OIDSMPLDT = OIDSMPLDT.IndexOf(',') > -1 ? OIDSMPLDT.Substring(0, OIDSMPLDT.Length - 1) : OIDSMPLDT;
                        lblOIDSMPLDT.Text = OIDSMPLDT.Replace(" ", "");

                        runRow++;
                    }

                    gcFBPart.DataSource = null;
                    sbSQL.Clear();
                    sbSQL.Append("SELECT DISTINCT SFP.OIDGParts AS FBPartID, GMP.GarmentParts AS FBPart ");
                    sbSQL.Append("FROM   SMPLRequestFabric AS SFB INNER JOIN ");
                    sbSQL.Append("       SMPLRequestFabricParts AS SFP ON SFB.OIDSMPLFB = SFP.OIDSMPLFB AND SFB.OIDSMPLDT = SFP.OIDSMPLDT INNER JOIN ");
                    sbSQL.Append("       GarmentParts AS GMP ON SFP.OIDGParts = GMP.OIDGParts INNER JOIN ");
                    sbSQL.Append("       SMPLQuantityRequired AS SQR ON SFB.OIDSMPLDT = SQR.OIDSMPLDT INNER JOIN ");
                    sbSQL.Append("       SMPLRequest AS SRQ ON SQR.OIDSMPL = SRQ.OIDSMPL AND SRQ.OIDSMPL = '" + lblID.Text.Trim() + "' ");
                    sbSQL.Append("WHERE (SFB.OIDITEM = '" + OIDITEM + "') ");
                    sbSQL.Append("ORDER BY FBPartID, FBPart ");
                    DataTable dtFBP = this.DBC.DBQuery(sbSQL).getDataTable();

                    DataTable dtPart = new DataTable();
                    dtPart.Columns.Add("SelPart", typeof(Boolean));
                    dtPart.Columns.Add("FBPartID", typeof(String));
                    dtPart.Columns.Add("FBPart", typeof(String));

                    if (dtFBP != null)
                    {
                        foreach (DataRow drFBP in dtFBP.Rows)
                        {
                            dtPart.Rows.Add(true, drFBP["FBPartID"].ToString(), drFBP["FBPart"].ToString());
                        }
                    }
                    gcFBPart.DataSource = dtPart;
                    gcFBPart.EndUpdate();
                    gcFBPart.ResumeLayout();
                    gvFBPart.ClearSelection();

                    //STD
                    sbSQL.Clear();
                    sbSQL.Append("SELECT    '" + OIDITEM + "' AS OIDITEM, SQR.OIDSIZE, PS.SizeName, FORMAT(SUM(ISNULL(MKDT.PracticalLengthCM, 0)), '##0.####') AS ActualLengthCm, FORMAT(SUM(ISNULL(MKDT.QuantityPCS, SQR.Quantity)), '##0.####') AS QtyPcs, FORMAT(SUM(ISNULL(MKDT.LengthPer1CM, 0)), '##0.####') AS LengthBodyCm, ");
                    sbSQL.Append("          FORMAT(SUM(ISNULL(MKDT.LengthPer1M, 0)), '##0.####') AS LengthBodyM, FORMAT(SUM(ISNULL(MKDT.LengthPer1INCH, 0)), '##0.####') AS LengthBodyInc, FORMAT(SUM(ISNULL(MKDT.LengthPer1YARD, 0)), '##0.####') AS LengthBodyYrd, ");
                    sbSQL.Append("          FORMAT(SUM(ISNULL(MKDT.WeightG, 0)), '##0.####') AS WeightMg, FORMAT(SUM(ISNULL(MKDT.WeightKG, 0)), '##0.####') AS WeightPcs ");
                    sbSQL.Append("FROM      SMPLRequest AS SRQ INNER JOIN ");
                    sbSQL.Append("          SMPLQuantityRequired AS SQR ON SRQ.OIDSMPL = SQR.OIDSMPL AND SRQ.OIDSMPL = '" + lblID.Text.Trim() + "' INNER JOIN ");
                    sbSQL.Append("          ProductSize AS PS ON PS.OIDSIZE = SQR.OIDSIZE LEFT OUTER JOIN ");
                    sbSQL.Append("          Marking AS MK ON SRQ.OIDSMPL = MK.OIDSMPL LEFT OUTER JOIN ");
                    sbSQL.Append("          MarkingDetails AS MKDT ON MK.OIDMARK = MKDT.OIDMARK AND PS.OIDSIZE = MKDT.OIDSIZE AND MKDT.OIDITEM = '" + OIDITEM + "' AND MKDT.DetailsType = 0 ");
                    sbSQL.Append("GROUP BY SQR.OIDSIZE, PS.SizeNo, PS.SizeName ");
                    sbSQL.Append("ORDER BY PS.SizeNo ");
                    new ObjDE.setGridControl(gcSTD, gvSTD, sbSQL).getData(false, false, true, false);

                    //POS
                    sbSQL.Clear();
                    sbSQL.Append("SELECT    '" + OIDITEM + "' AS OIDITEM, SQR.OIDSIZE, PS.SizeName, FORMAT(SUM(ISNULL(MKDT.PracticalLengthCM, 0)), '##0.####') AS ActualLengthCm, FORMAT(SUM(ISNULL(MKDT.QuantityPCS, SQR.Quantity)), '##0.####') AS QtyPcs, FORMAT(SUM(ISNULL(MKDT.LengthPer1CM, 0)), '##0.####') AS LengthBodyCm, ");
                    sbSQL.Append("          FORMAT(SUM(ISNULL(MKDT.LengthPer1M, 0)), '##0.####') AS LengthBodyM, FORMAT(SUM(ISNULL(MKDT.LengthPer1INCH, 0)), '##0.####') AS LengthBodyInc, FORMAT(SUM(ISNULL(MKDT.LengthPer1YARD, 0)), '##0.####') AS LengthBodyYrd, ");
                    sbSQL.Append("          FORMAT(SUM(ISNULL(MKDT.WeightG, 0)), '##0.####') AS WeightMg, FORMAT(SUM(ISNULL(MKDT.WeightKG, 0)), '##0.####') AS WeightPcs ");
                    sbSQL.Append("FROM      SMPLRequest AS SRQ INNER JOIN ");
                    sbSQL.Append("          SMPLQuantityRequired AS SQR ON SRQ.OIDSMPL = SQR.OIDSMPL AND SRQ.OIDSMPL = '" + lblID.Text.Trim() + "' INNER JOIN ");
                    sbSQL.Append("          ProductSize AS PS ON PS.OIDSIZE = SQR.OIDSIZE LEFT OUTER JOIN ");
                    sbSQL.Append("          Marking AS MK ON SRQ.OIDSMPL = MK.OIDSMPL LEFT OUTER JOIN ");
                    sbSQL.Append("          MarkingDetails AS MKDT ON MK.OIDMARK = MKDT.OIDMARK AND PS.OIDSIZE = MKDT.OIDSIZE AND MKDT.OIDITEM = '" + OIDITEM + "' AND MKDT.DetailsType = 1 ");
                    sbSQL.Append("GROUP BY SQR.OIDSIZE, PS.SizeNo, PS.SizeName ");
                    sbSQL.Append("ORDER BY PS.SizeNo ");
                    new ObjDE.setGridControl(gcPOS, gvPOS, sbSQL).getData(false, false, true, false);

                    //NEG
                    sbSQL.Clear();
                    sbSQL.Append("SELECT    '" + OIDITEM + "' AS OIDITEM, SQR.OIDSIZE, PS.SizeName, FORMAT(SUM(ISNULL(MKDT.PracticalLengthCM, 0)), '##0.####') AS ActualLengthCm, FORMAT(SUM(ISNULL(MKDT.QuantityPCS, SQR.Quantity)), '##0.####') AS QtyPcs, FORMAT(SUM(ISNULL(MKDT.LengthPer1CM, 0)), '##0.####') AS LengthBodyCm, ");
                    sbSQL.Append("          FORMAT(SUM(ISNULL(MKDT.LengthPer1M, 0)), '##0.####') AS LengthBodyM, FORMAT(SUM(ISNULL(MKDT.LengthPer1INCH, 0)), '##0.####') AS LengthBodyInc, FORMAT(SUM(ISNULL(MKDT.LengthPer1YARD, 0)), '##0.####') AS LengthBodyYrd, ");
                    sbSQL.Append("          FORMAT(SUM(ISNULL(MKDT.WeightG, 0)), '##0.####') AS WeightMg, FORMAT(SUM(ISNULL(MKDT.WeightKG, 0)), '##0.####') AS WeightPcs ");
                    sbSQL.Append("FROM      SMPLRequest AS SRQ INNER JOIN ");
                    sbSQL.Append("          SMPLQuantityRequired AS SQR ON SRQ.OIDSMPL = SQR.OIDSMPL AND SRQ.OIDSMPL = '" + lblID.Text.Trim() + "' INNER JOIN ");
                    sbSQL.Append("          ProductSize AS PS ON PS.OIDSIZE = SQR.OIDSIZE LEFT OUTER JOIN ");
                    sbSQL.Append("          Marking AS MK ON SRQ.OIDSMPL = MK.OIDSMPL LEFT OUTER JOIN ");
                    sbSQL.Append("          MarkingDetails AS MKDT ON MK.OIDMARK = MKDT.OIDMARK AND PS.OIDSIZE = MKDT.OIDSIZE AND MKDT.OIDITEM = '" + OIDITEM + "' AND MKDT.DetailsType = 2 ");
                    sbSQL.Append("GROUP BY SQR.OIDSIZE, PS.SizeNo, PS.SizeName ");
                    sbSQL.Append("ORDER BY PS.SizeNo ");
                    new ObjDE.setGridControl(gcNEG, gvNEG, sbSQL).getData(false, false, true, false);

                }

                //ดึงข้อมูลจากตารางขึ้นไปแสดง
                DataTable dtMDT = (DataTable)gcMDT.DataSource;
                if (dtMDT != null)
                {
                    //STD
                    DataTable dtSTD = (DataTable)gcSTD.DataSource;
                    if (dtSTD != null)
                    {
                        int rowSTD = 0;
                        foreach (DataRow drSTD in dtSTD.Rows)
                        {
                            string STDSIZE = drSTD["OIDSIZE"].ToString();
                            string STDITEM = drSTD["OIDITEM"].ToString();

                            int rowMDT = 0;
                            foreach (DataRow drMDT in dtMDT.Rows)
                            {
                                string MDTSIZE = drMDT["OIDSIZE"].ToString();
                                string MDTITEM = drMDT["OIDITEM"].ToString();
                                string DetailsType = drMDT["DetailsType"].ToString();
                                if (STDSIZE == MDTSIZE && STDITEM == MDTITEM && DetailsType == "0")
                                {
                                    string TotalWidth = drMDT["TotalWidth"].ToString();
                                    string UsableWidth = drMDT["UsableWidth"].ToString();
                                    string WeightGM2 = drMDT["WeightGM2"].ToString();
                                    string ActualLengthCm = drMDT["ActualLengthCm"].ToString();
                                    string QtyPcs = drMDT["QtyPcs"].ToString();
                                    string LengthBodyCm = drMDT["LengthBodyCm"].ToString();
                                    string LengthBodyM = drMDT["LengthBodyM"].ToString();
                                    string LengthBodyInc = drMDT["LengthBodyInc"].ToString();
                                    string LengthBodyYrd = drMDT["LengthBodyYrd"].ToString();
                                    string WeightMg = drMDT["WeightMg"].ToString();
                                    string WeightPcs = drMDT["WeightPcs"].ToString();

                                    txtTotal_Standard.Text = TotalWidth;
                                    txtUsable_Standard.Text = UsableWidth;
                                    txtWeight_Standard.Text = WeightGM2;

                                    dtSTD.Rows[rowSTD].SetField("ActualLengthCm", ActualLengthCm);
                                    dtSTD.Rows[rowSTD].SetField("QtyPcs", QtyPcs);
                                    dtSTD.Rows[rowSTD].SetField("LengthBodyCm", LengthBodyCm);
                                    dtSTD.Rows[rowSTD].SetField("LengthBodyM", LengthBodyM);
                                    dtSTD.Rows[rowSTD].SetField("LengthBodyInc", LengthBodyInc);
                                    dtSTD.Rows[rowSTD].SetField("LengthBodyYrd", LengthBodyYrd);
                                    dtSTD.Rows[rowSTD].SetField("WeightMg", WeightMg);
                                    dtSTD.Rows[rowSTD].SetField("WeightPcs", WeightPcs);
                                    break;
                                }

                                rowMDT++;
                            }

                            rowSTD++;
                        }
                        gcSTD.DataSource = dtSTD;
                        gcSTD.Update();
                    }


                    //POS
                    DataTable dtPOS = (DataTable)gcPOS.DataSource;
                    if (dtPOS != null)
                    {
                        int rowPOS = 0;
                        foreach (DataRow drPOS in dtPOS.Rows)
                        {
                            string POSSIZE = drPOS["OIDSIZE"].ToString();
                            string POSITEM = drPOS["OIDITEM"].ToString();

                            int rowMDT = 0;
                            foreach (DataRow drMDT in dtMDT.Rows)
                            {
                                string MDTSIZE = drMDT["OIDSIZE"].ToString();
                                string MDTITEM = drMDT["OIDITEM"].ToString();
                                string DetailsType = drMDT["DetailsType"].ToString();
                                if (POSSIZE == MDTSIZE && POSITEM == MDTITEM && DetailsType == "1")
                                {
                                    string TotalWidth = drMDT["TotalWidth"].ToString();
                                    string UsableWidth = drMDT["UsableWidth"].ToString();
                                    string WeightGM2 = drMDT["WeightGM2"].ToString();
                                    string ActualLengthCm = drMDT["ActualLengthCm"].ToString();
                                    string QtyPcs = drMDT["QtyPcs"].ToString();
                                    string LengthBodyCm = drMDT["LengthBodyCm"].ToString();
                                    string LengthBodyM = drMDT["LengthBodyM"].ToString();
                                    string LengthBodyInc = drMDT["LengthBodyInc"].ToString();
                                    string LengthBodyYrd = drMDT["LengthBodyYrd"].ToString();
                                    string WeightMg = drMDT["WeightMg"].ToString();
                                    string WeightPcs = drMDT["WeightPcs"].ToString();

                                    txtTotal_Positive.Text = TotalWidth;
                                    txtUsable_Positive.Text = UsableWidth;
                                    txtWeight_Positive.Text = WeightGM2;

                                    dtPOS.Rows[rowPOS].SetField("ActualLengthCm", ActualLengthCm);
                                    dtPOS.Rows[rowPOS].SetField("QtyPcs", QtyPcs);
                                    dtPOS.Rows[rowPOS].SetField("LengthBodyCm", LengthBodyCm);
                                    dtPOS.Rows[rowPOS].SetField("LengthBodyM", LengthBodyM);
                                    dtPOS.Rows[rowPOS].SetField("LengthBodyInc", LengthBodyInc);
                                    dtPOS.Rows[rowPOS].SetField("LengthBodyYrd", LengthBodyYrd);
                                    dtPOS.Rows[rowPOS].SetField("WeightMg", WeightMg);
                                    dtPOS.Rows[rowPOS].SetField("WeightPcs", WeightPcs);
                                    break;
                                }

                                rowMDT++;
                            }

                            rowPOS++;
                        }
                        gcPOS.DataSource = dtPOS;
                        gcPOS.Update();
                    }


                    //NEG
                    DataTable dtNEG = (DataTable)gcNEG.DataSource;
                    if (dtNEG != null)
                    {
                        int rowNEG = 0;
                        foreach (DataRow drNEG in dtNEG.Rows)
                        {
                            string NEGSIZE = drNEG["OIDSIZE"].ToString();
                            string NEGITEM = drNEG["OIDITEM"].ToString();

                            int rowMDT = 0;
                            foreach (DataRow drMDT in dtMDT.Rows)
                            {
                                string MDTSIZE = drMDT["OIDSIZE"].ToString();
                                string MDTITEM = drMDT["OIDITEM"].ToString();
                                string DetailsType = drMDT["DetailsType"].ToString();
                                if (NEGSIZE == MDTSIZE && NEGITEM == MDTITEM && DetailsType == "2")
                                {
                                    string TotalWidth = drMDT["TotalWidth"].ToString();
                                    string UsableWidth = drMDT["UsableWidth"].ToString();
                                    string WeightGM2 = drMDT["WeightGM2"].ToString();
                                    string ActualLengthCm = drMDT["ActualLengthCm"].ToString();
                                    string QtyPcs = drMDT["QtyPcs"].ToString();
                                    string LengthBodyCm = drMDT["LengthBodyCm"].ToString();
                                    string LengthBodyM = drMDT["LengthBodyM"].ToString();
                                    string LengthBodyInc = drMDT["LengthBodyInc"].ToString();
                                    string LengthBodyYrd = drMDT["LengthBodyYrd"].ToString();
                                    string WeightMg = drMDT["WeightMg"].ToString();
                                    string WeightPcs = drMDT["WeightPcs"].ToString();

                                    txtTotal_Negative.Text = TotalWidth;
                                    txtUsable_Negative.Text = UsableWidth;
                                    txtWeight_Negative.Text = WeightGM2;

                                    dtNEG.Rows[rowNEG].SetField("ActualLengthCm", ActualLengthCm);
                                    dtNEG.Rows[rowNEG].SetField("QtyPcs", QtyPcs);
                                    dtNEG.Rows[rowNEG].SetField("LengthBodyCm", LengthBodyCm);
                                    dtNEG.Rows[rowNEG].SetField("LengthBodyM", LengthBodyM);
                                    dtNEG.Rows[rowNEG].SetField("LengthBodyInc", LengthBodyInc);
                                    dtNEG.Rows[rowNEG].SetField("LengthBodyYrd", LengthBodyYrd);
                                    dtNEG.Rows[rowNEG].SetField("WeightMg", WeightMg);
                                    dtNEG.Rows[rowNEG].SetField("WeightPcs", WeightPcs);
                                    break;
                                }

                                rowMDT++;
                            }

                            rowNEG++;
                        }
                        gcNEG.DataSource = dtNEG;
                        gcNEG.Update();
                    }

                }
            }
        }

        public void updateMarkingDetail()
        {

        }

        private void bbiEdit_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {

        }

        private void gvSMPL_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator) e.Info.DisplayText = (e.RowHandle + 1).ToString();
            gvSMPL.IndicatorWidth = 40;
        }

        private void sbRefreshUser_Click(object sender, EventArgs e)
        {
            LoadUserSMPL();
        }

        private void glueUSER_EditValueChanged(object sender, EventArgs e)
        {
            LoadSMPL();
        }

        private void gvSMPL_DoubleClick(object sender, EventArgs e)
        {
            DXMouseEventArgs ea = e as DXMouseEventArgs;
            GridView view = sender as GridView;
            GridHitInfo info = view.CalcHitInfo(ea.Location);
            if (info.InRow || info.InRowCell)
            {
                GridView gv = gvSMPL;
                string SMPLID = gv.GetFocusedRowCellValue("ID").ToString();
                SMPLData(SMPLID);
            }

        }

        private void ClearSMPL()
        {
            //**** Clear Value ******
            lblStatus.Text = "NEW";

            txeMarkingNo.Text = "";
            lblMarkID.Text = "";
            lblID.Text = "";
            txeSampleRequestNo.Text = "";
            txeItemNo.Text = "";
            dteRequestDate.EditValue = null;
            txtModelName.Text = "";
            txeSpecofSize.Text = "";
            txeCategory.Text = "";
            txeUseFor.Text = "";
            txeStyleName.Text = "";
            txeCustomer.Text = "";
            txeRequestBy.Text = "";
            dtDeliveryRequest.EditValue = null;
            slueBranch.EditValue = "";
            txeSeason.Text = "";
            txeSaleSection.Text = "";
            gcQR.DataSource = null;
            //***********************
            ClearVFBCode();
        }

        private void ClearVFBCode()
        {
            lblITEMSEL.Text = "";
            lblOIDITEM.Text = "";
            lblItemCode.Text = "";
            lblItemDescription.Text = "";
            lblParts.Text = "";
            lblOIDSMPLDT.Text = "";
            lblFBPartID.Text = "";

            txePatternNo.Text = "";
            sluePatternSizeZone.EditValue = "";
            txeVendFBCode.Text = "";
            txeColor.Text = "";
            txeSampleLotNo.Text = "";
            txeFBType.Text = "";
           
            gcFBPart.DataSource = null;
            btnInsert.Enabled = false;

            gcSTD.DataSource = null;
            gcNEG.DataSource = null;
            gcPOS.DataSource = null;

            txtTotal_Standard.Text = "0";
            txtUsable_Standard.Text = "0";
            txtWeight_Standard.Text = "0";

            txtTotal_Positive.Text = "0";
            txtUsable_Positive.Text = "0";
            txtWeight_Positive.Text = "0";

            txtTotal_Negative.Text = "0";
            txtUsable_Negative.Text = "0";
            txtWeight_Negative.Text = "0";

            tabMarkDetail.SelectedTabPage = lcgStd;
        }

        private void SMPLData(string SMPLID)
        {
            ClearSMPL();

            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT DISTINCT  ");
            sbSQL.Append("       SRQ.OIDSMPL AS ID, SRQ.SMPLNo AS [SMPL No.], CASE WHEN SRQ.SMPLRevise > 0 THEN CONVERT(VARCHAR, SRQ.SMPLRevise) ELSE '' END AS Revise, CONVERT(VARCHAR(10), SRQ.RequestDate, 103) AS RequestDate, SRQ.ContactName AS RequestBy, CONVERT(VARCHAR(10), SRQ.DeliveryRequest, 103) AS DeliveryRequest, SRQ.ReferenceNo, SRQ.Season, SRQ.SMPLItem, SRQ.ModelName, (CASE SRQ.PatternSizeZone WHEN 0 THEN 'Japan' WHEN 1 THEN 'Europe' WHEN 2 THEN 'US' END) AS PatternSizeZone, SRQ.SMPLPatternNo AS PatternNo, SUF.UseFor, GC.CategoryName, PS.StyleName,  ");
            sbSQL.Append("       SUBSTRING((SELECT GarmentParts + ', ' AS 'data()' FROM(SELECT DISTINCT GP.OIDGParts, GP.GarmentParts FROM SMPLQuantityRequired AS SQR INNER JOIN SMPLRequestFabric AS SFB ON SQR.OIDSMPLDT = SFB.OIDSMPLDT AND SQR.OIDSMPL = SRQ.OIDSMPL  INNER JOIN SMPLRequestFabricParts AS SFBP ON SFB.OIDSMPLFB = SFBP.OIDSMPLFB AND SFB.OIDSMPLDT = SFBP.OIDSMPLDT INNER JOIN GarmentParts AS GP ON SFBP.OIDGParts = GP.OIDGParts) AS FParts FOR XML PATH('')), 1, LEN((SELECT GarmentParts + ', ' AS 'data()' FROM(SELECT DISTINCT GP.OIDGParts, GP.GarmentParts FROM SMPLQuantityRequired AS SQR INNER JOIN SMPLRequestFabric AS SFB ON SQR.OIDSMPLDT = SFB.OIDSMPLDT AND SQR.OIDSMPL = SRQ.OIDSMPL  INNER JOIN SMPLRequestFabricParts AS SFBP ON SFB.OIDSMPLFB = SFBP.OIDSMPLFB AND SFB.OIDSMPLDT = SFBP.OIDSMPLDT INNER JOIN GarmentParts AS GP ON SFBP.OIDGParts = GP.OIDGParts) AS FParts FOR XML PATH(''))) -1) AS FabricParts, ");
            sbSQL.Append("       CUS.Name AS Customer, SRQ.OIDBranch AS Branch, DP.Name AS Department, (SELECT TOP(1) FullName FROM Users WHERE OIDUSER = SRQ.CreatedBy) AS CreatedBy, SRQ.CreatedDate, (SELECT TOP(1) FullName FROM Users WHERE OIDUSER = SRQ.UpdatedBy) AS UpdatedBy, SRQ.UpdatedDate ");
            sbSQL.Append("FROM   SMPLRequest AS SRQ INNER JOIN ");
            sbSQL.Append("       SMPLUseFor AS SUF ON SRQ.UseFor = SUF.OIDUF INNER JOIN ");
            sbSQL.Append("       GarmentCategory AS GC ON SRQ.OIDCATEGORY = GC.OIDGCATEGORY INNER JOIN ");
            sbSQL.Append("       ProductStyle AS PS ON SRQ.OIDSTYLE = PS.OIDSTYLE INNER JOIN ");
            sbSQL.Append("       Customer AS CUS ON SRQ.OIDCUST = CUS.OIDCUST INNER JOIN ");
            sbSQL.Append("       Branchs AS BN ON SRQ.OIDBranch = BN.OIDBranch INNER JOIN ");
            sbSQL.Append("       Departments AS DP ON SRQ.OIDDEPT = DP.OIDDEPT ");
            sbSQL.Append("WHERE (SRQ.OIDSMPL = '" + SMPLID + "')  ");

            DataTable dtSMPL = this.DBC.DBQuery(sbSQL).getDataTable();
            if (dtSMPL != null)
            {
                foreach (DataRow drSMPL in dtSMPL.Rows)
                {
                    lblID.Text = drSMPL["ID"].ToString();
                    txeSampleRequestNo.Text = drSMPL["SMPL No."].ToString();
                    txeItemNo.Text = drSMPL["SMPLItem"].ToString();
                    dteRequestDate.EditValue = null;
                    if (drSMPL["RequestDate"].ToString() != "")
                        dteRequestDate.EditValue = Convert.ToDateTime(drSMPL["RequestDate"].ToString());

                    txtModelName.Text = drSMPL["ModelName"].ToString();
                    txeSpecofSize.Text = drSMPL["PatternSizeZone"].ToString();
                    txeCategory.Text = drSMPL["CategoryName"].ToString();
                    txeUseFor.Text = drSMPL["UseFor"].ToString();
                    txeStyleName.Text = drSMPL["StyleName"].ToString();
                    txeCustomer.Text = drSMPL["Customer"].ToString();
                    txeRequestBy.Text = drSMPL["RequestBy"].ToString();
                    dtDeliveryRequest.EditValue = null;
                    if (drSMPL["DeliveryRequest"].ToString() != "")
                        dtDeliveryRequest.EditValue = Convert.ToDateTime(drSMPL["DeliveryRequest"].ToString());

                    slueBranch.EditValue = drSMPL["Branch"].ToString();
                    txeSeason.Text = drSMPL["Season"].ToString();
                    txeSaleSection.Text = drSMPL["Department"].ToString();
                }

                sbSQL.Clear();
                sbSQL.Append("SELECT SQR.OIDSMPLDT AS ID, PC.ColorName AS Color, PS.SizeName AS Size, SQR.Quantity, UN.UnitName AS Unit ");
                sbSQL.Append("FROM   SMPLQuantityRequired AS SQR LEFT OUTER JOIN ");
                sbSQL.Append("       ProductColor AS PC ON SQR.OIDCOLOR = PC.OIDCOLOR LEFT OUTER JOIN ");
                sbSQL.Append("       ProductSize AS PS ON SQR.OIDSIZE = PS.OIDSIZE LEFT OUTER JOIN ");
                sbSQL.Append("       Unit AS UN ON SQR.OIDUnit = UN.OIDUNIT ");
                sbSQL.Append("WHERE  (SQR.OIDSMPL = '" + SMPLID + "') ");
                sbSQL.Append("ORDER BY ID ");
                new ObjDE.setGridControl(gcQR, gvQR, sbSQL).getData(false, false, true, false);

                sbSQL.Clear();
                sbSQL.Append("SELECT DISTINCT  ");
                sbSQL.Append("       SFB.OIDITEM, IT.Code AS ItemCode, IT.Description AS ItemDescription, ");
                sbSQL.Append("       (SELECT VendFBCode + ', ' AS 'data()' FROM(SELECT DISTINCT VendFBCode FROM SMPLRequestFabric WHERE OIDSMPLDT IN(SELECT OIDSMPLDT FROM SMPLQuantityRequired WHERE OIDSMPL = SRQ.OIDSMPL) AND OIDITEM = SFB.OIDITEM AND VendFBCode IS NOT NULL) AS VendFBCode FOR XML PATH('')) AS VendorFabricCode, ");
                sbSQL.Append("       (SELECT Composition + ', ' AS 'data()' FROM(SELECT DISTINCT Composition FROM SMPLRequestFabric WHERE OIDSMPLDT IN(SELECT OIDSMPLDT FROM SMPLQuantityRequired WHERE OIDSMPL = SRQ.OIDSMPL) AND OIDITEM = SFB.OIDITEM AND Composition IS NOT NULL) AS FBCode FOR XML PATH('')) AS Composition, ");
                sbSQL.Append("       (SELECT FBWeight + ', ' AS 'data()' FROM(SELECT DISTINCT FBWeight FROM SMPLRequestFabric WHERE OIDSMPLDT IN(SELECT OIDSMPLDT FROM SMPLQuantityRequired WHERE OIDSMPL = SRQ.OIDSMPL) AND OIDITEM = SFB.OIDITEM AND FBWeight IS NOT NULL) AS FBCode FOR XML PATH('')) AS Weight, ");
                sbSQL.Append("       (SELECT OIDCOLOR + ', ' AS 'data()' FROM(SELECT DISTINCT CONVERT(VARCHAR, OIDCOLOR) AS OIDCOLOR FROM SMPLRequestFabric WHERE OIDSMPLDT IN(SELECT OIDSMPLDT FROM SMPLQuantityRequired WHERE OIDSMPL = SRQ.OIDSMPL) AND OIDITEM = SFB.OIDITEM AND OIDCOLOR IS NOT NULL) AS OIDCOLOR FOR XML PATH('')) AS OIDCOLOR, ");
                sbSQL.Append("       (SELECT ColorName + ', ' AS 'data()' FROM(SELECT DISTINCT APC.ColorName FROM SMPLRequestFabric AS AFB INNER JOIN ProductColor AS APC ON AFB.OIDCOLOR = APC.OIDCOLOR WHERE AFB.OIDSMPLDT IN(SELECT OIDSMPLDT FROM SMPLQuantityRequired WHERE OIDSMPL = SRQ.OIDSMPL) AND AFB.OIDITEM = SFB.OIDITEM AND AFB.OIDCOLOR IS NOT NULL) AS OIDCOLOR FOR XML PATH('')) AS Color, ");
                sbSQL.Append("       (SELECT OIDVEND + ', ' AS 'data()' FROM(SELECT DISTINCT CONVERT(VARCHAR, OIDVEND) AS OIDVEND FROM SMPLRequestFabric WHERE OIDSMPLDT IN(SELECT OIDSMPLDT FROM SMPLQuantityRequired WHERE OIDSMPL = SRQ.OIDSMPL) AND OIDITEM = SFB.OIDITEM AND OIDVEND IS NOT NULL) AS OIDVEND FOR XML PATH('')) AS OIDVEND, ");
                sbSQL.Append("       (SELECT Name + ', ' AS 'data()' FROM(SELECT DISTINCT BVD.Name FROM SMPLRequestFabric AS BFB INNER JOIN Vendor AS BVD ON BFB.OIDVEND = BVD.OIDVEND WHERE BFB.OIDSMPLDT IN(SELECT OIDSMPLDT FROM SMPLQuantityRequired WHERE OIDSMPL = SRQ.OIDSMPL) AND BFB.OIDITEM = SFB.OIDITEM AND BFB.OIDVEND IS NOT NULL) AS OIDVEND FOR XML PATH('')) AS Vendor, ");
                sbSQL.Append("       (SELECT SMPLotNo + ', ' AS 'data()' FROM(SELECT DISTINCT SMPLotNo FROM SMPLRequestFabric WHERE OIDSMPLDT IN(SELECT OIDSMPLDT FROM SMPLQuantityRequired WHERE OIDSMPL = SRQ.OIDSMPL) AND OIDITEM = SFB.OIDITEM AND SMPLotNo IS NOT NULL) AS FBLot FOR XML PATH('')) AS SMPLotNo, ");
                sbSQL.Append("       (SELECT FBType + ', ' AS 'data()' FROM(SELECT DISTINCT FBType FROM SMPLRequestFabric WHERE OIDSMPLDT IN(SELECT OIDSMPLDT FROM SMPLQuantityRequired WHERE OIDSMPL = SRQ.OIDSMPL) AND OIDITEM = SFB.OIDITEM AND FBType IS NOT NULL) AS FBType FOR XML PATH('')) AS FabricType, ");
                sbSQL.Append("       (SELECT GarmentParts + ', ' AS 'data()' FROM(SELECT DISTINCT GP.OIDGParts, GP.GarmentParts FROM SMPLQuantityRequired AS SQR INNER JOIN SMPLRequestFabric AS FB ON SQR.OIDSMPLDT = FB.OIDSMPLDT AND SQR.OIDSMPL = SRQ.OIDSMPL INNER JOIN SMPLRequestFabricParts AS SFBP ON FB.OIDSMPLFB = SFBP.OIDSMPLFB AND FB.OIDSMPLDT = SFBP.OIDSMPLDT AND FB.OIDITEM = SFB.OIDITEM INNER JOIN GarmentParts AS GP ON SFBP.OIDGParts = GP.OIDGParts AND GP.GarmentParts IS NOT NULL) AS FParts FOR XML PATH('')) AS FabricParts ");
                sbSQL.Append("FROM   SMPLRequestFabric AS SFB INNER JOIN ");
                sbSQL.Append("       Items AS IT ON SFB.OIDITEM = IT.OIDITEM INNER JOIN ");
                sbSQL.Append("       SMPLQuantityRequired AS SQR ON SFB.OIDSMPLDT = SQR.OIDSMPLDT INNER JOIN ");
                sbSQL.Append("       SMPLRequest AS SRQ ON SQR.OIDSMPL = SRQ.OIDSMPL AND SRQ.OIDSMPL = '" + SMPLID + "' ");
                sbSQL.Append("ORDER BY IT.Code ");
                new ObjDE.setGridControl(gcListofFabric, gvListofFabric, sbSQL).getData(false, false, false, true);

                gvListofFabric.Columns["OIDITEM"].Visible = false;
                gvListofFabric.Columns["OIDCOLOR"].Visible = false;
                gvListofFabric.Columns["OIDVEND"].Visible = false;
                gvListofFabric.Columns["Weight"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;

                gvListofFabric.Columns[0].Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;

                DataTable dtFB = (DataTable)gcListofFabric.DataSource;
                if (dtFB != null)
                {
                    int rowIndex = 0;
                    foreach (DataRow drFB in dtFB.Rows)
                    {
                        string VendorFabricCode = drFB["VendorFabricCode"].ToString().Trim();
                        VendorFabricCode = VendorFabricCode.IndexOf(',') > -1 ? VendorFabricCode.Substring(0, VendorFabricCode.Length - 1) : VendorFabricCode;
                        dtFB.Rows[rowIndex].SetField("VendorFabricCode", VendorFabricCode);

                        string Composition = drFB["Composition"].ToString().Trim();
                        Composition = Composition.IndexOf(',') > -1 ? Composition.Substring(0, Composition.Length - 1) : Composition;
                        dtFB.Rows[rowIndex].SetField("Composition", Composition);

                        string Weight = drFB["Weight"].ToString().Trim();
                        Weight = Weight.IndexOf(',') > -1 ? Weight.Substring(0, Weight.Length - 1) : Weight;
                        dtFB.Rows[rowIndex].SetField("Weight", Weight);

                        string OIDCOLOR = drFB["OIDCOLOR"].ToString().Trim();
                        OIDCOLOR = OIDCOLOR.IndexOf(',') > -1 ? OIDCOLOR.Substring(0, OIDCOLOR.Length - 1) : OIDCOLOR;
                        dtFB.Rows[rowIndex].SetField("OIDCOLOR", OIDCOLOR);

                        string Color = drFB["Color"].ToString().Trim();
                        Color = Color.IndexOf(',') > -1 ? Color.Substring(0, Color.Length - 1) : Color;
                        dtFB.Rows[rowIndex].SetField("Color", Color);

                        string OIDVEND = drFB["OIDVEND"].ToString().Trim();
                        OIDVEND = OIDVEND.IndexOf(',') > -1 ? OIDVEND.Substring(0, OIDVEND.Length - 1) : OIDVEND;
                        dtFB.Rows[rowIndex].SetField("OIDVEND", OIDVEND);

                        string Vendor = drFB["Vendor"].ToString().Trim();
                        Vendor = Vendor.IndexOf(',') > -1 ? Vendor.Substring(0, Vendor.Length - 1) : Vendor;
                        dtFB.Rows[rowIndex].SetField("Vendor", Vendor);

                        string SMPLotNo = drFB["SMPLotNo"].ToString().Trim();
                        SMPLotNo = SMPLotNo.IndexOf(',') > -1 ? SMPLotNo.Substring(0, SMPLotNo.Length - 1) : SMPLotNo;
                        dtFB.Rows[rowIndex].SetField("SMPLotNo", SMPLotNo);

                        string FabricType = drFB["FabricType"].ToString().Trim();
                        FabricType = FabricType.IndexOf(',') > -1 ? FabricType.Substring(0, FabricType.Length - 1) : FabricType;
                        dtFB.Rows[rowIndex].SetField("FabricType", FabricType);

                        string FabricParts = drFB["FabricParts"].ToString().Trim();
                        FabricParts = FabricParts.IndexOf(',') > -1 ? FabricParts.Substring(0, FabricParts.Length - 1) : FabricParts;
                        dtFB.Rows[rowIndex].SetField("FabricParts", FabricParts);

                        rowIndex++;
                    }
                    gcListofFabric.DataSource = dtFB;
                }

            }

        }

        private void gvSTD_CustomColumnDisplayText(object sender, DevExpress.XtraGrid.Views.Base.CustomColumnDisplayTextEventArgs e)
        {
            if (e.Column.FieldName != "FabricType" && e.Column.FieldName != "OIDSIZE" && e.Column.FieldName != "SizeName")
            {
                if (Convert.ToString(e.Value) != "")
                {
                    if (Convert.ToDouble(e.Value) == 0)
                    {
                        e.DisplayText = "";
                    }
                }
            }
        }

        private void gvPOS_CustomColumnDisplayText(object sender, DevExpress.XtraGrid.Views.Base.CustomColumnDisplayTextEventArgs e)
        {
            if (e.Column.FieldName != "FabricType" && e.Column.FieldName != "OIDSIZE" && e.Column.FieldName != "SizeName")
            {
                if (Convert.ToString(e.Value) != "")
                {
                    if (Convert.ToDouble(e.Value) == 0)
                    {
                        e.DisplayText = "";
                    }
                }
            }
        }

        private void gvNEG_CustomColumnDisplayText(object sender, DevExpress.XtraGrid.Views.Base.CustomColumnDisplayTextEventArgs e)
        {
            if (e.Column.FieldName != "FabricType" && e.Column.FieldName != "OIDSIZE" && e.Column.FieldName != "SizeName")
            {
                if (Convert.ToDouble(e.Value) == 0)
                {
                    e.DisplayText = "";
                }
            }
        }

        private void gvSTD_CalcRowHeight(object sender, RowHeightEventArgs e)
        {
            DataTable dtR = (DataTable)gcSTD.DataSource;
            if (dtR != null)
            {
                GridViewInfo vi = gvSTD.GetViewInfo() as GridViewInfo;
                int headerHeight = vi.ColumnRowHeight;

                int xRowHeight = 0;

                GridView view = sender as GridView;
                if (view == null) return;
                if (e.RowHandle >= 0)
                {
                    if (dtR.Rows.Count > 3)
                        xRowHeight = (int)Math.Floor(Convert.ToDouble(gcSTD.Height - headerHeight) / Convert.ToDouble(dtR.Rows.Count)) - 2;
                    else
                        xRowHeight = (int)Math.Floor(Convert.ToDouble(gcSTD.Height - headerHeight) / Convert.ToDouble(3)) - 2;
                }
                e.RowHeight = xRowHeight;
                //MessageBox.Show("GHeight:" + gcSTD.Height.ToString() + ", Header:" + headerHeight.ToString() + ", RowHeight:" + xRowHeight.ToString());
            }
        }

        private void gvPOS_CalcRowHeight(object sender, RowHeightEventArgs e)
        {
            DataTable dtR = (DataTable)gcPOS.DataSource;
            if (dtR != null)
            {
                GridViewInfo vi = gvPOS.GetViewInfo() as GridViewInfo;
                int headerHeight = vi.ColumnRowHeight;

                int xRowHeight = 0;

                GridView view = sender as GridView;
                if (view == null) return;
                if (e.RowHandle >= 0)
                {
                    if (dtR.Rows.Count > 3)
                        xRowHeight = (int)Math.Floor(Convert.ToDouble(gcPOS.Height - headerHeight) / Convert.ToDouble(dtR.Rows.Count)) - 2;
                    else
                        xRowHeight = (int)Math.Floor(Convert.ToDouble(gcPOS.Height - headerHeight) / Convert.ToDouble(3)) - 2;
                }
                e.RowHeight = xRowHeight;
                //MessageBox.Show("GHeight:" + gcPOS.Height.ToString() + ", Header:" + headerHeight.ToString() + ", RowHeight:" + xRowHeight.ToString());
            }
        }

        private void gvNEG_CalcRowHeight(object sender, RowHeightEventArgs e)
        {
            DataTable dtR = (DataTable)gcNEG.DataSource;
            if (dtR != null)
            {
                GridViewInfo vi = gvNEG.GetViewInfo() as GridViewInfo;
                int headerHeight = vi.ColumnRowHeight;

                int xRowHeight = 0;

                GridView view = sender as GridView;
                if (view == null) return;
                if (e.RowHandle >= 0)
                {
                    if (dtR.Rows.Count > 3 && dtR.Rows.Count < 7)
                        xRowHeight = (int)Math.Floor(Convert.ToDouble(gcNEG.Height - headerHeight) / Convert.ToDouble(dtR.Rows.Count)) - 2;
                    else
                        xRowHeight = (int)Math.Floor(Convert.ToDouble(gcNEG.Height - headerHeight) / Convert.ToDouble(3)) - 2;
                }
                e.RowHeight = xRowHeight;
                //MessageBox.Show("GHeight:" + gcNEG.Height.ToString() + ", Header:" + headerHeight.ToString() + ", RowHeight:" + xRowHeight.ToString());
            }
        }

        private void gvMDT_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator) e.Info.DisplayText = (e.RowHandle + 1).ToString();
        }

        private void sbClear_Click(object sender, EventArgs e)
        {
            ClearVFBCode();
        }

        private void gvMARK_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator) e.Info.DisplayText = (e.RowHandle + 1).ToString();
        }

        private void gvMARK_SelectionChanged(object sender, DevExpress.Data.SelectionChangedEventArgs e)
        {
            GridView gv = gvMARK;
            int RowSelect = gv.GetFocusedDataSourceRowIndex();

            for (int i = 0; i < gvMARK.DataRowCount; i++)
            {
                //DataRow row = gridView1.GetDataRow(i);
                if (i != RowSelect)
                    gvMARK.UnselectRow(i);
            }

            int[] selectedRowHandles = gvMARK.GetSelectedRows();
            if (selectedRowHandles.Length > 0)
            {
                bbiPrint.Enabled = true;
                bbiPrintPDF.Enabled = true;
                //bbiCLONE.Enabled = true;

                string OIDUSER = gv.GetFocusedRowCellValue("ByCreated").ToString();
                if (UserLogin.OIDUser.ToString() == OIDUSER)
                {
                    bbiUPDATE.Enabled = true;
                    //bbiREVISE.Enabled = true;
                    //bbiDELBILL.Enabled = true;

                    string Status = gv.GetFocusedRowCellValue("Status").ToString();
                    if (Status == "0")
                        bbiDELBILL.Enabled = false;
                    else
                        bbiDELBILL.Enabled = true;
                }
                else
                {
                    bbiUPDATE.Enabled = false;
                    //bbiREVISE.Enabled = false;
                    bbiDELBILL.Enabled = false;
                }
            }
            else
            {
                bbiPrint.Enabled = false;
                bbiPrintPDF.Enabled = false;
                bbiUPDATE.Enabled = false;
                //bbiREVISE.Enabled = false;
                //bbiCLONE.Enabled = false;
                bbiDELBILL.Enabled = false;
            }
        }

        private void rgDocActive_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (rgDocActive.EditValue == null)
                rgDocActive.EditValue = 1;
            if (rgDocUser.EditValue == null)
                rgDocUser.EditValue = 0;

            getGrid_MARK(gcMARK, gvMARK, UserLogin.OIDUser, Convert.ToInt32(rgDocActive.EditValue.ToString()), Convert.ToInt32(rgDocUser.EditValue.ToString()));
            HideSelectDoc();
        }

        private void rgDocUser_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (rgDocActive.EditValue == null)
                rgDocActive.EditValue = 1;
            if (rgDocUser.EditValue == null)
                rgDocUser.EditValue = 0;

            getGrid_MARK(gcMARK, gvMARK, UserLogin.OIDUser, Convert.ToInt32(rgDocActive.EditValue.ToString()), Convert.ToInt32(rgDocUser.EditValue.ToString()));
            HideSelectDoc();
        }

        private void DEV02_Shown(object sender, EventArgs e)
        {
            if (this.DBC.chkCONNECTION_STING() == false)
            {
                FUNC.msgError("Connection string is null.");
                return;
            }

            rgDocActive.EditValue = 1;
            rgDocUser.EditValue = 0;

            tabMARKING.SelectedTabPage = lcgList;
        }

        private void bbiUPDATE_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                string MarkNo = "";
                if (tabMARKING.SelectedTabPage == lcgList) //LIST
                {
                    GridView gv = gvMARK;
                    MarkNo = gv.GetFocusedRowCellValue("Marking No.").ToString();
                }
                else
                {
                    MarkNo = txeMarkingNo.Text.Trim();
                }

                LoadMarkingDocument(MarkNo, "UPDATE");
                lblStatus.Text = "UPDATE";
                tabMARKING.SelectedTabPage = lcgMark;
                //SetWrite();
                txeMarkingNo.Focus();
            }
            catch (Exception exc)
            {
                FUNC.msgError(exc.ToString());
            }
        }

        private void LoadMarkingDocument(string MARKNo, string MARKModify)
        {
            LoadNewData(MARKModify);
            MARKNo = MARKNo.ToUpper().Trim();
            if (MARKNo != "")
            {
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT DISTINCT  ");
                sbSQL.Append("       MK.OIDMARK, MK.MarkingNo, SRQ.OIDSMPL AS ID, SRQ.SMPLNo AS [SMPL No.], CASE WHEN SRQ.SMPLRevise > 0 THEN CONVERT(VARCHAR, SRQ.SMPLRevise) ELSE '' END AS Revise, CONVERT(VARCHAR(10), SRQ.RequestDate, 103) AS RequestDate, SRQ.ContactName AS RequestBy, CONVERT(VARCHAR(10), SRQ.DeliveryRequest, 103) AS DeliveryRequest, SRQ.ReferenceNo, SRQ.Season, SRQ.SMPLItem, SRQ.ModelName, (CASE SRQ.PatternSizeZone WHEN 0 THEN 'Japan' WHEN 1 THEN 'Europe' WHEN 2 THEN 'US' END) AS PatternSizeZone, SRQ.SMPLPatternNo AS PatternNo, SUF.UseFor, GC.CategoryName, PS.StyleName,  ");
                sbSQL.Append("       SUBSTRING((SELECT GarmentParts + ', ' AS 'data()' FROM(SELECT DISTINCT GP.OIDGParts, GP.GarmentParts FROM SMPLQuantityRequired AS SQR INNER JOIN SMPLRequestFabric AS SFB ON SQR.OIDSMPLDT = SFB.OIDSMPLDT AND SQR.OIDSMPL = SRQ.OIDSMPL  INNER JOIN SMPLRequestFabricParts AS SFBP ON SFB.OIDSMPLFB = SFBP.OIDSMPLFB AND SFB.OIDSMPLDT = SFBP.OIDSMPLDT INNER JOIN GarmentParts AS GP ON SFBP.OIDGParts = GP.OIDGParts) AS FParts FOR XML PATH('')), 1, LEN((SELECT GarmentParts + ', ' AS 'data()' FROM(SELECT DISTINCT GP.OIDGParts, GP.GarmentParts FROM SMPLQuantityRequired AS SQR INNER JOIN SMPLRequestFabric AS SFB ON SQR.OIDSMPLDT = SFB.OIDSMPLDT AND SQR.OIDSMPL = SRQ.OIDSMPL  INNER JOIN SMPLRequestFabricParts AS SFBP ON SFB.OIDSMPLFB = SFBP.OIDSMPLFB AND SFB.OIDSMPLDT = SFBP.OIDSMPLDT INNER JOIN GarmentParts AS GP ON SFBP.OIDGParts = GP.OIDGParts) AS FParts FOR XML PATH(''))) -1) AS FabricParts, ");
                sbSQL.Append("       CUS.Name AS Customer, SRQ.OIDBranch AS Branch, DP.Name AS Department, (SELECT TOP(1) FullName FROM Users WHERE OIDUSER = SRQ.CreatedBy) AS CreatedBy, SRQ.CreatedDate, (SELECT TOP(1) FullName FROM Users WHERE OIDUSER = SRQ.UpdatedBy) AS UpdatedBy, SRQ.UpdatedDate ");
                sbSQL.Append("FROM   Marking AS MK INNER JOIN ");
                sbSQL.Append("       SMPLRequest AS SRQ ON MK.OIDSMPL = SRQ.OIDSMPL INNER JOIN ");
                sbSQL.Append("       SMPLUseFor AS SUF ON SRQ.UseFor = SUF.OIDUF INNER JOIN ");
                sbSQL.Append("       GarmentCategory AS GC ON SRQ.OIDCATEGORY = GC.OIDGCATEGORY INNER JOIN ");
                sbSQL.Append("       ProductStyle AS PS ON SRQ.OIDSTYLE = PS.OIDSTYLE INNER JOIN ");
                sbSQL.Append("       Customer AS CUS ON SRQ.OIDCUST = CUS.OIDCUST INNER JOIN ");
                sbSQL.Append("       Branchs AS BN ON SRQ.OIDBranch = BN.OIDBranch INNER JOIN ");
                sbSQL.Append("       Departments AS DP ON SRQ.OIDDEPT = DP.OIDDEPT ");
                sbSQL.Append("WHERE (MK.MarkingNo = N'" + MARKNo + "')  ");

                DataTable dtSMPL = this.DBC.DBQuery(sbSQL).getDataTable();
                if (dtSMPL != null)
                {
                    lblStatus.Text = MARKModify;
                    foreach (DataRow drSMPL in dtSMPL.Rows)
                    {
                        lblMarkID.Text = drSMPL["OIDMARK"].ToString();
                        txeMarkingNo.Text = drSMPL["MarkingNo"].ToString();
                        lblID.Text = drSMPL["ID"].ToString();
                        txeSampleRequestNo.Text = drSMPL["SMPL No."].ToString();
                        txeItemNo.Text = drSMPL["SMPLItem"].ToString();
                        dteRequestDate.EditValue = null;
                        if (drSMPL["RequestDate"].ToString() != "")
                            dteRequestDate.EditValue = Convert.ToDateTime(drSMPL["RequestDate"].ToString());

                        txtModelName.Text = drSMPL["ModelName"].ToString();
                        txeSpecofSize.Text = drSMPL["PatternSizeZone"].ToString();
                        txeCategory.Text = drSMPL["CategoryName"].ToString();
                        txeUseFor.Text = drSMPL["UseFor"].ToString();
                        txeStyleName.Text = drSMPL["StyleName"].ToString();
                        txeCustomer.Text = drSMPL["Customer"].ToString();
                        txeRequestBy.Text = drSMPL["RequestBy"].ToString();
                        dtDeliveryRequest.EditValue = null;
                        if (drSMPL["DeliveryRequest"].ToString() != "")
                            dtDeliveryRequest.EditValue = Convert.ToDateTime(drSMPL["DeliveryRequest"].ToString());

                        slueBranch.EditValue = drSMPL["Branch"].ToString();
                        txeSeason.Text = drSMPL["Season"].ToString();
                        txeSaleSection.Text = drSMPL["Department"].ToString();
                    }

                    sbSQL.Clear();
                    sbSQL.Append("SELECT SQR.OIDSMPLDT AS ID, PC.ColorName AS Color, PS.SizeName AS Size, SQR.Quantity, UN.UnitName AS Unit ");
                    sbSQL.Append("FROM   SMPLQuantityRequired AS SQR LEFT OUTER JOIN ");
                    sbSQL.Append("       ProductColor AS PC ON SQR.OIDCOLOR = PC.OIDCOLOR LEFT OUTER JOIN ");
                    sbSQL.Append("       ProductSize AS PS ON SQR.OIDSIZE = PS.OIDSIZE LEFT OUTER JOIN ");
                    sbSQL.Append("       Unit AS UN ON SQR.OIDUnit = UN.OIDUNIT ");
                    sbSQL.Append("WHERE  (SQR.OIDSMPL = '" + lblID.Text + "') ");
                    sbSQL.Append("ORDER BY ID ");
                    new ObjDE.setGridControl(gcQR, gvQR, sbSQL).getData(false, false, true, false);

                    sbSQL.Clear();
                    sbSQL.Append("SELECT DISTINCT  ");
                    sbSQL.Append("       SFB.OIDITEM, IT.Code AS ItemCode, IT.Description AS ItemDescription, ");
                    sbSQL.Append("       (SELECT VendFBCode + ', ' AS 'data()' FROM(SELECT DISTINCT VendFBCode FROM SMPLRequestFabric WHERE OIDSMPLDT IN(SELECT OIDSMPLDT FROM SMPLQuantityRequired WHERE OIDSMPL = SRQ.OIDSMPL) AND OIDITEM = SFB.OIDITEM AND VendFBCode IS NOT NULL) AS VendFBCode FOR XML PATH('')) AS VendorFabricCode, ");
                    sbSQL.Append("       (SELECT Composition + ', ' AS 'data()' FROM(SELECT DISTINCT Composition FROM SMPLRequestFabric WHERE OIDSMPLDT IN(SELECT OIDSMPLDT FROM SMPLQuantityRequired WHERE OIDSMPL = SRQ.OIDSMPL) AND OIDITEM = SFB.OIDITEM AND Composition IS NOT NULL) AS FBCode FOR XML PATH('')) AS Composition, ");
                    sbSQL.Append("       (SELECT FBWeight + ', ' AS 'data()' FROM(SELECT DISTINCT FBWeight FROM SMPLRequestFabric WHERE OIDSMPLDT IN(SELECT OIDSMPLDT FROM SMPLQuantityRequired WHERE OIDSMPL = SRQ.OIDSMPL) AND OIDITEM = SFB.OIDITEM AND FBWeight IS NOT NULL) AS FBCode FOR XML PATH('')) AS Weight, ");
                    sbSQL.Append("       (SELECT OIDCOLOR + ', ' AS 'data()' FROM(SELECT DISTINCT CONVERT(VARCHAR, OIDCOLOR) AS OIDCOLOR FROM SMPLRequestFabric WHERE OIDSMPLDT IN(SELECT OIDSMPLDT FROM SMPLQuantityRequired WHERE OIDSMPL = SRQ.OIDSMPL) AND OIDITEM = SFB.OIDITEM AND OIDCOLOR IS NOT NULL) AS OIDCOLOR FOR XML PATH('')) AS OIDCOLOR, ");
                    sbSQL.Append("       (SELECT ColorName + ', ' AS 'data()' FROM(SELECT DISTINCT APC.ColorName FROM SMPLRequestFabric AS AFB INNER JOIN ProductColor AS APC ON AFB.OIDCOLOR = APC.OIDCOLOR WHERE AFB.OIDSMPLDT IN(SELECT OIDSMPLDT FROM SMPLQuantityRequired WHERE OIDSMPL = SRQ.OIDSMPL) AND AFB.OIDITEM = SFB.OIDITEM AND AFB.OIDCOLOR IS NOT NULL) AS OIDCOLOR FOR XML PATH('')) AS Color, ");
                    sbSQL.Append("       (SELECT OIDVEND + ', ' AS 'data()' FROM(SELECT DISTINCT CONVERT(VARCHAR, OIDVEND) AS OIDVEND FROM SMPLRequestFabric WHERE OIDSMPLDT IN(SELECT OIDSMPLDT FROM SMPLQuantityRequired WHERE OIDSMPL = SRQ.OIDSMPL) AND OIDITEM = SFB.OIDITEM AND OIDVEND IS NOT NULL) AS OIDVEND FOR XML PATH('')) AS OIDVEND, ");
                    sbSQL.Append("       (SELECT Name + ', ' AS 'data()' FROM(SELECT DISTINCT BVD.Name FROM SMPLRequestFabric AS BFB INNER JOIN Vendor AS BVD ON BFB.OIDVEND = BVD.OIDVEND WHERE BFB.OIDSMPLDT IN(SELECT OIDSMPLDT FROM SMPLQuantityRequired WHERE OIDSMPL = SRQ.OIDSMPL) AND BFB.OIDITEM = SFB.OIDITEM AND BFB.OIDVEND IS NOT NULL) AS OIDVEND FOR XML PATH('')) AS Vendor, ");
                    sbSQL.Append("       (SELECT SMPLotNo + ', ' AS 'data()' FROM(SELECT DISTINCT SMPLotNo FROM SMPLRequestFabric WHERE OIDSMPLDT IN(SELECT OIDSMPLDT FROM SMPLQuantityRequired WHERE OIDSMPL = SRQ.OIDSMPL) AND OIDITEM = SFB.OIDITEM AND SMPLotNo IS NOT NULL) AS FBLot FOR XML PATH('')) AS SMPLotNo, ");
                    sbSQL.Append("       (SELECT FBType + ', ' AS 'data()' FROM(SELECT DISTINCT FBType FROM SMPLRequestFabric WHERE OIDSMPLDT IN(SELECT OIDSMPLDT FROM SMPLQuantityRequired WHERE OIDSMPL = SRQ.OIDSMPL) AND OIDITEM = SFB.OIDITEM AND FBType IS NOT NULL) AS FBType FOR XML PATH('')) AS FabricType, ");
                    sbSQL.Append("       (SELECT GarmentParts + ', ' AS 'data()' FROM(SELECT DISTINCT GP.OIDGParts, GP.GarmentParts FROM SMPLQuantityRequired AS SQR INNER JOIN SMPLRequestFabric AS FB ON SQR.OIDSMPLDT = FB.OIDSMPLDT AND SQR.OIDSMPL = SRQ.OIDSMPL INNER JOIN SMPLRequestFabricParts AS SFBP ON FB.OIDSMPLFB = SFBP.OIDSMPLFB AND FB.OIDSMPLDT = SFBP.OIDSMPLDT AND FB.OIDITEM = SFB.OIDITEM INNER JOIN GarmentParts AS GP ON SFBP.OIDGParts = GP.OIDGParts AND GP.GarmentParts IS NOT NULL) AS FParts FOR XML PATH('')) AS FabricParts ");
                    sbSQL.Append("FROM   SMPLRequestFabric AS SFB INNER JOIN ");
                    sbSQL.Append("       Items AS IT ON SFB.OIDITEM = IT.OIDITEM INNER JOIN ");
                    sbSQL.Append("       SMPLQuantityRequired AS SQR ON SFB.OIDSMPLDT = SQR.OIDSMPLDT INNER JOIN ");
                    sbSQL.Append("       SMPLRequest AS SRQ ON SQR.OIDSMPL = SRQ.OIDSMPL AND SRQ.OIDSMPL = '" + lblID.Text + "' ");
                    sbSQL.Append("ORDER BY IT.Code ");
                    new ObjDE.setGridControl(gcListofFabric, gvListofFabric, sbSQL).getData(false, false, false, true);

                    gvListofFabric.Columns["OIDITEM"].Visible = false;
                    gvListofFabric.Columns["OIDCOLOR"].Visible = false;
                    gvListofFabric.Columns["OIDVEND"].Visible = false;
                    gvListofFabric.Columns["Weight"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;

                    gvListofFabric.Columns[0].Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;

                    DataTable dtFB = (DataTable)gcListofFabric.DataSource;
                    if (dtFB != null)
                    {
                        int rowIndex = 0;
                        foreach (DataRow drFB in dtFB.Rows)
                        {
                            string VendorFabricCode = drFB["VendorFabricCode"].ToString().Trim();
                            VendorFabricCode = VendorFabricCode.IndexOf(',') > -1 ? VendorFabricCode.Substring(0, VendorFabricCode.Length - 1) : VendorFabricCode;
                            dtFB.Rows[rowIndex].SetField("VendorFabricCode", VendorFabricCode);

                            string Composition = drFB["Composition"].ToString().Trim();
                            Composition = Composition.IndexOf(',') > -1 ? Composition.Substring(0, Composition.Length - 1) : Composition;
                            dtFB.Rows[rowIndex].SetField("Composition", Composition);

                            string Weight = drFB["Weight"].ToString().Trim();
                            Weight = Weight.IndexOf(',') > -1 ? Weight.Substring(0, Weight.Length - 1) : Weight;
                            dtFB.Rows[rowIndex].SetField("Weight", Weight);

                            string OIDCOLOR = drFB["OIDCOLOR"].ToString().Trim();
                            OIDCOLOR = OIDCOLOR.IndexOf(',') > -1 ? OIDCOLOR.Substring(0, OIDCOLOR.Length - 1) : OIDCOLOR;
                            dtFB.Rows[rowIndex].SetField("OIDCOLOR", OIDCOLOR);

                            string Color = drFB["Color"].ToString().Trim();
                            Color = Color.IndexOf(',') > -1 ? Color.Substring(0, Color.Length - 1) : Color;
                            dtFB.Rows[rowIndex].SetField("Color", Color);

                            string OIDVEND = drFB["OIDVEND"].ToString().Trim();
                            OIDVEND = OIDVEND.IndexOf(',') > -1 ? OIDVEND.Substring(0, OIDVEND.Length - 1) : OIDVEND;
                            dtFB.Rows[rowIndex].SetField("OIDVEND", OIDVEND);

                            string Vendor = drFB["Vendor"].ToString().Trim();
                            Vendor = Vendor.IndexOf(',') > -1 ? Vendor.Substring(0, Vendor.Length - 1) : Vendor;
                            dtFB.Rows[rowIndex].SetField("Vendor", Vendor);

                            string SMPLotNo = drFB["SMPLotNo"].ToString().Trim();
                            SMPLotNo = SMPLotNo.IndexOf(',') > -1 ? SMPLotNo.Substring(0, SMPLotNo.Length - 1) : SMPLotNo;
                            dtFB.Rows[rowIndex].SetField("SMPLotNo", SMPLotNo);

                            string FabricType = drFB["FabricType"].ToString().Trim();
                            FabricType = FabricType.IndexOf(',') > -1 ? FabricType.Substring(0, FabricType.Length - 1) : FabricType;
                            dtFB.Rows[rowIndex].SetField("FabricType", FabricType);

                            string FabricParts = drFB["FabricParts"].ToString().Trim();
                            FabricParts = FabricParts.IndexOf(',') > -1 ? FabricParts.Substring(0, FabricParts.Length - 1) : FabricParts;
                            dtFB.Rows[rowIndex].SetField("FabricParts", FabricParts);

                            rowIndex++;
                        }
                        gcListofFabric.DataSource = dtFB;
                    }

                    sbSQL.Clear();
                    sbSQL.Append("SELECT    MKD.OIDMARKDT AS ID, MKD.OIDITEM, IT.Code AS ItemCode, IT.Description AS ItemDescription, SRQ.SMPLPatternNo, MKD.GPartsStuff AS FBPartsID, ");
                    sbSQL.Append("          (SELECT GarmentParts + ', ' AS 'data()' FROM(SELECT DISTINCT GP.OIDGParts, GP.GarmentParts FROM SMPLQuantityRequired AS SQR INNER JOIN SMPLRequestFabric AS FB ON SQR.OIDSMPLDT = FB.OIDSMPLDT AND SQR.OIDSMPL = SRQ.OIDSMPL INNER JOIN SMPLRequestFabricParts AS SFBP ON FB.OIDSMPLFB = SFBP.OIDSMPLFB AND FB.OIDSMPLDT = SFBP.OIDSMPLDT AND FB.OIDITEM = MKD.OIDITEM INNER JOIN GarmentParts AS GP ON SFBP.OIDGParts = GP.OIDGParts AND GP.GarmentParts IS NOT NULL) AS FBParts FOR XML PATH('')) AS FBParts, ");
                    sbSQL.Append("          (SELECT VendFBCode + ', ' AS 'data()' FROM(SELECT DISTINCT VendFBCode FROM SMPLRequestFabric WHERE OIDSMPLDT IN(SELECT OIDSMPLDT FROM SMPLQuantityRequired  WHERE OIDSMPL = MK.OIDSMPL) AND OIDITEM = MKD.OIDITEM AND OIDSMPLDT IN(SELECT value FROM MarkingDetails CROSS APPLY string_split(OIDSMPLDTStuff, ',')  WHERE OIDMARKDT = MKD.OIDMARKDT) AND VendFBCode IS NOT NULL) AS VendFBCode FOR XML PATH('')) AS VendorFBCode, ");
                    sbSQL.Append("          (SELECT SMPLotNo + ', ' AS 'data()' FROM(SELECT DISTINCT SMPLotNo FROM SMPLRequestFabric WHERE OIDSMPLDT IN(SELECT OIDSMPLDT FROM SMPLQuantityRequired WHERE OIDSMPL = MK.OIDSMPL) AND OIDITEM = MKD.OIDITEM AND SMPLotNo IS NOT NULL) AS FBLot FOR XML PATH('')) AS SampleLotNo, ");
                    sbSQL.Append("          MKD.DetailsType, FORMAT(MKD.TotalWidthSTD, '###0.####') AS TotalWidth, FORMAT(MKD.UsableWidth, '###0.####') AS UsableWidth, FORMAT(MKD.GM2, '###0.####') AS WeightGM2, MKD.OIDSIZE, PS.SizeName, FORMAT(MKD.PracticalLengthCM, '###0.####') AS ActualLengthCm, MKD.QuantityPCS AS QtyPcs, FORMAT(MKD.LengthPer1CM, '###0.####') AS LengthBodyCm, FORMAT(MKD.LengthPer1M, '###0.####') AS LengthBodyM, FORMAT(MKD.LengthPer1INCH, '###0.####') AS LengthBodyInc, FORMAT(MKD.LengthPer1YARD, '###0.####') AS LengthBodyYrd,  ");
                    sbSQL.Append("          FORMAT(MKD.WeightG, '###0.####') AS WeightMg, FORMAT(MKD.WeightKG, '###0.####') AS WeightPcs, MKD.OIDSMPLDTStuff AS OIDSMPLDT, MKD.OIDSIZEZONE AS PatternSizeZone ");
                    sbSQL.Append("FROM      MarkingDetails AS MKD INNER JOIN ");
                    sbSQL.Append("          Items AS IT ON MKD.OIDITEM = IT.OIDITEM INNER JOIN ");
                    sbSQL.Append("          Marking AS MK ON MKD.OIDMARK = MK.OIDMARK INNER JOIN ");
                    sbSQL.Append("          SMPLRequest AS SRQ ON MK.OIDSMPL = SRQ.OIDSMPL INNER JOIN ");
                    sbSQL.Append("          ProductSize AS PS ON MKD.OIDSIZE = PS.OIDSIZE ");
                    sbSQL.Append("WHERE (MKD.OIDMARK = '" + lblMarkID.Text + "') ");
                    sbSQL.Append("ORDER BY MKD.OIDITEM, MKD.DetailsType ");
                    DataTable dtMK = this.DBC.DBQuery(sbSQL).getDataTable();
                    if (dtMK != null)
                    {
                        int rowIndex = 0;
                        foreach (DataRow drMK in dtMK.Rows)
                        {
                            string FBParts = drMK["FBParts"].ToString().Trim();
                            FBParts = FBParts.IndexOf(',') > -1 ? FBParts.Substring(0, FBParts.Length - 1) : FBParts;
                            dtMK.Rows[rowIndex].SetField("FBParts", FBParts);

                            string VendorFBCode = drMK["VendorFBCode"].ToString().Trim();
                            VendorFBCode = VendorFBCode.IndexOf(',') > -1 ? VendorFBCode.Substring(0, VendorFBCode.Length - 1) : VendorFBCode;
                            dtMK.Rows[rowIndex].SetField("VendorFBCode", VendorFBCode);

                            string SampleLotNo = drMK["SampleLotNo"].ToString().Trim();
                            SampleLotNo = SampleLotNo.IndexOf(',') > -1 ? SampleLotNo.Substring(0, SampleLotNo.Length - 1) : SampleLotNo;
                            dtMK.Rows[rowIndex].SetField("SampleLotNo", SampleLotNo);

                            rowIndex++;
                        }
                        gcMDT.DataSource = dtMK;
                    }

                }
                else
                {
                    lblStatus.Text = "NEW";
                }
            }

            if (MARKModify == "READ-ONLY")
            {
                SetReadOnly();
            }
            else
            {
                SetWrite();
            }
        }

        private void LoadNewData(string Status = "NEW")
        {
            LoadUserSMPL();
            SetWrite();

            if (rgDocActive.EditValue == null)
                rgDocActive.EditValue = 1;
            if (rgDocUser.EditValue == null)
                rgDocUser.EditValue = 0;

            //Tab : Main Load
            getGrid_MARK(gcMARK, gvMARK, UserLogin.OIDUser, Convert.ToInt32(rgDocActive.EditValue.ToString()), Convert.ToInt32(rgDocUser.EditValue.ToString()));
            HideSelectDoc();

            txeMarkingNo.Text = "";
            lblID.Text = "";

            lblStatus.Text = Status;

            NewData();

        }

        private void SetReadOnly()
        {
            //Marking Tab
            dtDocumentDate.ReadOnly = true;
            dtDocumentDate.BackColor = Color.White;
            dtDocumentDate.ForeColor = Color.Black;

            glueMarkingRequestType.ReadOnly = true;
            glueMarkingRequestType.BackColor = Color.White;
            glueMarkingRequestType.ForeColor = Color.Black;

            glueCuttingFac.ReadOnly = true;
            glueCuttingFac.BackColor = Color.White;
            glueCuttingFac.ForeColor = Color.Black;

            glueSewingFac.ReadOnly = true;
            glueSewingFac.BackColor = Color.White;
            glueSewingFac.ForeColor = Color.Black;

            mmRemark.ReadOnly = true;
            mmRemark.BackColor = Color.White;
            mmRemark.ForeColor = Color.Black;


            //Marking Detail Tab
            gvSTD.OptionsBehavior.Editable = false;
            gvPOS.OptionsBehavior.Editable = false;
            gvNEG.OptionsBehavior.Editable = false;

            txtTotal_Standard.ReadOnly = true;
            txtTotal_Standard.BackColor = Color.White;
            txtTotal_Standard.ForeColor = Color.Black;

            txtUsable_Standard.ReadOnly = true;
            txtUsable_Standard.BackColor = Color.White;
            txtUsable_Standard.ForeColor = Color.Black;

            txtWeight_Standard.ReadOnly = true;
            txtWeight_Standard.BackColor = Color.White;
            txtWeight_Standard.ForeColor = Color.Black;

            txtTotal_Positive.ReadOnly = true;
            txtTotal_Positive.BackColor = Color.White;
            txtTotal_Positive.ForeColor = Color.Black;

            txtUsable_Positive.ReadOnly = true;
            txtUsable_Positive.BackColor = Color.White;
            txtUsable_Positive.ForeColor = Color.Black;

            txtWeight_Positive.ReadOnly = true;
            txtWeight_Positive.BackColor = Color.White;
            txtWeight_Positive.ForeColor = Color.Black;

            txtTotal_Negative.ReadOnly = true;
            txtTotal_Negative.BackColor = Color.White;
            txtTotal_Negative.ForeColor = Color.Black;

            txtUsable_Negative.ReadOnly = true;
            txtUsable_Negative.BackColor = Color.White;
            txtUsable_Negative.ForeColor = Color.Black;

            txtWeight_Negative.ReadOnly = true;
            txtWeight_Negative.BackColor = Color.White;
            txtWeight_Negative.ForeColor = Color.Black;

            btnInsert.Enabled = false;
            sbClear.Enabled = false;
        }

        private void SetWrite()
        {
            //Marking Tab
            dtDocumentDate.ReadOnly = false;
            glueMarkingRequestType.ReadOnly = false;
            glueCuttingFac.ReadOnly = false;
            glueSewingFac.ReadOnly = false;
            mmRemark.ReadOnly = false;

            //Marking Detail Tab
            gvSTD.OptionsBehavior.Editable = true;
            gvPOS.OptionsBehavior.Editable = true;
            gvNEG.OptionsBehavior.Editable = true;

            txtTotal_Standard.ReadOnly = false;
            txtUsable_Standard.ReadOnly = false;
            txtWeight_Standard.ReadOnly = false;
            txtTotal_Positive.ReadOnly = false;
            txtUsable_Positive.ReadOnly = false;
            txtWeight_Positive.ReadOnly = false;
            txtTotal_Negative.ReadOnly = false;
            txtUsable_Negative.ReadOnly = false;
            txtWeight_Negative.ReadOnly = false;

            btnInsert.Enabled = true;
            sbClear.Enabled = true;
        }

        private void lblStatus_TextChanged(object sender, EventArgs e)
        {
            lblStatus.ForeColor = Color.White;
            bbiSave.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
            if (lblStatus.Text.Trim() == "NEW")
            {
                lblStatus.BackColor = Color.Green;
            }
            else if (lblStatus.Text.Trim() == "UPDATE")
            {
                lblStatus.BackColor = Color.Navy;
            }
            else if (lblStatus.Text.Trim() == "REVISE")
            {
                lblStatus.BackColor = Color.Purple;
            }
            else if (lblStatus.Text.Trim() == "CLONE")
            {
                lblStatus.BackColor = Color.Teal;
            }
            else if (lblStatus.Text.Trim() == "READ-ONLY")
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

            if (tabMARKING.SelectedTabPage == lcgList) //LIST
                if (chkReadWrite == 1)
                    rpgManage.Visible = true;
                else
                {
                    if (lblStatus.Text.Trim() == "READ-ONLY")
                        if (chkReadWrite == 1)
                            rpgManage.Visible = true;
                        else
                            rpgManage.Visible = false;
                }
        }

        private void gvMDT_CustomColumnDisplayText(object sender, DevExpress.XtraGrid.Views.Base.CustomColumnDisplayTextEventArgs e)
        {
            if (e.Column.FieldName == "TotalWidth" || e.Column.FieldName == "UsableWidth" || 
                e.Column.FieldName == "WeightGM2" || e.Column.FieldName == "ActualLengthCm" || 
                e.Column.FieldName == "QtyPcs" || e.Column.FieldName == "LengthBodyCm" || 
                e.Column.FieldName == "LengthBodyM" || e.Column.FieldName == "LengthBodyInc" || 
                e.Column.FieldName == "LengthBodyYrd" || e.Column.FieldName == "WeightMg" || e.Column.FieldName == "WeightPcs")
            {
                if (Convert.ToString(e.Value) != "")
                {
                    if (Convert.ToDouble(e.Value) == 0)
                    {
                        e.DisplayText = "";
                    }
                }
            }
        }

        private void gvSTD_ValidateRow(object sender, ValidateRowEventArgs e)
        {

        }

        private void gvSTD_CellValueChanged(object sender, CellValueChangedEventArgs e)
        {

        }

        private void txtTotal_Standard_CustomDisplayText(object sender, DevExpress.XtraEditors.Controls.CustomDisplayTextEventArgs e)
        {
            if (e.Value != null)
                if(e.Value.ToString() != "")
                    if (Convert.ToDouble(e.Value) == 0)
                        e.DisplayText = "";
        }

        private void txtUsable_Standard_CustomDisplayText(object sender, DevExpress.XtraEditors.Controls.CustomDisplayTextEventArgs e)
        {
            if (e.Value != null)
                if (e.Value.ToString() != "")
                    if (Convert.ToDouble(e.Value) == 0)
                        e.DisplayText = "";
        }

        private void txtWeight_Standard_CustomDisplayText(object sender, DevExpress.XtraEditors.Controls.CustomDisplayTextEventArgs e)
        {
            if (e.Value != null)
                if (e.Value.ToString() != "")
                    if (Convert.ToDouble(e.Value) == 0)
                        e.DisplayText = "";
        }

        private void txtTotal_Positive_CustomDisplayText(object sender, DevExpress.XtraEditors.Controls.CustomDisplayTextEventArgs e)
        {
            if (e.Value != null)
                if (e.Value.ToString() != "")
                    if (Convert.ToDouble(e.Value) == 0)
                        e.DisplayText = "";
        }

        private void txtUsable_Positive_CustomDisplayText(object sender, DevExpress.XtraEditors.Controls.CustomDisplayTextEventArgs e)
        {
            if (e.Value != null)
                if (e.Value.ToString() != "")
                    if (Convert.ToDouble(e.Value) == 0)
                        e.DisplayText = "";
        }

        private void txtWeight_Positive_CustomDisplayText(object sender, DevExpress.XtraEditors.Controls.CustomDisplayTextEventArgs e)
        {
            if (e.Value != null)
                if (e.Value.ToString() != "")
                    if (Convert.ToDouble(e.Value) == 0)
                        e.DisplayText = "";
        }

        private void txtTotal_Negative_CustomDisplayText(object sender, DevExpress.XtraEditors.Controls.CustomDisplayTextEventArgs e)
        {
            if (e.Value != null)
                if (e.Value.ToString() != "")
                    if (Convert.ToDouble(e.Value) == 0)
                        e.DisplayText = "";
        }

        private void txtUsable_Negative_CustomDisplayText(object sender, DevExpress.XtraEditors.Controls.CustomDisplayTextEventArgs e)
        {
            if (e.Value != null)
                if (e.Value.ToString() != "")
                    if (Convert.ToDouble(e.Value) == 0)
                        e.DisplayText = "";
        }

        private void txtWeight_Negative_CustomDisplayText(object sender, DevExpress.XtraEditors.Controls.CustomDisplayTextEventArgs e)
        {
            if (e.Value != null)
                if (e.Value.ToString() != "")
                    if (Convert.ToDouble(e.Value) == 0)
                        e.DisplayText = "";
        }

        private void txtTotal_Standard_Click(object sender, EventArgs e)
        {
            txtTotal_Standard.SelectAll();
        }

        private void txtUsable_Standard_Click(object sender, EventArgs e)
        {
            txtUsable_Standard.SelectAll();
        }

        private void txtWeight_Standard_Click(object sender, EventArgs e)
        {
            txtWeight_Standard.SelectAll();
        }

        private void txtTotal_Positive_Click(object sender, EventArgs e)
        {
            txtTotal_Positive.SelectAll();
        }

        private void txtUsable_Positive_Click(object sender, EventArgs e)
        {
            txtUsable_Positive.SelectAll();
        }

        private void txtWeight_Positive_Click(object sender, EventArgs e)
        {
            txtWeight_Positive.SelectAll();
        }

        private void txtTotal_Negative_Click(object sender, EventArgs e)
        {
            txtTotal_Negative.SelectAll();
        }

        private void txtUsable_Negative_Click(object sender, EventArgs e)
        {
            txtUsable_Negative.SelectAll();
        }

        private void txtWeight_Negative_Click(object sender, EventArgs e)
        {
            txtWeight_Negative.SelectAll();
        }

        private void txtTotal_Standard_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtUsable_Standard.Focus();
                txtUsable_Standard.SelectAll();
            }
        }

        private void txtUsable_Standard_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtWeight_Standard.Focus();
                txtWeight_Standard.SelectAll();
            }
        }

        private void txtWeight_Standard_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                gcSTD.Focus();
            }
        }

        private void txtTotal_Positive_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtUsable_Positive.Focus();
                txtUsable_Positive.SelectAll();
            }
        }

        private void txtUsable_Positive_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtWeight_Positive.Focus();
                txtWeight_Positive.SelectAll();
            }
        }

        private void txtWeight_Positive_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                gcPOS.Focus();
            }
        }

        private void txtTotal_Negative_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtUsable_Negative.Focus();
                txtUsable_Negative.SelectAll();
            }
        }

        private void txtUsable_Negative_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtWeight_Negative.Focus();
                txtWeight_Negative.SelectAll();
            }
        }

        private void txtWeight_Negative_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                gvNEG.Focus();
            }
        }

        private void gcSTD_ProcessGridKey(object sender, KeyEventArgs e)
        {
            //if (e.KeyCode == Keys.Enter)
            //{
            //    if (gvSTD.FocusedColumn.VisibleIndex == gvSTD.VisibleColumns.Count - 1)
            //        gvSTD.FocusedRowHandle++;
                //gvSTD.FocusedColumn = gvSTD.GetNearestCanFocusedColumn(gvSTD.FocusedColumn);
            //}
            
        }

        private void bbiDELBILL_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            string MARKID = "";
            string MARKNo = "";
            if (tabMARKING.SelectedTabPage == lcgList) //LIST
            {
                GridView gv = gvMARK;
                MARKID = gv.GetFocusedRowCellValue("ID").ToString();
                MARKNo = gv.GetFocusedRowCellValue("Marking No.").ToString();
            }
            else
            {
                MARKID = lblMarkID.Text.Trim();
                MARKNo = txeMarkingNo.Text.Trim();
            }

            if (MARKID != "")
            {
                if (FUNC.msgQuiz("Confirm delete this marking ?\nยืนยันลบเอกสารนี้") == true)
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT TOP (1) Status FROM SMPLRequest WHERE (OIDSMPL = (SELECT TOP(1) OIDSMPL FROM Marking WHERE (OIDMARK = '" + MARKID + "')))");
                    string chkMK = DBC.DBQuery(sbSQL.ToString()).getString();
                    if (chkMK != "2")
                    {
                        sbSQL.Clear();
                        sbSQL.Append("DELETE FROM MarkingDetails WHERE (OIDMARK = '" + MARKID + "')   ");

                        sbSQL.Append("UPDATE SMPLRequest ");
                        sbSQL.Append("SET Status = CASE WHEN CustApproved = 0 THEN 0 ELSE 2 END ");
                        sbSQL.Append("WHERE (OIDSMPL = (SELECT OIDSMPL FROM Marking WHERE (OIDMARK = '" + MARKID + "')))  ");

                        sbSQL.Append("DELETE FROM Marking WHERE (OIDMARK = '" + MARKID + "')   ");

                        try
                        {
                            bool chkSave = this.DBC.DBQuery(sbSQL).runSQL();
                            if (chkSave == true)
                            {
                                FUNC.msgInfo("Delete complete.");
                                //Delete Success
                                if (MARKNo == txeMarkingNo.Text.Trim())
                                    LoadNewData();
                                else
                                {
                                    getGrid_MARK(gcMARK, gvMARK, UserLogin.OIDUser, Convert.ToInt32(rgDocActive.EditValue.ToString()), Convert.ToInt32(rgDocUser.EditValue.ToString()));
                                    HideSelectDoc();
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            FUNC.msgError("Error : " + ex.ToString());
                        }

                    }
                    else
                    {
                        FUNC.msgError("The document cannot be deleted. Because sample request document bound to this marking document has been approved by customers.\nไม่สามารถลบเอกสารได้ เนื่องจากเอกสาร sample request ที่ผูกกับมาร์คกิ้งนี้ได้รับอนุมัติจากลูกค้าแล้ว");
                    }
                }
            }
            else
            {
                FUNC.msgWarning("Please select marking document.");
            }
        }

        private void bbiPrintPDF_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            int[] selectedRowHandles = gvMARK.GetSelectedRows();
            if (selectedRowHandles.Length > 0)
            {
                gvMARK.FocusedRowHandle = selectedRowHandles[0];
                string MARKID = gvMARK.GetRowCellDisplayText(selectedRowHandles[0], "ID");
                string MARKNo = gvMARK.GetRowCellDisplayText(selectedRowHandles[0], "Marking No.");
                if (FUNC.msgQuiz("Confirm print marking (pdf file)  : " + MARKNo + " ?") == true)
                {
                    layoutControlItem120.Text = "Print pdf file processing ..";
                    layoutControlItem120.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                    pbcEXPORT.Properties.Step = 1;
                    pbcEXPORT.Properties.PercentView = true;
                    pbcEXPORT.Properties.Maximum = 11;
                    pbcEXPORT.Properties.Minimum = 0;

                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT    TOP (1) SMPL.SMPLItem, SMPL.ModelName, SMPL.SMPLPatternNo, US.UserName, US.FullName, CUS.Name AS Customer, SMPL.Season ");
                    sbSQL.Append("FROM      Marking AS MK INNER JOIN ");
                    sbSQL.Append("          SMPLRequest AS SMPL ON MK.OIDSMPL = SMPL.OIDSMPL AND MK.OIDMARK = '" + MARKID + "' LEFT OUTER JOIN ");
                    sbSQL.Append("          Users AS US ON MK.CreatedBy = US.OIDUSER LEFT OUTER JOIN ");
                    sbSQL.Append("          Customer AS CUS ON SMPL.OIDCUST = CUS.OIDCUST ");

                    string[] MARK = this.DBC.DBQuery(sbSQL).getMultipleValue();
                    if (MARK.Length > 0)
                    {
                        //****** BEGIN EXPORT *******

                        String sFilePath = System.IO.Path.Combine(new ObjSet.Folder(@"C:\MDS\Export\").GetPath() + MARKNo + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");
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
                            objWorkBook = objApp.Workbooks.Open(this.reportPath + "MARK.xlsx");

                            int LastRow = 9;
                            //** Standard ***
                            LastRow = 9;
                            objSheet = (Microsoft.Office.Interop.Excel.Worksheet)objWorkBook.Sheets[1];

                            objSheet.Cells[2, 1] = MARK[5].ToUpper().Trim();
                            objSheet.Cells[4, 4] = MARK[0];
                            objSheet.Cells[4, 9] = MARK[1];
                            objSheet.Cells[4, 18] = MARK[2];
                            objSheet.Cells[6, 14] = MARK[6];
                            objSheet.Cells[5, 21] = MARK[3].ToUpper().Trim();

                            pbcEXPORT.PerformStep();
                            pbcEXPORT.Update();

                            sbSQL.Clear();
                            sbSQL.Append("SELECT PS.SizeName AS Size, ");
                            sbSQL.Append("       (SELECT GarmentParts + ', ' AS 'data()' FROM(SELECT DISTINCT GP.OIDGParts, GP.GarmentParts FROM SMPLQuantityRequired AS SQR INNER JOIN SMPLRequestFabric AS FB ON SQR.OIDSMPLDT = FB.OIDSMPLDT AND SQR.OIDSMPL = MK.OIDSMPL INNER JOIN SMPLRequestFabricParts AS SFBP ON FB.OIDSMPLFB = SFBP.OIDSMPLFB AND FB.OIDSMPLDT = SFBP.OIDSMPLDT AND FB.OIDITEM = MKD.OIDITEM INNER JOIN GarmentParts AS GP ON SFBP.OIDGParts = GP.OIDGParts AND GP.GarmentParts IS NOT NULL) AS FBParts FOR XML PATH('')) AS FBParts, ");
                            sbSQL.Append("       (SELECT VendFBCode + ', ' AS 'data()' FROM(SELECT DISTINCT VendFBCode FROM SMPLRequestFabric WHERE OIDSMPLDT IN(SELECT OIDSMPLDT FROM SMPLQuantityRequired  WHERE OIDSMPL = MK.OIDSMPL) AND OIDITEM = MKD.OIDITEM AND OIDSMPLDT IN(SELECT value FROM MarkingDetails CROSS APPLY string_split(OIDSMPLDTStuff, ',')  WHERE OIDMARKDT = MKD.OIDMARKDT) AND VendFBCode IS NOT NULL) AS VendFBCode FOR XML PATH('')) AS VendorFBCode, MKD.TotalWidthSTD, MKD.UsableWidth, MKD.GM2, ");
                            sbSQL.Append("       MKD.PracticalLengthCM, MKD.QuantityPCS, MKD.LengthPer1CM, MKD.LengthPer1M, MKD.LengthPer1INCH, MKD.LengthPer1YARD, MKD.WeightG, MKD.WeightKG ");
                            sbSQL.Append("FROM   MarkingDetails AS MKD INNER JOIN ");
                            sbSQL.Append("       Marking AS MK ON MKD.OIDMARK = MK.OIDMARK INNER JOIN ");
                            sbSQL.Append("       ProductSize AS PS ON MKD.OIDSIZE = PS.OIDSIZE ");
                            sbSQL.Append("WHERE (MKD.OIDMARK = '" + MARKID + "') AND(MKD.DetailsType = 0) ");
                            sbSQL.Append("ORDER BY MKD.OIDITEM, MKD.OIDSIZE ");
                            DataTable dtSTD = this.DBC.DBQuery(sbSQL).getDataTable();
                            if (dtSTD != null)
                            {
                                int totalRow = dtSTD.Rows.Count;
                                int diffRow = totalRow > 8 ? totalRow - 8 : 0;

                                if (diffRow > 0)
                                {
                                    //ลบแถว
                                    for (int i = 0; i < diffRow; i++)
                                    {
                                        objSheet.Rows[17].Delete();
                                    }

                                    //แทรกแถว + Merge cell
                                    for (int i = 0; i < diffRow; i++)
                                    {
                                        objSheet.Rows[16].Insert();
                                        objSheet.Range[objSheet.Cells[16, 3], objSheet.Cells[16, 4]].Merge();
                                        objSheet.Range[objSheet.Cells[16, 5], objSheet.Cells[16, 7]].Merge();
                                        objSheet.Range[objSheet.Cells[16, 8], objSheet.Cells[16, 9]].Merge();
                                        objSheet.Range[objSheet.Cells[16, 11], objSheet.Cells[16, 12]].Merge();
                                        objSheet.Range[objSheet.Cells[16, 13], objSheet.Cells[16, 15]].Merge();
                                        objSheet.Range[objSheet.Cells[16, 17], objSheet.Cells[16, 18]].Merge();
                                    }
                                }

                                string chkSize = "";
                                string chkFBParts = "";
                                string chkVendorFBCode = "";
                                string chkTotalWidthSTD = "";
                                string chkUsableWidth = "";
                                string chkGM2 = "";

                                int runRow = 0;
                                foreach (DataRow drMARK in dtSTD.Rows)
                                {
                                    string Size = drMARK["Size"].ToString().ToUpper().Trim();
                                    if (chkSize == Size)
                                        objSheet.Range[objSheet.Cells[LastRow - 1, 2], objSheet.Cells[LastRow, 2]].Merge();
                                    objSheet.Cells[LastRow, 2] = Size;

                                    string FBParts = drMARK["FBParts"].ToString().Trim();
                                    FBParts = FBParts.IndexOf(',') > -1 ? FBParts.ToUpper().Trim().Substring(0, FBParts.ToUpper().Trim().Length - 1) : FBParts.ToUpper().Trim();
                                    if (chkFBParts == FBParts)
                                        objSheet.Range[objSheet.Cells[LastRow - 1, 3], objSheet.Cells[LastRow, 3]].Merge();
                                    objSheet.Cells[LastRow, 3] = FBParts;

                                    string VendorFBCode = drMARK["VendorFBCode"].ToString().Trim();
                                    VendorFBCode = VendorFBCode.IndexOf(',') > -1 ? VendorFBCode.ToUpper().Trim().Substring(0, VendorFBCode.ToUpper().Trim().Length - 1) : VendorFBCode.ToUpper().Trim();
                                    if (chkVendorFBCode == VendorFBCode)
                                        objSheet.Range[objSheet.Cells[LastRow - 1, 5], objSheet.Cells[LastRow, 5]].Merge();
                                    objSheet.Cells[LastRow, 5] = VendorFBCode;

                                    string TotalWidthSTD = drMARK["TotalWidthSTD"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["TotalWidthSTD"].ToString()).ToString("###0.####");
                                    if (chkTotalWidthSTD == TotalWidthSTD)
                                        objSheet.Range[objSheet.Cells[LastRow - 1, 8], objSheet.Cells[LastRow, 8]].Merge();
                                    objSheet.Cells[LastRow, 8] = TotalWidthSTD;

                                    string UsableWidth = drMARK["UsableWidth"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["UsableWidth"].ToString()).ToString("###0.####");
                                    if (chkUsableWidth == UsableWidth)
                                        objSheet.Range[objSheet.Cells[LastRow - 1, 10], objSheet.Cells[LastRow, 10]].Merge();
                                    objSheet.Cells[LastRow, 10] = UsableWidth;

                                    string GM2 = drMARK["GM2"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["GM2"].ToString()).ToString("###0.####");
                                    if (chkGM2 == GM2)
                                        objSheet.Range[objSheet.Cells[LastRow - 1, 11], objSheet.Cells[LastRow, 11]].Merge();
                                    objSheet.Cells[LastRow, 11] = GM2;

                                    string PracticalLengthCM = drMARK["PracticalLengthCM"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["PracticalLengthCM"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 13] = PracticalLengthCM;

                                    string QuantityPCS = drMARK["QuantityPCS"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["QuantityPCS"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 16] = QuantityPCS;

                                    string LengthPer1CM = drMARK["LengthPer1CM"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["LengthPer1CM"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 17] = LengthPer1CM;

                                    string LengthPer1M = drMARK["LengthPer1M"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["LengthPer1M"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 19] = drMARK["LengthPer1M"].ToString().ToUpper().Trim();

                                    string LengthPer1INCH = drMARK["LengthPer1INCH"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["LengthPer1INCH"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 20] = LengthPer1INCH;

                                    string LengthPer1YARD = drMARK["LengthPer1YARD"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["LengthPer1YARD"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 21] = LengthPer1YARD;

                                    string WeightG = drMARK["WeightG"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["WeightG"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 22] = WeightG;

                                    string WeightKG = drMARK["WeightKG"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["WeightKG"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 23] = WeightKG;

                                    if (chkSize != Size || chkFBParts != FBParts || chkVendorFBCode != VendorFBCode || chkTotalWidthSTD != TotalWidthSTD || chkUsableWidth != UsableWidth || chkGM2 != GM2)
                                    {
                                        chkSize = Size;
                                        chkFBParts = FBParts;
                                        chkVendorFBCode = VendorFBCode;
                                        chkTotalWidthSTD = TotalWidthSTD;
                                        chkUsableWidth = UsableWidth;
                                        chkGM2 = GM2;
                                    }

                                    LastRow++;
                                    runRow++;
                                }
                            }

                            pbcEXPORT.PerformStep();
                            pbcEXPORT.Update();

                            //** Positive ***
                            LastRow = 9;
                            objSheet = (Microsoft.Office.Interop.Excel.Worksheet)objWorkBook.Sheets[2];

                            objSheet.Cells[2, 1] = MARK[5].ToUpper().Trim();
                            objSheet.Cells[4, 4] = MARK[0];
                            objSheet.Cells[4, 9] = MARK[1];
                            objSheet.Cells[4, 18] = MARK[2];
                            objSheet.Cells[6, 14] = MARK[6];
                            objSheet.Cells[5, 21] = MARK[3].ToUpper().Trim();

                            pbcEXPORT.PerformStep();
                            pbcEXPORT.Update();

                            sbSQL.Clear();
                            sbSQL.Append("SELECT PS.SizeName AS Size, ");
                            sbSQL.Append("       (SELECT GarmentParts + ', ' AS 'data()' FROM(SELECT DISTINCT GP.OIDGParts, GP.GarmentParts FROM SMPLQuantityRequired AS SQR INNER JOIN SMPLRequestFabric AS FB ON SQR.OIDSMPLDT = FB.OIDSMPLDT AND SQR.OIDSMPL = MK.OIDSMPL INNER JOIN SMPLRequestFabricParts AS SFBP ON FB.OIDSMPLFB = SFBP.OIDSMPLFB AND FB.OIDSMPLDT = SFBP.OIDSMPLDT AND FB.OIDITEM = MKD.OIDITEM INNER JOIN GarmentParts AS GP ON SFBP.OIDGParts = GP.OIDGParts AND GP.GarmentParts IS NOT NULL) AS FBParts FOR XML PATH('')) AS FBParts, ");
                            sbSQL.Append("       (SELECT VendFBCode + ', ' AS 'data()' FROM(SELECT DISTINCT VendFBCode FROM SMPLRequestFabric WHERE OIDSMPLDT IN(SELECT OIDSMPLDT FROM SMPLQuantityRequired  WHERE OIDSMPL = MK.OIDSMPL) AND OIDITEM = MKD.OIDITEM AND OIDSMPLDT IN(SELECT value FROM MarkingDetails CROSS APPLY string_split(OIDSMPLDTStuff, ',')  WHERE OIDMARKDT = MKD.OIDMARKDT) AND VendFBCode IS NOT NULL) AS VendFBCode FOR XML PATH('')) AS VendorFBCode, MKD.TotalWidthSTD, MKD.UsableWidth, MKD.GM2, ");
                            sbSQL.Append("       MKD.PracticalLengthCM, MKD.QuantityPCS, MKD.LengthPer1CM, MKD.LengthPer1M, MKD.LengthPer1INCH, MKD.LengthPer1YARD, MKD.WeightG, MKD.WeightKG ");
                            sbSQL.Append("FROM   MarkingDetails AS MKD INNER JOIN ");
                            sbSQL.Append("       Marking AS MK ON MKD.OIDMARK = MK.OIDMARK INNER JOIN ");
                            sbSQL.Append("       ProductSize AS PS ON MKD.OIDSIZE = PS.OIDSIZE ");
                            sbSQL.Append("WHERE (MKD.OIDMARK = '" + MARKID + "') AND(MKD.DetailsType = 1) ");
                            sbSQL.Append("ORDER BY MKD.OIDITEM, MKD.OIDSIZE ");
                            DataTable dtPOS = this.DBC.DBQuery(sbSQL).getDataTable();
                            if (dtPOS != null)
                            {
                                int totalRow = dtPOS.Rows.Count;
                                int diffRow = totalRow > 8 ? totalRow - 8 : 0;

                                if (diffRow > 0)
                                {
                                    //ลบแถว
                                    for (int i = 0; i < diffRow; i++)
                                    {
                                        objSheet.Rows[17].Delete();
                                    }

                                    //แทรกแถว + Merge cell
                                    for (int i = 0; i < diffRow; i++)
                                    {
                                        objSheet.Rows[16].Insert();
                                        objSheet.Range[objSheet.Cells[16, 3], objSheet.Cells[16, 4]].Merge();
                                        objSheet.Range[objSheet.Cells[16, 5], objSheet.Cells[16, 7]].Merge();
                                        objSheet.Range[objSheet.Cells[16, 8], objSheet.Cells[16, 9]].Merge();
                                        objSheet.Range[objSheet.Cells[16, 11], objSheet.Cells[16, 12]].Merge();
                                        objSheet.Range[objSheet.Cells[16, 13], objSheet.Cells[16, 15]].Merge();
                                        objSheet.Range[objSheet.Cells[16, 17], objSheet.Cells[16, 18]].Merge();
                                    }
                                }

                                string chkSize = "";
                                string chkFBParts = "";
                                string chkVendorFBCode = "";
                                string chkTotalWidthSTD = "";
                                string chkUsableWidth = "";
                                string chkGM2 = "";

                                int runRow = 0;
                                foreach (DataRow drMARK in dtPOS.Rows)
                                {
                                    string Size = drMARK["Size"].ToString().ToUpper().Trim();
                                    if (chkSize == Size)
                                        objSheet.Range[objSheet.Cells[LastRow - 1, 2], objSheet.Cells[LastRow, 2]].Merge();
                                    objSheet.Cells[LastRow, 2] = Size;

                                    string FBParts = drMARK["FBParts"].ToString().Trim();
                                    FBParts = FBParts.IndexOf(',') > -1 ? FBParts.ToUpper().Trim().Substring(0, FBParts.ToUpper().Trim().Length - 1) : FBParts.ToUpper().Trim();
                                    if (chkFBParts == FBParts)
                                        objSheet.Range[objSheet.Cells[LastRow - 1, 3], objSheet.Cells[LastRow, 3]].Merge();
                                    objSheet.Cells[LastRow, 3] = FBParts;

                                    string VendorFBCode = drMARK["VendorFBCode"].ToString().Trim();
                                    VendorFBCode = VendorFBCode.IndexOf(',') > -1 ? VendorFBCode.ToUpper().Trim().Substring(0, VendorFBCode.ToUpper().Trim().Length - 1) : VendorFBCode.ToUpper().Trim();
                                    if (chkVendorFBCode == VendorFBCode)
                                        objSheet.Range[objSheet.Cells[LastRow - 1, 5], objSheet.Cells[LastRow, 5]].Merge();
                                    objSheet.Cells[LastRow, 5] = VendorFBCode;

                                    string TotalWidthSTD = drMARK["TotalWidthSTD"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["TotalWidthSTD"].ToString()).ToString("###0.####");
                                    if (chkTotalWidthSTD == TotalWidthSTD)
                                        objSheet.Range[objSheet.Cells[LastRow - 1, 8], objSheet.Cells[LastRow, 8]].Merge();
                                    objSheet.Cells[LastRow, 8] = TotalWidthSTD;

                                    string UsableWidth = drMARK["UsableWidth"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["UsableWidth"].ToString()).ToString("###0.####");
                                    if (chkUsableWidth == UsableWidth)
                                        objSheet.Range[objSheet.Cells[LastRow - 1, 10], objSheet.Cells[LastRow, 10]].Merge();
                                    objSheet.Cells[LastRow, 10] = UsableWidth;

                                    string GM2 = drMARK["GM2"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["GM2"].ToString()).ToString("###0.####");
                                    if (chkGM2 == GM2)
                                        objSheet.Range[objSheet.Cells[LastRow - 1, 11], objSheet.Cells[LastRow, 11]].Merge();
                                    objSheet.Cells[LastRow, 11] = GM2;

                                    string PracticalLengthCM = drMARK["PracticalLengthCM"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["PracticalLengthCM"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 13] = PracticalLengthCM;

                                    string QuantityPCS = drMARK["QuantityPCS"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["QuantityPCS"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 16] = QuantityPCS;

                                    string LengthPer1CM = drMARK["LengthPer1CM"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["LengthPer1CM"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 17] = LengthPer1CM;

                                    string LengthPer1M = drMARK["LengthPer1M"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["LengthPer1M"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 19] = drMARK["LengthPer1M"].ToString().ToUpper().Trim();

                                    string LengthPer1INCH = drMARK["LengthPer1INCH"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["LengthPer1INCH"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 20] = LengthPer1INCH;

                                    string LengthPer1YARD = drMARK["LengthPer1YARD"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["LengthPer1YARD"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 21] = LengthPer1YARD;

                                    string WeightG = drMARK["WeightG"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["WeightG"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 22] = WeightG;

                                    string WeightKG = drMARK["WeightKG"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["WeightKG"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 23] = WeightKG;

                                    if (chkSize != Size || chkFBParts != FBParts || chkVendorFBCode != VendorFBCode || chkTotalWidthSTD != TotalWidthSTD || chkUsableWidth != UsableWidth || chkGM2 != GM2)
                                    {
                                        chkSize = Size;
                                        chkFBParts = FBParts;
                                        chkVendorFBCode = VendorFBCode;
                                        chkTotalWidthSTD = TotalWidthSTD;
                                        chkUsableWidth = UsableWidth;
                                        chkGM2 = GM2;
                                    }

                                    LastRow++;
                                    runRow++;
                                }
                            }

                            pbcEXPORT.PerformStep();
                            pbcEXPORT.Update();

                            //** Nagative ***
                            LastRow = 9;
                            objSheet = (Microsoft.Office.Interop.Excel.Worksheet)objWorkBook.Sheets[3];

                            objSheet.Cells[2, 1] = MARK[5].ToUpper().Trim();
                            objSheet.Cells[4, 4] = MARK[0];
                            objSheet.Cells[4, 9] = MARK[1];
                            objSheet.Cells[4, 18] = MARK[2];
                            objSheet.Cells[6, 14] = MARK[6];
                            objSheet.Cells[5, 21] = MARK[3].ToUpper().Trim();

                            pbcEXPORT.PerformStep();
                            pbcEXPORT.Update();

                            sbSQL.Clear();
                            sbSQL.Append("SELECT PS.SizeName AS Size, ");
                            sbSQL.Append("       (SELECT GarmentParts + ', ' AS 'data()' FROM(SELECT DISTINCT GP.OIDGParts, GP.GarmentParts FROM SMPLQuantityRequired AS SQR INNER JOIN SMPLRequestFabric AS FB ON SQR.OIDSMPLDT = FB.OIDSMPLDT AND SQR.OIDSMPL = MK.OIDSMPL INNER JOIN SMPLRequestFabricParts AS SFBP ON FB.OIDSMPLFB = SFBP.OIDSMPLFB AND FB.OIDSMPLDT = SFBP.OIDSMPLDT AND FB.OIDITEM = MKD.OIDITEM INNER JOIN GarmentParts AS GP ON SFBP.OIDGParts = GP.OIDGParts AND GP.GarmentParts IS NOT NULL) AS FBParts FOR XML PATH('')) AS FBParts, ");
                            sbSQL.Append("       (SELECT VendFBCode + ', ' AS 'data()' FROM(SELECT DISTINCT VendFBCode FROM SMPLRequestFabric WHERE OIDSMPLDT IN(SELECT OIDSMPLDT FROM SMPLQuantityRequired  WHERE OIDSMPL = MK.OIDSMPL) AND OIDITEM = MKD.OIDITEM AND OIDSMPLDT IN(SELECT value FROM MarkingDetails CROSS APPLY string_split(OIDSMPLDTStuff, ',')  WHERE OIDMARKDT = MKD.OIDMARKDT) AND VendFBCode IS NOT NULL) AS VendFBCode FOR XML PATH('')) AS VendorFBCode, MKD.TotalWidthSTD, MKD.UsableWidth, MKD.GM2, ");
                            sbSQL.Append("       MKD.PracticalLengthCM, MKD.QuantityPCS, MKD.LengthPer1CM, MKD.LengthPer1M, MKD.LengthPer1INCH, MKD.LengthPer1YARD, MKD.WeightG, MKD.WeightKG ");
                            sbSQL.Append("FROM   MarkingDetails AS MKD INNER JOIN ");
                            sbSQL.Append("       Marking AS MK ON MKD.OIDMARK = MK.OIDMARK INNER JOIN ");
                            sbSQL.Append("       ProductSize AS PS ON MKD.OIDSIZE = PS.OIDSIZE ");
                            sbSQL.Append("WHERE (MKD.OIDMARK = '" + MARKID + "') AND(MKD.DetailsType = 2) ");
                            sbSQL.Append("ORDER BY MKD.OIDITEM, MKD.OIDSIZE ");
                            DataTable dtNEG = this.DBC.DBQuery(sbSQL).getDataTable();
                            if (dtNEG != null)
                            {
                                int totalRow = dtNEG.Rows.Count;
                                int diffRow = totalRow > 8 ? totalRow - 8 : 0;

                                if (diffRow > 0)
                                {
                                    //ลบแถว
                                    for (int i = 0; i < diffRow; i++)
                                    {
                                        objSheet.Rows[17].Delete();
                                    }

                                    //แทรกแถว + Merge cell
                                    for (int i = 0; i < diffRow; i++)
                                    {
                                        objSheet.Rows[16].Insert();
                                        objSheet.Range[objSheet.Cells[16, 3], objSheet.Cells[16, 4]].Merge();
                                        objSheet.Range[objSheet.Cells[16, 5], objSheet.Cells[16, 7]].Merge();
                                        objSheet.Range[objSheet.Cells[16, 8], objSheet.Cells[16, 9]].Merge();
                                        objSheet.Range[objSheet.Cells[16, 11], objSheet.Cells[16, 12]].Merge();
                                        objSheet.Range[objSheet.Cells[16, 13], objSheet.Cells[16, 15]].Merge();
                                        objSheet.Range[objSheet.Cells[16, 17], objSheet.Cells[16, 18]].Merge();
                                    }
                                }

                                string chkSize = "";
                                string chkFBParts = "";
                                string chkVendorFBCode = "";
                                string chkTotalWidthSTD = "";
                                string chkUsableWidth = "";
                                string chkGM2 = "";

                                int runRow = 0;
                                foreach (DataRow drMARK in dtNEG.Rows)
                                {
                                    string Size = drMARK["Size"].ToString().ToUpper().Trim();
                                    if (chkSize == Size)
                                        objSheet.Range[objSheet.Cells[LastRow - 1, 2], objSheet.Cells[LastRow, 2]].Merge();
                                    objSheet.Cells[LastRow, 2] = Size;

                                    string FBParts = drMARK["FBParts"].ToString().Trim();
                                    FBParts = FBParts.IndexOf(',') > -1 ? FBParts.ToUpper().Trim().Substring(0, FBParts.ToUpper().Trim().Length - 1) : FBParts.ToUpper().Trim();
                                    if (chkFBParts == FBParts)
                                        objSheet.Range[objSheet.Cells[LastRow - 1, 3], objSheet.Cells[LastRow, 3]].Merge();
                                    objSheet.Cells[LastRow, 3] = FBParts;

                                    string VendorFBCode = drMARK["VendorFBCode"].ToString().Trim();
                                    VendorFBCode = VendorFBCode.IndexOf(',') > -1 ? VendorFBCode.ToUpper().Trim().Substring(0, VendorFBCode.ToUpper().Trim().Length - 1) : VendorFBCode.ToUpper().Trim();
                                    if (chkVendorFBCode == VendorFBCode)
                                        objSheet.Range[objSheet.Cells[LastRow - 1, 5], objSheet.Cells[LastRow, 5]].Merge();
                                    objSheet.Cells[LastRow, 5] = VendorFBCode;

                                    string TotalWidthSTD = drMARK["TotalWidthSTD"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["TotalWidthSTD"].ToString()).ToString("###0.####");
                                    if (chkTotalWidthSTD == TotalWidthSTD)
                                        objSheet.Range[objSheet.Cells[LastRow - 1, 8], objSheet.Cells[LastRow, 8]].Merge();
                                    objSheet.Cells[LastRow, 8] = TotalWidthSTD;

                                    string UsableWidth = drMARK["UsableWidth"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["UsableWidth"].ToString()).ToString("###0.####");
                                    if (chkUsableWidth == UsableWidth)
                                        objSheet.Range[objSheet.Cells[LastRow - 1, 10], objSheet.Cells[LastRow, 10]].Merge();
                                    objSheet.Cells[LastRow, 10] = UsableWidth;

                                    string GM2 = drMARK["GM2"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["GM2"].ToString()).ToString("###0.####");
                                    if (chkGM2 == GM2)
                                        objSheet.Range[objSheet.Cells[LastRow - 1, 11], objSheet.Cells[LastRow, 11]].Merge();
                                    objSheet.Cells[LastRow, 11] = GM2;

                                    string PracticalLengthCM = drMARK["PracticalLengthCM"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["PracticalLengthCM"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 13] = PracticalLengthCM;

                                    string QuantityPCS = drMARK["QuantityPCS"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["QuantityPCS"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 16] = QuantityPCS;

                                    string LengthPer1CM = drMARK["LengthPer1CM"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["LengthPer1CM"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 17] = LengthPer1CM;

                                    string LengthPer1M = drMARK["LengthPer1M"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["LengthPer1M"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 19] = drMARK["LengthPer1M"].ToString().ToUpper().Trim();

                                    string LengthPer1INCH = drMARK["LengthPer1INCH"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["LengthPer1INCH"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 20] = LengthPer1INCH;

                                    string LengthPer1YARD = drMARK["LengthPer1YARD"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["LengthPer1YARD"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 21] = LengthPer1YARD;

                                    string WeightG = drMARK["WeightG"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["WeightG"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 22] = WeightG;

                                    string WeightKG = drMARK["WeightKG"].ToString() == "" ? "0" : Convert.ToDouble(drMARK["WeightKG"].ToString()).ToString("###0.####");
                                    objSheet.Cells[LastRow, 23] = WeightKG;

                                    if (chkSize != Size || chkFBParts != FBParts || chkVendorFBCode != VendorFBCode || chkTotalWidthSTD != TotalWidthSTD || chkUsableWidth != UsableWidth || chkGM2 != GM2)
                                    {
                                        chkSize = Size;
                                        chkFBParts = FBParts;
                                        chkVendorFBCode = VendorFBCode;
                                        chkTotalWidthSTD = TotalWidthSTD;
                                        chkUsableWidth = UsableWidth;
                                        chkGM2 = GM2;
                                    }

                                    LastRow++;
                                    runRow++;
                                }
                            }

                            pbcEXPORT.PerformStep();
                            pbcEXPORT.Update();


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
                        FUNC.msgError("ไม่พบข้อมูลเอกสาร Marking: " + MARKNo);
                    }
                    layoutControlItem120.Text = "Status ..";
                    layoutControlItem120.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                }

            }
        }
    }
}