using System;
using System.Text;
using DBConnect;
using System.Windows.Forms;
using DevExpress.LookAndFeel;
using DevExpress.Utils.Extensions;
using System.Drawing;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraEditors;
using System.Data;
using DevExpress.Data.Extensions;
using System.Linq;
using TheepClass;
using DevExpress.Utils;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;

namespace MDS.Master
{
    public partial class M09 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private Functionality.Function FUNC = new Functionality.Function();
        public LogIn UserLogin { get; set; }
        public int Company { get; set; }
        public string ConnectionString { get; set; }
        string CONNECT_STRING = "";
        DatabaseConnect DBC;

        public M09()
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

            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT TOP (1) ReadWriteStatus FROM FunctionAccess WHERE (OIDUser = '" + UserLogin.OIDUser + "') AND(FunctionNo = 'M09') ");
            int chkReadWrite = this.DBC.DBQuery(sbSQL).getInt();
            if (chkReadWrite == 0)
                ribbonPageGroup1.Visible = false;

            sbSQL.Clear();
            sbSQL.Append("SELECT FullName, OIDUSER FROM Users ORDER BY OIDUSER ");
            new ObjDE.setGridLookUpEdit(glueCREATE, sbSQL, "FullName", "OIDUSER").getData();

            glueCREATE.EditValue = UserLogin.OIDUser;

            bbiNew.PerformClick();
        }

        private void LoadData()
        {

            StringBuilder sbSQL = new StringBuilder();

            sbSQL.Clear();
            sbSQL.Append("SELECT DISTINCT CUS.Code, CUS.ShortName, CUS.Name, CUS.OIDCUST AS ID ");
            sbSQL.Append("FROM   Customer AS CUS INNER JOIN ");
            sbSQL.Append("       ProductionLine AS PL ON CUS.OIDCUST = PL.OIDCUST ");
            sbSQL.Append("ORDER BY CUS.ShortName ");
            new ObjDE.setSearchLookUpEdit(slueCustomer, sbSQL, "ShortName", "ID").getData(true);
            //MessageBox.Show("3");

            LoadCapacity();
        }

        private void LoadCapacity()
        {
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT PC.OIDPCAP AS CapacityID, PC.OIDCUST AS CustomerID, CUS.ShortName AS CustomerName, PC.OIDGCATEGORY AS CategoryID, GC.CategoryName, PC.OIDSTYLE AS StyleID, PS.StyleName, PC.QTYPerHour, PC.QTYPerDay, ");
            sbSQL.Append("       PC.QTYPerOT, FORMAT(PC.STDTimeCUT, '###0.####') AS StandardTimeCutting, FORMAT(PC.STDTimePAD, '###0.####') AS StandardTimePadPrint, FORMAT(PC.STDTimeSEW, '###0.####') AS StandardTimeSewing, FORMAT(PC.STDTimePACK, '###0.####') AS StandardTimePacking, FORMAT(PC.STDTime, '###0.####') AS StandardTime, ");
            sbSQL.Append("       CASE WHEN ISNULL(PC.ProductionStartDate, '') = '' THEN '' ELSE CONVERT(VARCHAR(10), PC.ProductionStartDate, 103) END AS ProductionStartDate, PC.CreatedBy, PC.CreatedDate ");
            sbSQL.Append("FROM   ProductionCapacity AS PC INNER JOIN ");
            sbSQL.Append("       Customer AS CUS ON PC.OIDCUST = CUS.OIDCUST INNER JOIN ");
            sbSQL.Append("       ProductStyle AS PS ON PC.OIDSTYLE = PS.OIDSTYLE INNER JOIN ");
            sbSQL.Append("       GarmentCategory AS GC ON PC.OIDGCATEGORY = GC.OIDGCATEGORY ");
            sbSQL.Append("WHERE (PC.OIDPCAP <> '') ");
            if (slueCustomer.Text.Trim() != "")
            {
                sbSQL.Append("AND (PC.OIDCUST = '" + slueCustomer.EditValue.ToString() + "') ");
            }
            sbSQL.Append("ORDER BY CustomerID, GC.CategoryName, PS.StyleName ");
            new ObjDE.setGridControl(gcCapacity, gvCapacity, sbSQL).getData(false, false, false, true);
            gvCapacity.Columns["CapacityID"].Visible = false;
            gvCapacity.Columns["CustomerID"].Visible = false;
            gvCapacity.Columns["CategoryID"].Visible = false;
            gvCapacity.Columns["StyleID"].Visible = false;
            gvCapacity.Columns["CreatedBy"].Visible = false;
            gvCapacity.Columns["CreatedDate"].Visible = false;
        }

        private void CreateTabPage(string CusID="", string CategoryID="")
        {
            //gcCapacity.DataSource = null;
            tabBranch.TabPages.Clear();

            if (CusID != "" && CategoryID != "")
            {
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT B.Name AS Branch, B.OIDBranch AS ID ");
                sbSQL.Append("FROM ProductionLine AS PL INNER JOIN ");
                sbSQL.Append("     Branchs AS B ON PL.Branch = B.OIDBranch ");
                sbSQL.Append("WHERE (PL.OIDCUST = '" + CusID + "') AND(PL.OIDCATEGORY = '" + CategoryID + "') ");
                sbSQL.Append("ORDER BY B.OIDBranch ");
                DataTable dtBranch = this.DBC.DBQuery(sbSQL).getDataTable();
                foreach (DataRow drRow in dtBranch.Rows)
                {
                    DevExpress.XtraTab.XtraTabPage tabPage = new DevExpress.XtraTab.XtraTabPage();
                    tabPage.Name = "B" + drRow["ID"].ToString();
                    tabPage.Text = drRow["Branch"].ToString();

                    DevExpress.XtraGrid.GridControl grid = new DevExpress.XtraGrid.GridControl();
                    grid.Name = "gc" + drRow["ID"].ToString();
                    GridView view = new GridView();
                    view.Name = "gv" + drRow["ID"].ToString();

                    tabPage.Controls.Add(grid);

                    grid.Dock = DockStyle.Fill;
                    grid.ViewCollection.Add(view);
                    grid.MainView = view;
                    view.GridControl = grid;
                    view.OptionsView.ShowAutoFilterRow = false;
                    view.OptionsBehavior.Editable = false;
                    view.OptionsView.EnableAppearanceEvenRow = true;
                    view.OptionsView.EnableAppearanceOddRow = true;

                    StringBuilder sbLINE = new StringBuilder();
                    sbLINE.Append("SELECT LN.LINENAME AS LineName, LN.OIDLINE AS LineID, LN.Branch AS BranchID  ");
                    sbLINE.Append("FROM   ProductionLine AS PL INNER JOIN ");
                    sbLINE.Append("        LineNumber AS LN ON PL.OIDLine = LN.OIDLINE ");
                    sbLINE.Append("WHERE (PL.OIDCUST = '" + CusID + "') AND(PL.OIDCATEGORY = '" + CategoryID + "') AND(PL.Branch = '" + drRow["ID"].ToString() + "') ");
                    sbLINE.Append("ORDER By LN.LINENAME ");
                    DataTable dtLINE = this.DBC.DBQuery(sbLINE).getDataTable();

                    grid.DataSource = dtLINE;

                    grid.EndUpdate();
                    grid.ResumeLayout();
                    view.OptionsView.ColumnAutoWidth = true;
                    view.BestFitColumns();
                    view.RowCellClick += gvLine_RowCellClick;
                    view.DataSourceChanged += gvLine_DataSourceChanged;
                    view.RefreshData();
                    grid.RefreshDataSource();

                    tabBranch.TabPages.Add(tabPage);
                    tabBranch.ResumeLayout(false);
                    tabBranch.LayoutChanged();
                }

            }
        }

        private void CreateTabPageLine(string CusID = "", string CategoryID = "", string StyleID = "")
        {
            //gcCapacity.DataSource = null;
            tabBranch.TabPages.Clear();

            if (CusID != "" && CategoryID != "")
            {
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT DISTINCT B.Name AS Branch, B.OIDBranch AS ID ");
                sbSQL.Append("FROM ProductionLine AS PL INNER JOIN ");
                sbSQL.Append("     Branchs AS B ON PL.Branch = B.OIDBranch ");
                sbSQL.Append("WHERE (PL.OIDCUST = '" + CusID + "') AND(PL.OIDCATEGORY = '" + CategoryID + "') ");
                sbSQL.Append("ORDER BY B.OIDBranch ");
                DataTable dtBranch = this.DBC.DBQuery(sbSQL).getDataTable();
                int BCount = this.DBC.DBQuery(sbSQL).getCount();

                foreach (DataRow drRow in dtBranch.Rows)
                {
                    DevExpress.XtraTab.XtraTabPage tabPage = new DevExpress.XtraTab.XtraTabPage();
                    tabPage.Name = "B" + drRow["ID"].ToString();
                    tabPage.Text = drRow["Branch"].ToString();

                    CheckedListBoxControl clbLine = new CheckedListBoxControl();
                    clbLine.Name = "LN" + drRow["ID"].ToString();


                    tabPage.Controls.Add(clbLine);
                    clbLine.Dock = DockStyle.Fill;

                    StringBuilder sbLINE = new StringBuilder();
                    sbLINE.Append("SELECT LN.LINENAME AS LineName, LN.OIDLINE AS LineID, LN.Branch AS BranchID  ");
                    sbLINE.Append("FROM   ProductionLine AS PL INNER JOIN ");
                    sbLINE.Append("        LineNumber AS LN ON PL.OIDLine = LN.OIDLINE ");
                    sbLINE.Append("WHERE (PL.OIDCUST = '" + CusID + "') AND(PL.OIDCATEGORY = '" + CategoryID + "') AND(PL.Branch = '" + drRow["ID"].ToString() + "') ");
                    sbLINE.Append("ORDER By LN.LINENAME ");
                    DataTable dtLINE = this.DBC.DBQuery(sbLINE).getDataTable();

                    clbLine.ValueMember = "LineName";
                    clbLine.DisplayMember = "LineName";
                    clbLine.DataSource = dtLINE;

                    StringBuilder sbCapacity = new StringBuilder();
                    sbCapacity.Append("SELECT DISTINCT LINEID AS LineName ");
                    sbCapacity.Append("FROM   ProductionCapacityLine ");
                    sbCapacity.Append("WHERE(OIDCAP IN ");
                    sbCapacity.Append("             (SELECT OIDPCAP ");
                    sbCapacity.Append("              FROM   ProductionCapacity ");
                    sbCapacity.Append("              WHERE (OIDCUST = '" + CusID + "') AND (OIDGCATEGORY = '" + CategoryID + "'))) AND (OIDBranch = '" + drRow["ID"].ToString() + "') ");
                    DataTable dtQC = this.DBC.DBQuery(sbCapacity).getDataTable();
                    foreach (DataRow row in dtQC.Rows)
                    {
                        for (int i = 0; i < clbLine.ItemCount; i++)
                        {
                            if (row["LineName"].ToString() == clbLine.GetItemValue(i).ToString())
                            {
                                clbLine.SetItemCheckState(i, CheckState.Checked);
                                break;
                            }
                        }
                    }

                    tabBranch.TabPages.Add(tabPage);
                    tabBranch.ResumeLayout(false);
                    tabBranch.LayoutChanged();
                }

                sbSQL.Clear();
                sbSQL.Append("SELECT PC.OIDPCAP AS CapacityID, PC.OIDCUST AS CustomerID, CUS.ShortName AS CustomerName, PC.OIDGCATEGORY AS CategoryID, GC.CategoryName, PC.OIDSTYLE AS StyleID, PS.StyleName, PC.QTYPerHour, PC.QTYPerDay, ");
                sbSQL.Append("       PC.QTYPerOT, FORMAT(PC.STDTimeCUT, '###0.####') AS StandardTimeCutting, FORMAT(PC.STDTimePAD, '###0.####') AS StandardTimePadPrint, FORMAT(PC.STDTimeSEW, '###0.####') AS StandardTimeSewing, FORMAT(PC.STDTimePACK, '###0.####') AS StandardTimePacking, FORMAT(PC.STDTime, '###0.####') AS StandardTime, ");
                sbSQL.Append("       CASE WHEN ISNULL(PC.ProductionStartDate, '') = '' THEN '' ELSE CONVERT(VARCHAR(10), PC.ProductionStartDate, 103) END AS ProductionStartDate, PC.CreatedBy, PC.CreatedDate ");
                sbSQL.Append("FROM   ProductionCapacity AS PC INNER JOIN ");
                sbSQL.Append("       Customer AS CUS ON PC.OIDCUST = CUS.OIDCUST INNER JOIN ");
                sbSQL.Append("       ProductStyle AS PS ON PC.OIDSTYLE = PS.OIDSTYLE INNER JOIN ");
                sbSQL.Append("       GarmentCategory AS GC ON PC.OIDGCATEGORY = GC.OIDGCATEGORY ");
                sbSQL.Append("WHERE  (PC.OIDCUST = '" + CusID + "') ");
                if (CusID != "")
                {
                    sbSQL.Append("AND (PC.OIDGCATEGORY = '" + CategoryID + "') ");
                }
                if (StyleID != "")
                {
                    sbSQL.Append("AND (PC.OIDSTYLE = '" + StyleID + "') ");
                }
                sbSQL.Append("ORDER BY CustomerID, GC.CategoryName, PS.StyleName ");
                new ObjDE.setGridControl(gcCapacity, gvCapacity, sbSQL).getData(false, false, false, true);
                gvCapacity.Columns["CapacityID"].Visible = false;
                gvCapacity.Columns["CustomerID"].Visible = false;
                gvCapacity.Columns["CategoryID"].Visible = false;
                gvCapacity.Columns["StyleID"].Visible = false;
                gvCapacity.Columns["CreatedBy"].Visible = false;
                gvCapacity.Columns["CreatedDate"].Visible = false;
            }
        }

        private void gvLine_DataSourceChanged(object sender, EventArgs e)
        {
            GridView view = sender as GridView;
            view.Columns["LineID"].Visible = false;
            view.Columns["BranchID"].Visible = false;
        }

        private void gvLine_RowLoaded(object sender, DevExpress.XtraGrid.Views.Base.RowEventArgs e)
        {
            GetVisible((GridView)sender, e.RowHandle);
        }

        private void GetVisible(GridView aView, int aRowHandle)
        {
            aView.Columns[1].Visible = false;
        }

        private void gvLine_RowCellClick(object sender, RowCellClickEventArgs e)
        {
            GetValue((GridView)sender, e.RowHandle);      
        }

        private void GetValue(GridView aView, int aRowHandle)
        {
            string LineName = aView.GetRowCellValue(aRowHandle, aView.Columns[0]).ToString();
            string LineID = aView.GetRowCellValue(aRowHandle, aView.Columns[1]).ToString();
            string BranchID = aView.GetRowCellValue(aRowHandle, aView.Columns[2]).ToString();
            MessageBox.Show("LineName:"+ LineName+", LineID:" + LineID+", BranchID:" + BranchID);
        }

        private void NewData()
        {
            lblStatus.Text = "* Add Capacity";
            lblStatus.ForeColor = Color.Green;

            txeID.Text = this.DBC.DBQuery("SELECT CASE WHEN ISNULL(MAX(OIDPCAP), '') = '' THEN 1 ELSE MAX(OIDPCAP) + 1 END AS NewNo FROM ProductionCapacity").getString();
            //txeID.Text = "";

            //gcCapacity.DataSource = null;
            tabBranch.TabPages.Clear();

            ClearData();
        }

        private void ClearData()
        {
            txe1Hr.Text = "";
            txe1Day.Text = "";
            txeOT.Text = "";

            txeCutting.Text = "";
            txePadPrint.Text = "";
            txeSewing.Text = "";
            txePacking.Text = "";
            txeStdTime.Text = "";

            dteStart.EditValue = DateTime.Now;
            glueCREATE.EditValue = UserLogin.OIDUser;
            txeDATE.Text = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
        }

        private void bbiNew_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            LoadData();
            NewData();
        }

        private void gvGarment_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {
            
        }

        private void slueCustomer_EditValueChanged(object sender, EventArgs e)
        {
            glueCategory.EditValue = "";
            glueCategory.Properties.DataSource = null;
            ClearData();

            if (slueCustomer.Text.Trim() != "")
            {
                string CUSID = slueCustomer.EditValue.ToString();
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT DISTINCT GC.OIDGCATEGORY AS ID, GC.CategoryName ");
                sbSQL.Append("FROM GarmentCategory AS GC INNER JOIN ");
                sbSQL.Append("     ProductionLine AS PL ON GC.OIDGCATEGORY = PL.OIDCATEGORY ");
                sbSQL.Append("WHERE (PL.OIDCUST = '" + CUSID + "') ");
                sbSQL.Append("ORDER BY ID ");
                new ObjDE.setGridLookUpEdit(glueCategory, sbSQL, "CategoryName", "ID").getData(true);

                CreateTabPageLine(CUSID);
            }
            LoadCapacity();
            glueCategory.Focus();
        }

        private void glueCategory_EditValueChanged(object sender, EventArgs e)
        {
            slueStyle.EditValue = "";
            slueStyle.Properties.DataSource = null;
            ClearData();

            if (glueCategory.Text.Trim() != "")
            {
                string CATEID = glueCategory.EditValue.ToString();
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT StyleName, OIDSTYLE AS ID ");
                sbSQL.Append("FROM   ProductStyle ");
                sbSQL.Append("WHERE (OIDGCATEGORY = '" + CATEID + "') ");
                new ObjDE.setSearchLookUpEdit(slueStyle, sbSQL, "StyleName", "ID").getData(true);
            }

            //***** LOAD BRANCH *************
            tabBranch.TabPages.Clear();
            string CUSID = "";
            string CATGID = "";

            if (slueCustomer.Text.Trim() != "")
            {
                CUSID = slueCustomer.EditValue.ToString();
            }
            if (glueCategory.Text.Trim() != "")
            {
                CATGID = glueCategory.EditValue.ToString();
            }
            CreateTabPageLine(CUSID, CATGID);
            //*******************************

            slueStyle.Focus();

        }

        private void gvCapacity_RowStyle(object sender, RowStyleEventArgs e)
        {
            
        }

        private void slueStyle_EditValueChanged(object sender, EventArgs e)
        {
            string CUSID = "";
            string CATGID = "";
            string STYLEID = "";

            if (slueCustomer.Text.Trim() != "")
            {
                CUSID = slueCustomer.EditValue.ToString();
            }
            if (glueCategory.Text.Trim() != "")
            {
                CATGID = glueCategory.EditValue.ToString();
            }
            if (slueStyle.Text.Trim() != "")
            {
                STYLEID = slueStyle.EditValue.ToString();
            }
            CreateTabPageLine(CUSID, CATGID, STYLEID);

            //******* LOAD DATA ***********
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT TOP(1) OIDPCAP, QTYPerHour, QTYPerDay, ");
            sbSQL.Append("       QTYPerOT, FORMAT(STDTimeCUT, '###0.####') AS StandardTimeCutting, FORMAT(STDTimePAD, '###0.####') AS StandardTimePadPrint, FORMAT(STDTimeSEW, '###0.####') AS StandardTimeSewing, FORMAT(STDTimePACK, '###0.####') AS StandardTimePacking, FORMAT(STDTime, '###0.####') AS StandardTime, ");
            sbSQL.Append("       ProductionStartDate, CreatedBy, CreatedDate ");
            sbSQL.Append("FROM   ProductionCapacity ");
            sbSQL.Append("WHERE  (OIDCUST = '" + CUSID + "') AND (OIDGCATEGORY = '" + CATGID + "') AND (OIDSTYLE = '" + STYLEID + "') ");
            string[] arrCapacity = this.DBC.DBQuery(sbSQL).getMultipleValue();
            if (arrCapacity.Length > 0)
            {
                lblStatus.Text = "* Edit Capacity";
                lblStatus.ForeColor = Color.Red;

                txeID.Text = arrCapacity[0];
                txe1Hr.Text = arrCapacity[1];
                txe1Day.Text = arrCapacity[2];
                txeOT.Text = arrCapacity[3];

                txeCutting.Text = arrCapacity[4];
                txePadPrint.Text = arrCapacity[5];
                txeSewing.Text = arrCapacity[6];
                txePacking.Text = arrCapacity[7];
                txeStdTime.Text = arrCapacity[8];

                dteStart.EditValue = Convert.ToDateTime(arrCapacity[9]);

                glueCREATE.EditValue = arrCapacity[10];
                txeDATE.Text = arrCapacity[11];
            }
            else
            {
                lblStatus.Text = "* Add Capacity";
                lblStatus.ForeColor = Color.Green;

                txeID.Text = "";
                txe1Hr.Text = "";
                txe1Day.Text = "";
                txeOT.Text = "";

                txeCutting.Text = "";
                txePadPrint.Text = "";
                txeSewing.Text = "";
                txePacking.Text = "";
                txeStdTime.Text = "";

                dteStart.EditValue = DateTime.Now;

                glueCREATE.EditValue = UserLogin.OIDUser;
                txeDATE.Text = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
            }

            txe1Hr.Focus();
        }

        private void bbiSave_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (slueCustomer.Text.Trim() == "")
            {
                FUNC.msgWarning("Please select customer.");
                slueCustomer.Focus();
            }
            else if (glueCategory.Text.Trim() == "")
            {
                FUNC.msgWarning("Please product category.");
                glueCategory.Focus();
            }
            else if (slueStyle.Text.Trim() == "")
            {
                FUNC.msgWarning("Please select product style.");
                slueStyle.Focus();
            }
            else
            {
                if (FUNC.msgQuiz("Confirm save data ?") == true)
                {
                    StringBuilder sbSAVE = new StringBuilder();
                    string strCREATE = UserLogin.OIDUser.ToString() != "" ? UserLogin.OIDUser.ToString() : "0";

                    //******** save ProductionCapacity table ************
                    string CUSTOMER = slueCustomer.EditValue.ToString();
                    string CATEGORY = glueCategory.EditValue.ToString();
                    string STYLEID = slueStyle.EditValue.ToString();

                    string QTYPerHour = txe1Hr.Text.Trim();
                    QTYPerHour = QTYPerHour != "" ? QTYPerHour : "0";
                    string QTYPerDay = txe1Day.Text.Trim();
                    QTYPerDay = QTYPerDay != "" ? QTYPerDay : "0";
                    string QTYPerOT = txeOT.Text.Trim();
                    QTYPerOT = QTYPerOT != "" ? QTYPerOT : "0";

                    string STDTimeCUT = txeCutting.Text.Trim();
                    STDTimeCUT = STDTimeCUT != "" ? STDTimeCUT : "0";
                    string STDTimePAD = txePadPrint.Text.Trim();
                    STDTimePAD = STDTimePAD != "" ? STDTimePAD : "0";
                    string STDTimeSEW = txeSewing.Text.Trim();
                    STDTimeSEW = STDTimeSEW != "" ? STDTimeSEW : "0";
                    string STDTimePACK = txePacking.Text.Trim();
                    STDTimePACK = STDTimePACK != "" ? STDTimePACK : "0";
                    string STDTime = txeStdTime.Text.Trim();
                    STDTime = STDTime != "" ? STDTime : "0";

                    if (lblStatus.Text == "* Add Capacity")
                    {
                        sbSAVE.Append(" INSERT INTO ProductionCapacity(OIDCUST, OIDGCATEGORY, OIDSTYLE, QTYPerHour, QTYPerDay, QTYPerOT, STDTimeCUT, STDTimePAD, STDTimeSEW, STDTimePACK, STDTime, ProductionStartDate, CreatedBy, CreatedDate) ");
                        sbSAVE.Append("  VALUES('" + CUSTOMER + "', '" + CATEGORY + "', '" + STYLEID + "', '" + QTYPerHour + "', '" + QTYPerDay + "', '" + QTYPerOT + "', '" + STDTimeCUT + "', '" + STDTimePAD + "', '" + STDTimeSEW + "', '" + STDTimePACK + "', '" + STDTime + "', '" + Convert.ToDateTime(dteStart.Text).ToString("yyyy-MM-dd") + "', '" + strCREATE + "', GETDATE()) ");
                    }
                    else if (lblStatus.Text == "* Edit Capacity")
                    {
                        sbSAVE.Append(" UPDATE ProductionCapacity SET ");
                        sbSAVE.Append("  OIDCUST='" + CUSTOMER + "', OIDGCATEGORY='" + CATEGORY + "', OIDSTYLE='" + STYLEID + "', QTYPerHour='" + QTYPerHour + "', QTYPerDay='" + QTYPerDay + "', ");
                        sbSAVE.Append("  QTYPerOT='" + QTYPerOT + "', STDTimeCUT='" + STDTimeCUT + "', STDTimePAD='" + STDTimePAD + "', STDTimeSEW='" + STDTimeSEW + "', STDTimePACK='" + STDTimePACK + "', ");
                        sbSAVE.Append("  STDTime='" + STDTime + "', ProductionStartDate='" + Convert.ToDateTime(dteStart.Text).ToString("yyyy-MM-dd") + "' ");
                        sbSAVE.Append(" WHERE (OIDPCAP = '" + txeID.Text.Trim() + "') ");
                    }

                    if (sbSAVE.Length > 0)
                    {
                        try
                        {
                            bool chkSAVE = this.DBC.DBQuery(sbSAVE).runSQL();
                            if (chkSAVE == true)
                            {
                                string strCAP = this.DBC.DBQuery("SELECT TOP (1) OIDPCAP FROM ProductionCapacity WHERE (OIDCUST = '" + CUSTOMER + "') AND(OIDGCATEGORY = '" + CATEGORY + "') AND(OIDSTYLE = '" + STYLEID + "')").getString();

                                if (strCAP != "") //Save ProductionCapacityLine
                                {
                                    sbSAVE.Clear();
                                    StringBuilder sbSQL = new StringBuilder();
                                    sbSQL.Append("SELECT DISTINCT Branch ");
                                    sbSQL.Append("FROM   ProductionLine ");
                                    sbSQL.Append("ORDER BY Branch ");
                                    DataTable dtLINE = this.DBC.DBQuery(sbSQL).getDataTable();
                                    foreach (DataRow row in dtLINE.Rows)
                                    {
                                        CheckedListBoxControl clb = this.Controls.Find("LN" + row["Branch"], true).FirstOrDefault() as CheckedListBoxControl;
                                        if (clb != null)
                                        {
                                            string strBRANCH = row["Branch"].ToString();
                                            string strLINE = "";
                                            int iCQC = 0;
                                            foreach (DataRowView item in clb.CheckedItems)
                                            {
                                                if (iCQC != 0)
                                                {
                                                    strLINE += ", ";
                                                }
                                                strLINE += "'" + item["LineName"].ToString() + "'";
                                                sbSAVE.Append("IF NOT EXISTS(SELECT OIDLCAPLine FROM ProductionCapacityLine WHERE (OIDCAP = '" + strCAP + "') AND (OIDBranch = '" + strBRANCH + "') AND (LINEID = '" + item["LineName"].ToString() + "')) ");
                                                sbSAVE.Append(" BEGIN ");
                                                sbSAVE.Append("  INSERT INTO ProductionCapacityLine(OIDCAP, OIDBranch, LINEID, CreatedBy, CreatedDate) ");
                                                sbSAVE.Append("  VALUES('" + strCAP + "', '" + strBRANCH + "', '" + item["LineName"].ToString() + "', '" + strCREATE + "', GETDATE()) ");
                                                sbSAVE.Append(" END ");
                                                iCQC++;
                                            }

                                            if (strLINE == "")
                                            {
                                                sbSAVE.Append("DELETE FROM ProductionCapacityLine WHERE (OIDCAP = '" + strCAP + "') AND (OIDBranch = '" + strBRANCH + "')  ");
                                            }
                                            else
                                            {
                                                sbSAVE.Append("DELETE FROM ProductionCapacityLine WHERE (OIDCAP = '" + strCAP + "') AND (OIDBranch = '" + strBRANCH + "') AND (LINEID NOT IN (" + strLINE + "))  ");
                                            }
                                        }
                                    }


                                    if (sbSAVE.Length > 0)
                                    {
                                        //MessageBox.Show(sbSAVE.ToString());
                                        try
                                        {
                                            bool chkSAVECAPA = this.DBC.DBQuery(sbSAVE).runSQL();
                                            if (chkSAVECAPA == true)
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
                        catch (Exception)
                        { }
                    }
                }
            }
        }

        private void bbiExcel_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            string pathFile = new ObjSet.Folder(@"C:\MDS\Export\").GetPath() + "ProductionCapacityList_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
            gvCapacity.ExportToXlsx(pathFile);
            System.Diagnostics.Process.Start(pathFile);
        }

        private void txe1Hr_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txe1Day.Focus();
            }
        }

        private void txe1Day_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeOT.Focus();
            }
        }

        private void txeOT_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeCutting.Focus();
            }
        }

        private void txeCutting_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txePadPrint.Focus();
            }
        }

        private void txePadPrint_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeSewing.Focus();
            }
        }

        private void txeSewing_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txePacking.Focus();
            }
        }

        private void txePacking_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeStdTime.Focus();
            }
        }

        private void txeStdTime_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                dteStart.Focus();
            }
        }

        private void gvCapacity_RowClick(object sender, RowClickEventArgs e)
        {
            if (gvCapacity.IsFilterRow(e.RowHandle)) return;
        }

        private void gvCapacity_DoubleClick(object sender, EventArgs e)
        {
            DXMouseEventArgs ea = e as DXMouseEventArgs;
            GridView view = sender as GridView;
            GridHitInfo info = view.CalcHitInfo(ea.Location);
            if (info.InRow || info.InRowCell)
            {
                GridView gv = gvCapacity;
                string CUSID = gv.GetFocusedRowCellValue("CustomerID").ToString();
                string CATEID = gv.GetFocusedRowCellValue("CategoryID").ToString();
                string STYLEID = gv.GetFocusedRowCellValue("StyleID").ToString();

                lblStatus.Text = "* Edit Capacity";
                lblStatus.ForeColor = Color.Red;

                txeID.Text = gv.GetFocusedRowCellValue("CapacityID").ToString();

                slueCustomer.EditValue = CUSID;
                glueCategory.EditValue = CATEID;
                slueStyle.EditValue = STYLEID;
            }
        }

        private void bbiPrintPreview_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            gcCapacity.ShowPrintPreview();
        }

        private void bbiPrint_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            gcCapacity.Print();
        }

        private DateTime FIND_SUM_TIME(DateTime dtSUM, string strTIME)
        {
            string Minutes = "0";
            string Seconds = "0";

            if (strTIME.IndexOf('.') > -1)
            {
                string[] arrTIME = strTIME.Split('.');

                if (arrTIME.Length > 0)
                {
                    Minutes = arrTIME[0];
                    Seconds = arrTIME[1];
                    if (Seconds.Length == 1)
                        Seconds = (Convert.ToInt32(Seconds) * 10).ToString();
                }
            }
            else
                Minutes = strTIME;
           
            dtSUM = dtSUM.AddMinutes(Convert.ToInt32(Minutes)).AddSeconds(Convert.ToInt32(Seconds));

            return dtSUM;
        }

        private void SUM_TIME()
        {
            txeStdTime.Text = Convert.ToDateTime("01/01/2000 00:00:00").ToString("mm:ss");
            DateTime dtSUM = Convert.ToDateTime("01/01/2000 00:00:00");
            if (txeCutting.Text != "")
                dtSUM = FIND_SUM_TIME(dtSUM, txeCutting.Text.Trim());
            if (txePadPrint.Text != "")
                dtSUM = FIND_SUM_TIME(dtSUM, txePadPrint.Text.Trim());
            if (txeSewing.Text != "")
                dtSUM = FIND_SUM_TIME(dtSUM, txeSewing.Text.Trim());
            if (txePacking.Text != "")
                dtSUM = FIND_SUM_TIME(dtSUM, txePacking.Text.Trim());
            txeStdTime.Text = dtSUM.ToString("mm.ss");
        }

        private void txeCutting_Leave(object sender, EventArgs e)
        {
            if (txeCutting.Text != "")
            {
                DateTime dtSUM = Convert.ToDateTime("01/01/2000 00:00:00");
                dtSUM = FIND_SUM_TIME(dtSUM, txeCutting.Text.Trim());
                txeCutting.Text = Convert.ToDouble(dtSUM.ToString("mm.ss")).ToString("##.##");
            }
            SUM_TIME();
        }

        private void txePadPrint_Leave(object sender, EventArgs e)
        {
            if (txePadPrint.Text != "")
            {
                DateTime dtSUM = Convert.ToDateTime("01/01/2000 00:00:00");
                dtSUM = FIND_SUM_TIME(dtSUM, txePadPrint.Text.Trim());
                txePadPrint.Text = Convert.ToDouble(dtSUM.ToString("mm.ss")).ToString("##.##");
            }
            SUM_TIME();
        }

        private void txeSewing_Leave(object sender, EventArgs e)
        {
            if (txeSewing.Text != "")
            {
                DateTime dtSUM = Convert.ToDateTime("01/01/2000 00:00:00");
                dtSUM = FIND_SUM_TIME(dtSUM, txeSewing.Text.Trim());
                txeSewing.Text = Convert.ToDouble(dtSUM.ToString("mm.ss")).ToString("##.##");
            }
            SUM_TIME();
        }

        private void txePacking_Leave(object sender, EventArgs e)
        {
            if (txePacking.Text != "")
            {
                DateTime dtSUM = Convert.ToDateTime("01/01/2000 00:00:00");
                dtSUM = FIND_SUM_TIME(dtSUM, txePacking.Text.Trim());
                txePacking.Text = Convert.ToDouble(dtSUM.ToString("mm.ss")).ToString("##.##");
            }
            SUM_TIME();
        }

        private void gvCapacity_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator) e.Info.DisplayText = (e.RowHandle + 1).ToString();
            gvCapacity.IndicatorWidth = 40;
        }

      
    }
}