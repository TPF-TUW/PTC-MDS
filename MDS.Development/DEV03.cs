using System;
using System.Text;
using DBConnect;
using System.Windows.Forms;
using System.Globalization;
using System.Data;
using DevExpress.LookAndFeel;
using DevExpress.Utils;
using DevExpress.Utils.Extensions;
using System.Drawing;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Controls;
using DevExpress.XtraLayout.Utils;
using TheepClass;
using DBConnect;

namespace MDS.Development
{
    public partial class DEV03 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        cDatabase db;
        CultureInfo clinfo = new CultureInfo("en-US");
        DateTimeFormatInfo dtfinfo;
        goClass.dbConn db1 = new goClass.dbConn();
        goClass.ctool ct = new goClass.ctool();
        hardQuery q = new hardQuery();
        private Functionality.Function FUNC = new Functionality.Function();

        int chkReadWrite = 0;

        public LogIn UserLogin { get; set; }
        public int Company { get; set; }

        public string ConnectionString { get; set; }

        string CONNECT_STRING = "";
        DatabaseConnect DBC;

        public DEV03()
        {
            InitializeComponent();
        }

        private void NewData()
        {
            try
            {
                ClearSampleRequestDetail();
                ClearBOMDetail();
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


        }

        private void ClearBomList()
        {

        }
        private void ClearSampleRequestDetail()
        {
            gridControl2.DataSource = null;
            txtCreateBy.Text = "";
            txtCreateDate.Text = "";
            txtUpdateBy.Text = "";
            txtUpdateDate.Text = "";
        }
        private void ClearBOMDetail()
        {
            gridControl3.DataSource = null;

            txtBomNo.Text = "";
            txtReviseNo.Text = "";
            dtpLastDate.EditValue = DateTime.Today;
            txtSmplNo_Header.Text = "";
            txtPatternNo.Text = "";
            txtPatternSizeZone.Text = "";
            txtModelName.Text = "";
            txtStyleName.Text = "";
            txtCategory.Text = "";
            txtSeason_Header.Text = "";
            txtCustomer_Header.Text = "";
            txtFGProductCode.Text = "";
            txtColor.Text = "";
            txtSize.Text = "";
            gleUnit.Properties.DataSource = null;
            txtUnitCost.Text = "0";
            rdoStatus.SelectedIndex = -1;
            txtCostsheetNo.Text = "";
            treeBom.DataSource = null;
        }
        private void GetBOMList()
        {
            string strSQL = "EXEC spDEV03_GetBOM";
            DataTable dt = db.GetDataTable(strSQL);
            gridControl1.DataSource = dt;
            if (dt == null) return;
            gridView1.Columns["BOMNO"].Caption = "BOM No.";
            gridView1.Columns["REVISIONNO"].Caption = "Revise";
            gridView1.Columns["SMPLITEM"].Caption = "SMPL Item";
            gridView1.Columns["SMPLNO"].Caption = "SMPL No.";
            gridView1.Columns["SEASON"].Caption = "Season";
            gridView1.Columns["CUSTOMER"].Caption = "Customer";
            gridView1.Columns["ITEM"].Caption = "Item";
            gridView1.Columns["CATEGORYNAME"].Caption = "Category";
            gridView1.Columns["STYLENAME"].Caption = "Style";
            gridView1.Columns["SMPLPATTERNNO"].Caption = "Pattern No.";
            gridView1.Columns["USERNAME"].Caption = "Created By";
            gridView1.OptionsView.EnableAppearanceEvenRow = true;
            gridView1.OptionsView.EnableAppearanceOddRow = true;
            gridView1.OptionsView.ColumnAutoWidth = false;
            gridView1.BestFitColumns();
        }
        private void GetSampleRequest(int oidSMPL)
        {
            string strSQL = "EXEC spDEV03_GetSampleRequest " + oidSMPL;
            DataTable dt = db.GetDataTable(strSQL);
            gridControl2.DataSource = dt;
            if (dt == null) return;
            foreach (DataRow dr in dt.Rows)
            {
                txtCreateBy.Text = dr["CREATEDBY"].ToString();
                txtCreateDate.Text = dr["CREATEDDATE"].ToString();
                txtUpdateBy.Text = dr["UPDATEDBY"].ToString();
                txtUpdateDate.Text = dr["UPDATEDDATE"].ToString();
            }
            gridView2.Columns["OIDSMPLDT"].Caption = "ID";
            gridView2.Columns["SMPLITEM"].Caption = "Item No.";
            gridView2.Columns["COLORNAME"].Caption = "Color";
            gridView2.Columns["SIZENAME"].Caption = "Size";
            gridView2.Columns["OIDSMPLDT"].Visible = false;
            gridView2.Columns["OIDSMPL"].Visible = false;
            gridView2.Columns["OIDCOLOR"].Visible = false;
            gridView2.Columns["OIDSIZE"].Visible = false;
            gridView2.Columns["CREATEDBY"].Visible = false;
            gridView2.Columns["CREATEDDATE"].Visible = false;
            gridView2.Columns["UPDATEDBY"].Visible = false;
            gridView2.Columns["UPDATEDDATE"].Visible = false;
            gridView2.OptionsView.EnableAppearanceEvenRow = true;
            gridView2.OptionsView.EnableAppearanceOddRow = true;
            gridView2.OptionsView.ColumnAutoWidth = false;
            gridView2.BestFitColumns();

        }
        private void GetBOMDetail(int oidSMPL, int oidSize, int oidColor)
        {
            string strSQL = "EXEC spDEV03_GetBOMDetail " + oidSMPL + "," + oidSize + "," + oidColor;
            DataTable dt = db.GetDataTable(strSQL);
            if (dt == null) return;
            foreach (DataRow dr in dt.Rows)
            {
                txtBomNo.EditValue = dr["BOMNO"];
                txtReviseNo.EditValue = dr["REVISIONNO"];
                dtpLastDate.EditValue = dr["ISSUEDATE"] == System.DBNull.Value ? (DateTime?)null : (DateTime)dr["ISSUEDATE"];
                txtSmplNo_Header.EditValue = dr["SMPLNO"];
                txtPatternNo.EditValue = dr["SMPLPATTERNNO"];
                txtPatternSizeZone.EditValue = dr["PATTERNSIZEZONE"];
                txtItemNo.EditValue = dr["SMPLITEM"];
                txtModelName.EditValue = dr["MODELNAME"];
                txtStyleName.EditValue = dr["STYLENAME"];
                txtCategory.EditValue = dr["CATEGORYNAME"];
                txtSeason_Header.EditValue = dr["SEASON"];
                txtCustomer_Header.EditValue = dr["SHORTNAME"];

            }

        }





        private void MyStyleChanged(object sender, EventArgs e)
        {
            UserLookAndFeel userLookAndFeel = (UserLookAndFeel)sender;
            cUtility.SaveRegistry(@"Software\MDS", "SkinName", userLookAndFeel.SkinName);
            cUtility.SaveRegistry(@"Software\MDS", "SkinPalette", userLookAndFeel.ActiveSvgPaletteName);
        }
        private void DEV03_Load(object sender, EventArgs e)
        {
            UserLookAndFeel.Default.StyleChanged += MyStyleChanged;
            IniFile ini = new IniFile(@"\\192.168.101.3\Software_tuw\PTC-MDS\FileConfig\Configue.ini");
            db = new cDatabase("Server=" + ini.Read("Server", "ConnectionString") + ";uid=" + ini.Read("Uid", "ConnectionString") + ";pwd=" + ini.Read("Pwd", "ConnectionString") + ";database=" + ini.Read("Database", "ConnectionString"));
            dtfinfo = clinfo.DateTimeFormat;
            try
            {
                NewData();
                GetBOMList();
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


            ////***** SET CONNECT DB ********
            //if (this.ConnectionString != null)
            //{
            //    if (this.ConnectionString != "")
            //    {
            //        CONNECT_STRING = this.ConnectionString;
            //    }
            //}

            //this.DBC = new DatabaseConnect(CONNECT_STRING);

            //if (this.DBC.chkCONNECTION_STING() == false)
            //{
            //    this.DBC.setCONNECTION_STRING_INIFILE();
            //    if (this.DBC.chkCONNECTION_STING() == false)
            //    {
            //        return;
            //    }
            //}
            //new ObjDE.setDatabase(this.DBC);
            //******************************

            lblUser.Text = "Login : " + UserLogin.FullName;
            //StringBuilder sbSQL = new StringBuilder();
            //sbSQL.Append("SELECT TOP (1) ReadWriteStatus FROM FunctionAccess WHERE (OIDUser = '" + UserLogin.OIDUser + "') AND(FunctionNo = 'DEV03') ");
            //chkReadWrite = this.DBC.DBQuery(sbSQL).getInt();

            //if (chkReadWrite == 0)
            //{
            //    ribbonPageGroup1.Visible = false;
            //    //rpgManage.Visible = false;

            //    layoutControlItem29.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            //    //simpleButton2.Enabled = false;
            //    //simpleButton3.Enabled = false;
            //    //simpleButton4.Enabled = false;
            //    //sbColor.Enabled = false;
            //    //sbSize.Enabled = false;
            //    //sbFBColor.Enabled = false;
            //    //sbTempCode.Enabled = false;
            //    //sbMTColor.Enabled = false;
            //    //sbTempCodeMat.Enabled = false;
            //    //btnOpenImg_Main.Enabled = false;
            //    //sbDelete_S.Enabled = false;
            //    //sbClear.Enabled = false;
            //    //simpleButton5.Enabled = false;
            //    //sbDelete_F.Enabled = false;
            //    //sbMatClear.Enabled = false;
            //    //btnUploadMat.Enabled = false;
            //    //simpleButton1.Enabled = false;

            //    //sbUseFor.Enabled = false;
            //    //sbUnit.Enabled = false;

            //    //sbPart.Enabled = false;
            //    //sbFBSupplier.Enabled = false;

            //    //sbMTSupplier.Enabled = false;
            //}

            //LoadListBOM();
        }

        private void LoadListBOM()
        {
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("");
            new ObjDE.setGridControl(gridControl1, gridView1, sbSQL).getData(false, false, false, true);
        }

        private void LoadData()
        {
            //StringBuilder sbSQL = new StringBuilder();
            //sbSQL.Append("SELECT OIDPayment AS No, Name, Description, DuedateCalculation, Status, CreatedBy, CreatedDate ");
            //sbSQL.Append("FROM PaymentTerm ");
            //sbSQL.Append("ORDER BY OIDPayment ");
            //new ObjDevEx.setGridControl(gcPTerm, gvPTerm, sbSQL).getData(false, false, false, true);

        }
        private void bbiNew_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            tabbed_Master.SelectedTabPageIndex = 1;

        }

        private void gvGarment_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {

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
            //        if (new DBQuery(sbSQL).getString() != "")
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
            //        string strCHK = new DBQuery(sbSQL).getString();
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

        private void gvPTerm_RowStyle(object sender, RowStyleEventArgs e)
        {

        }

        private void bbiSave_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            //if (txeName.Text.Trim() == "")
            //{
            //    FUNC.msgWarning("Please name.");
            //    txeName.Focus();
            //}
            //else if (txeDescription.Text.Trim() == "")
            //{
            //    FUNC.msgWarning("Please input description.");
            //    txeDescription.Focus();
            //}
            //else
            //{
            //    if (FUNC.msgQuiz("Confirm save data ?") == true)
            //    {
            //        StringBuilder sbSQL = new StringBuilder();
            //        string strCREATE = "0";
            //        if (txeCREATE.Text.Trim() != "")
            //        {
            //            strCREATE = txeCREATE.Text.Trim();
            //        }

            //        bool chkGMP = chkDuplicate();
            //        if (chkGMP == true)
            //        {
            //            string Status = "NULL";
            //            if (rgStatus.SelectedIndex != -1)
            //            {
            //                Status = rgStatus.Properties.Items[rgStatus.SelectedIndex].Value.ToString();
            //            }

            //            if (lblStatus.Text == "* Add Payment Term")
            //            {
            //                sbSQL.Append("  INSERT INTO PaymentTerm(Name, Description, DueDateCalculation, Status, CreatedBy, CreatedDate) ");
            //                sbSQL.Append("  VALUES(N'" + txeName.Text.Trim().Replace("'", "''") + "', N'" + txeDescription.Text.Trim().Replace("'", "''") + "', N'" + txeDueDate.Text.Trim().Replace("'", "''") + "', " + Status + ", '" + strCREATE + "', GETDATE()) ");
            //            }
            //            else if (lblStatus.Text == "* Edit Payment Term")
            //            {
            //                sbSQL.Append("  UPDATE PaymentTerm SET ");
            //                sbSQL.Append("      Name=N'" + txeName.Text.Trim().Replace("'", "''") + "', ");
            //                sbSQL.Append("      Description=N'" + txeDescription.Text.Trim().Replace("'", "''") + "', ");
            //                sbSQL.Append("      DueDateCalculation=N'" + txeDueDate.Text.Trim().Replace("'", "''") + "', ");
            //                sbSQL.Append("      Status=" + Status + " ");
            //                sbSQL.Append("  WHERE(OIDPayment = '" + txeID.Text.Trim() + "') ");
            //            }

            //            //MessageBox.Show(sbSQL.ToString());
            //            if (sbSQL.Length > 0)
            //            {
            //                try
            //                {
            //                    bool chkSAVE = new DBQuery(sbSQL).runSQL();
            //                    if (chkSAVE == true)
            //                    {
            //                        FUNC.msgInfo("Save complete.");
            //                        bbiNew.PerformClick();
            //                    }
            //                }
            //                catch (Exception)
            //                { }
            //            }
            //        }
            //    }
            //}
        }

        private void bbiExcel_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            //string pathFile = new ObjSet.Folder(@"C:\MDS\Export\").GetPath() + "PaymentTermList_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
            //gvPTerm.ExportToXlsx(pathFile);
            //System.Diagnostics.Process.Start(pathFile);
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
        }

        private void bbiPrint_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            //gcPTerm.Print();
        }

        private void tabbed_Master_SelectedPageChanged(object sender, DevExpress.XtraLayout.LayoutTabPageChangedEventArgs e)
        {
            if (tabbed_Master.SelectedTabPageIndex == 1) //Entry
            {
                //q.get_sl_smplNo(sleSmplNoEntry);
                //q.get_gl_Branch(gleBranchEntry);
                //q.get_gl_Season(gleSeasonEntry);
                //q.get_sl_Customer(sleCustomerEntry);
                //q.get_gcListof_SMPL(gcListof_SMPL); gvListof_SMPL.OptionsBehavior.Editable = false;
                //txtCreateBy.EditValue = 0;
                //txtCreateDate.EditValue = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                //txtUpdateBy.EditValue = 0;
                //txtUpdateDate.EditValue = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                //// Header
                //txtBomNo.EditValue = q.get_running_BomNo(); txtBomNo.ReadOnly = true;
                //dtLastDate.EditValue = DateTime.Now;
                ////q.get_sl_StyleNmae(sl_StyleName);
                ////q.get_gl_Category(gl_Category);
                ////q.get_gl_Season(gl_Season_Header);
                ////q.get_sl_Customer(sl_Customer_Header);
                //q.get_sl_Color(sl_Color);
                //q.get_sl_Size(sl_Size);
                //q.get_gl_Unit(gl_Unit);
                //rdoStatus.SelectedIndex = 0;
            }
        }


        private void sleCustomerEntry_EditValueChanged(object sender, EventArgs e)
        {

        }

        private void ribbonControl_Click(object sender, EventArgs e)
        {

        }

        private void gridView1_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator) e.Info.DisplayText = (e.RowHandle + 1).ToString();
            gridView1.IndicatorWidth = 45;
        }
        private void gridView2_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator) e.Info.DisplayText = (e.RowHandle + 1).ToString();
            gridView2.IndicatorWidth = 45;
        }
        private void gridView1_ShowingEditor(object sender, System.ComponentModel.CancelEventArgs e)
        {
            e.Cancel = true;
        }
        private void gridView2_ShowingEditor(object sender, System.ComponentModel.CancelEventArgs e)
        {
            e.Cancel = true;
        }
        private void gridView1_DoubleClick(object sender, EventArgs e)
        {
            DXMouseEventArgs ea = e as DXMouseEventArgs;
            GridView view = sender as GridView;
            GridHitInfo info = view.CalcHitInfo(ea.Location);
            if (info.InRow || info.InRowCell)
            {
                //string colCaption = info.Column == null ? "N/A" : info.Column.GetCaption();
                ClearSampleRequestDetail();
                ClearBOMDetail();
                GetSampleRequest(Convert.ToInt32(view.GetRowCellValue(info.RowHandle, "OIDSMPL")));
                tabbed_Master.SelectedTabPageIndex = 1;
                //MessageBox.Show(string.Format("DoubleClick on row: {0}, column: {1}.", info.RowHandle, colCaption));
            }
        }
        private void gridView2_RowCellClick(object sender, RowCellClickEventArgs e)
        {
            try
            {
                ClearBOMDetail();
                GetBOMDetail(Convert.ToInt32(gridView2.GetRowCellValue(e.RowHandle, "OIDSMPL")), Convert.ToInt32(gridView2.GetRowCellValue(e.RowHandle, "OIDSIZE")), Convert.ToInt32(gridView2.GetRowCellValue(e.RowHandle, "OIDCOLOR")));
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}