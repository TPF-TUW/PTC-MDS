using System;
using System.Text;
using DBConnect;
using System.Windows.Forms;
using System.Globalization;
using System.Data;
using DevExpress.LookAndFeel;
using DevExpress.Utils.Extensions;
using System.Drawing;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Controls;
using DevExpress.XtraLayout.Utils;
using TheepClass;

namespace MDS.Development
{
    public partial class DEV03 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        cDatabase db;
        CultureInfo clinfo = new CultureInfo("en-US");
        DateTimeFormatInfo dtfinfo;
        //goClass.dbConn db1   = new goClass.dbConn();
        //goClass.ctool ct    = new goClass.ctool();
        //hardQuery q         = new hardQuery();
        private Functionality.Function FUNC = new Functionality.Function();
        
        public LogIn UserLogin { get; set; }
        //public int Company { get; set; }

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
                //gleBranchEntry.Properties.DataSource = null;
                //gleSeasonEntry.Properties.DataSource = null;
                //sleCustomerEntry.Properties.DataSource = null;
                //sleSmplNoEntry.Properties.DataSource = null;

                //gridControl2.DataSource = null;

                //txtCreateBy.Text = "";
                //txtCreateDate.Text = "";
                //txtUpdateBy.Text = "";
                //txtUpdateDate.Text = "";

                //gridControl3.DataSource = null;

                //txtBomNo.Text = "";
                //txtReviseNo.Text = "";
                //dtpLastDate.EditValue = DateTime.Today;
                //txtSmplNo_Header.Text = "";
                //txtPatternNo.Text = "";
                //txtPatternSizeZone.Text = "";
                //txtModelName.Text = "";
                //txtStyleName.Text = "";
                //txtCategory.Text = "";
                //txtSeason_Header.Text = "";
                //txtCustomer_Header.Text = "";
                //txtFGProductCode.Text = "";
                //txtColor.Properties.DataSource = null;
                //txtSize.Properties.DataSource = null;
                //gleUnit.Properties.DataSource = null;
                //txtUnitCost.Text = "0";
                //optStatus.SelectedIndex = -1 ;
                //txtCostsheetNo.Text = "";
                //treeBom.DataSource = null;

            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);                
            }

            //txeName.Text = "";
            //lblStatus.Text = "* Add Payment Term";
            //lblStatus.ForeColor = Color.Green;

            //txeID.Text = new DBQuery("SELECT CASE WHEN ISNULL(MAX(OIDPayment), '') = '' THEN 1 ELSE MAX(OIDPayment) + 1 END AS NewNo FROM PaymentTerm").getString();
            //txeDescription.Text = "";
            //txeDueDate.Text = "";
            //rgStatus.SelectedIndex = -1;

            //txeCREATE.Text = "0";
            //txeDATE.Text = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");

            //////txeID.Focus();
        }
        private void SaveData()
        { 
        
        }

        private void ClearBOMList() 
        {
            gridControl1.DataSource = null;
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
            //Header
            txtBomNo.Text = "";
            txtReviseNo.Text = "";
            dtpLastDate.EditValue = DateTime.Today;
            txtSmplNo_Header.Text = "";
            txtPatternNo.Text = "";
            txtPatternSizeZone.Text = "";
            txtItemNo.Text = "";
            txtModelName.Text = "";
            txtStyleName.Text = "";
            txtCategory.Text = "";
            txtSeason_Header.Text = "";
            txtCustomer_Header.Text = "";
            txtFGProductCode.Text = "";
            txtColor.Text = "";
            txtSize.Text = "";
            gleUnit.EditValue = null;
            txtUnitCost.Text = "";
            optStatus.SelectedIndex = -1;
            txtCostsheetNo.Text = "";
            treeBom.DataSource = null;
            //Details
            txtListNo.Text = "";
            gleMaterialType.EditValue = null;
            sleItemNo.EditValue = null;
            sleMatColor.EditValue = null;
            glematSize.EditValue = null;
            txtComposition.Text = "";
            gleCurrency.EditValue = null;
            txtPrice.Text = "";
            txtConsumption.Text = "";
            txtCost.Text = "";
            sleVendor.EditValue = null;
            txtVendMatCode.Text = "";
            txtSmplLotNo.Text = "";
            gleWorkStation.EditValue = null;
            txtMatLoss.Text = "";
            gleMatUnit.EditValue = null;

        }
        private void GetBOMList()
        {
            string strSQL = "EXEC SPDEV03_GETBOM";
            DataTable dt = db.GetDataTable(strSQL);
            gridControl1.DataSource = dt;
            gridView1.OptionsView.EnableAppearanceEvenRow = true;
            gridView1.OptionsView.EnableAppearanceOddRow = true;
            gridView1.OptionsView.ColumnAutoWidth = false;
            gridView1.BestFitColumns();

        }
        //private void GetBranch()
        //{
        //    string strSQL = "SELECT OIDBRANCH,NAME FROM BRANCHS";
        //    DataTable dt= db.GetDataTable(strSQL);
        //    gleBranchEntry.Properties.DataSource = dt;
        //    gleBranchEntry.Properties.DisplayMember = "NAME";
        //    gleBranchEntry.Properties.ValueMember = "OIDBRANCH";
        //    gleBranchEntry.Properties.PopulateViewColumns();
        //    gleBranchEntry.Properties.View.Columns["OIDBRANCH"].Visible = false;
        //    gleBranchEntry.Properties.BestFitMode = BestFitMode.BestFitResizePopup;
        //}
        //private void GetSeason(int oidBranch)
        //{
        //    string strSQL = "SELECT DISTINCT SEASON FROM SMPLREQUEST WHERE OIDBRANCH="+oidBranch;
        //    DataTable dt = db.GetDataTable(strSQL);
        //    gleSeasonEntry.Properties.DataSource = dt;
        //    gleSeasonEntry.Properties.DisplayMember = "SEASON";
        //    gleSeasonEntry.Properties.ValueMember = "SEASON";
        //    gleSeasonEntry.Properties.BestFitMode = BestFitMode.BestFitResizePopup;
        //}
        //private void GetCustomer(int oidBranch,string season)
        //{
        //    string strSQL = "SELECT OIDCUST,NAME FROM CUSTOMER WHERE OIDBRANCH="+oidBranch+" AND SEASON='"+season+"'";
        //    DataTable dt = db.GetDataTable(strSQL);
        //    sleCustomerEntry.Properties.DataSource = dt;
        //    sleCustomerEntry.Properties.DisplayMember = "NAME";
        //    sleCustomerEntry.Properties.ValueMember = "OIDCUST";
        //    sleCustomerEntry.Properties.PopulateViewColumns();
        //    sleCustomerEntry.Properties.View.Columns["OIDCUST"].Visible = false;
        //    sleCustomerEntry.Properties.BestFitMode = BestFitMode.BestFitResizePopup;
        //}
        //private void GetSampleRequest(int oidBranch,string season,int oidCust)
        //{
        //    string strSQL = "SELECT ";
        //}
        private void MyStyleChanged(object sender, EventArgs e)
        {
            UserLookAndFeel userLookAndFeel = (UserLookAndFeel)sender;
            cUtility.SaveRegistry(@"Software\MDS", "SkinName", userLookAndFeel.SkinName);
            cUtility.SaveRegistry(@"Software\MDS", "SkinPalette", userLookAndFeel.ActiveSvgPaletteName);
        }
        
        private void XtraForm1_Load(object sender, EventArgs e)
        {
            UserLookAndFeel.Default.StyleChanged += MyStyleChanged;
            IniFile ini = new IniFile(@"\\192.168.101.3\Software_tuw\PTC-MDS\FileConfig\Configue.ini");
            db = new cDatabase("Server=" + ini.Read("Server", "ConnectionString") + ";uid=" + ini.Read("Uid", "ConnectionString") + ";pwd=" + ini.Read("Pwd", "ConnectionString") + ";database=" + ini.Read("Database", "ConnectionString"));
            dtfinfo = clinfo.DateTimeFormat;
            try
            {
                NewData();
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            //bbiNew.PerformClick(); 

            // Set Tabbed
            tabbed_Master.SelectedTabPageIndex = 0;
            tabbedBom.SelectedTabPageIndex = 0;

            //q.get_sl_smplNo(sl_smplNo);
            //q.get_sl_Customer(sl_Customer);
            //q.get_gl_Season(gl_Season);
            q.get_gcListof_Bom(gridControl1); gridView1.OptionsBehavior.Editable = false;
            GetBranch();
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

        private void gleBranchEntry_EditValueChanged(object sender, EventArgs e)
        {
            try
            {
                if (gleBranchEntry.EditValue == null)
                    return;
                else
                    GetSeason((int)gleBranchEntry.EditValue);
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void gleSeasonEntry_EditValueChanged(object sender, EventArgs e)
        {
            try
            {
                if (gleSeasonEntry.EditValue == null)
                    return;
                else
                    GetCustomer((int)gleBranchEntry.EditValue, gleSeasonEntry.EditValue.ToString());
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void sleCustomerEntry_EditValueChanged(object sender, EventArgs e)
        {

        }
    }
}