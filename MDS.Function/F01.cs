using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using DBConnect;
using DevExpress.LookAndFeel;
using DevExpress.Utils.Drawing.Helpers;
using DevExpress.Utils.Extensions;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Views.Grid;
using TheepClass;
using DevExpress.Utils;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;

namespace MDS.Function
{
    public partial class F01 : DevExpress.XtraBars.ToolbarForm.ToolbarForm
    {
        private Functionality.Function FUNC = new Functionality.Function();
        public LogIn UserLogin { get; set; }
        public int Company { get; set; }
        public string ConnectionString { get; set; }
        string CONNECT_STRING = "";
        DatabaseConnect DBC;


        public F01()
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
            sbSQL.Append("SELECT TOP (1) ReadWriteStatus FROM FunctionAccess WHERE (OIDUser = '" + UserLogin.OIDUser + "') AND(FunctionNo = 'F01') ");
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

            glueCREATE.EditValue = UserLogin.OIDUser;

            LoadData();
            NewData();
        }

        private void LoadData()
        {
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT OIDCOMPANY AS ID, Code AS [Company Code], EngName AS [Company Name], EngAddress1 AS Address1, EngAddress2 AS Address2, EngAddress3 AS Address3, Telephone, FaxNo AS [Fax.], TaxID AS [Tax ID.], ");
            sbSQL.Append("       BranchNo AS[Branch No.], THName AS ชื่อบริษัท, THAddress1 AS ที่อยู่1, THAddress2 AS ที่อยู่2, THAddress3 AS ที่อยู่3, Status AS State, CASE WHEN Status = 0 THEN 'Non Active' ELSE CASE WHEN Status = 1 THEN 'Active' ELSE '' END END AS [Status], CreatedBy, CreatedDate ");
            sbSQL.Append("FROM   Company ");
            //sbSQL.Append("WHERE (Status = 1) ");
            sbSQL.Append("ORDER BY ID ");
            new ObjDE.setGridControl(gcCompany, gvCompany, sbSQL).getData(false, false, false, true);
            gvCompany.Columns[0].Visible = false; //ID
            gvCompany.Columns[14].Visible = false; //Status

            gvCompany.Columns["CreatedBy"].Visible = false;
            gvCompany.Columns["CreatedDate"].Visible = false;
        }

        private void NewData()
        {
            txeID.Text = this.DBC.DBQuery("SELECT CASE WHEN ISNULL(MAX(OIDCOMPANY), '') = '' THEN 1 ELSE MAX(OIDCOMPANY) + 1 END AS NewNo FROM Company").getString();
 
            lblStatus.Text = "* Add Company";
            lblStatus.ForeColor = Color.Green;

            txeCode.Text = "";
            txeName.Text = "";
            txeAddr1.Text = "";
            txeAddr2.Text = "";
            txeAddr3.Text = "";
            txeTel.Text = "";
            txeFax.Text = "";
            txeTax.Text = "";
            txeBranchNo.Text = "";

            txeNameTH.Text = "";
            txeAddr1TH.Text = "";
            txeAddr2TH.Text = "";
            txeAddr3TH.Text = "";
            txeTelTH.Text = "";
            txeFaxTH.Text = "";
            txeTaxTH.Text = "";
            txeBranchNoTH.Text = "";

            rgStatus.EditValue = 1;

            glueCREATE.EditValue = UserLogin.OIDUser;
            txeCDATE.EditValue = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
            txeCode.Focus();
        }

        private void bbiNew_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            LoadData();
            NewData();
        }

        private void gvGarment_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {
            
        }

        private bool chkDuplicateNo()
        {
            bool chkDup = true;
            if (txeCode.Text != "")
            {
                txeCode.Text = txeCode.Text.Trim();
                if (txeCode.Text.Trim() != "" && lblStatus.Text == "* Add Company")
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT TOP(1) Code FROM Company WHERE (Code = N'" + txeCode.Text.Trim() + "') ");
                    if (this.DBC.DBQuery(sbSQL).getString() != "")
                    {
                        chkDup = false;
                    }
                }
                else if (txeCode.Text.Trim() != "" && lblStatus.Text == "* Edit Company")
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT TOP(1) OIDCOMPANY ");
                    sbSQL.Append("FROM Company ");
                    sbSQL.Append("WHERE (Code = N'" + txeCode.Text.Trim() + "') ");
                    string strCHK = this.DBC.DBQuery(sbSQL).getString();
                    if (strCHK != "" && strCHK != txeID.Text.Trim())
                    {
                        chkDup = false;
                    }
                }
            }
            return chkDup;
        }

        private void bbiSave_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (txeCode.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input company code.");
                txeCode.Focus();
            }
            else if (txeName.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input company name.");
                txeName.Focus();
            }
            else
            {
                bool chkGMP = chkDuplicateNo();
                if (chkGMP == true)
                {
                    StringBuilder sbSQL = new StringBuilder();
                    string strCREATE = UserLogin.OIDUser.ToString() != "" ? UserLogin.OIDUser.ToString() : "0";
       
                    if (FUNC.msgQuiz("Confirm save data ?") == true)
                    {
                        string Status = "NULL";
                        if (rgStatus.SelectedIndex != -1)
                        {
                            Status = rgStatus.Properties.Items[rgStatus.SelectedIndex].Value.ToString();
                        }

                        if (lblStatus.Text == "* Add Company")
                        {
                            sbSQL.Append("  INSERT INTO Company(Code, EngName, EngAddress1, EngAddress2, EngAddress3, THName, THAddress1, THAddress2, THAddress3, Telephone, FaxNo, TaxID, BranchNo, Status, CreatedBy, CreatedDate) ");
                            sbSQL.Append("  VALUES(N'" + txeCode.Text.Trim().Replace("'", "''") + "', N'" + txeName.Text.Trim().Replace("'", "''") + "', N'" + txeAddr1.Text.Trim().Replace("'", "''") + "', N'" + txeAddr2.Text.Trim().Replace("'", "''") + "', N'" + txeAddr3.Text.Trim().Replace("'", "''") + "', N'" + txeNameTH.Text.Trim().Replace("'", "''") + "', N'" + txeAddr1TH.Text.Trim().Replace("'", "''") + "', N'" + txeAddr2TH.Text.Trim().Replace("'", "''") + "', N'" + txeAddr3TH.Text.Trim().Replace("'", "''") + "', N'" + txeTel.Text.Trim() + "', N'" + txeFax.Text.Trim() + "', N'" + txeTax.Text.Trim() + "', N'" + txeBranchNo.Text.Trim() + "', " + Status + ", '" + strCREATE + "', GETDATE()) ");
                        }
                        else if (lblStatus.Text == "* Edit Company")
                        {
                            sbSQL.Append("  UPDATE Company SET ");
                            sbSQL.Append("      Code = N'" + txeCode.Text.Trim().Replace("'", "''") + "', ");
                            sbSQL.Append("      EngName = N'" + txeName.Text.Trim().Replace("'", "''") + "', ");
                            sbSQL.Append("      EngAddress1 = N'" + txeAddr1.Text.Trim().Replace("'", "''") + "', ");
                            sbSQL.Append("      EngAddress2 = N'" + txeAddr2.Text.Trim().Replace("'", "''") + "', ");
                            sbSQL.Append("      EngAddress3 = N'" + txeAddr3.Text.Trim().Replace("'", "''") + "', ");
                            sbSQL.Append("      THName = N'" + txeNameTH.Text.Trim().Replace("'", "''") + "', ");
                            sbSQL.Append("      THAddress1 = N'" + txeAddr1TH.Text.Trim().Replace("'", "''") + "', ");
                            sbSQL.Append("      THAddress2 = N'" + txeAddr2TH.Text.Trim().Replace("'", "''") + "', ");
                            sbSQL.Append("      THAddress3 = N'" + txeAddr3TH.Text.Trim().Replace("'", "''") + "', ");
                            sbSQL.Append("      Telephone = N'" + txeTel.Text.Trim() + "', ");
                            sbSQL.Append("      FaxNo = N'" + txeFax.Text.Trim() + "', ");
                            sbSQL.Append("      TaxID = N'" + txeTax.Text.Trim() + "', ");
                            sbSQL.Append("      BranchNo = N'" + txeBranchNo.Text.Trim() + "', ");
                            sbSQL.Append("      Status = " + Status + " ");
                            sbSQL.Append("  WHERE(OIDCOMPANY = '" + txeID.Text.Trim() + "') ");
                        }

                        //MessageBox.Show(sbSQL.ToString());
                        if (sbSQL.Length > 0)
                        {
                            try
                            {
                                bool chkSAVE = this.DBC.DBQuery(sbSQL).runSQL();
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
                else
                {
                    txeCode.Text = "";
                    txeCode.Focus();
                    FUNC.msgWarning("Duplicate company code. !! Please Change.");
                }
            }
        }

        private void bbiExcel_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            string pathFile = new ObjSet.Folder(@"C:\MDS\Export\").GetPath() + "CompanyList_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
            gvCompany.ExportToXlsx(pathFile);
            System.Diagnostics.Process.Start(pathFile);
        }

        private void bbiPrintPreview_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            gcCompany.ShowPrintPreview();
        }

        private void bbiPrint_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            gcCompany.Print();
        }

        private void F01_Shown(object sender, EventArgs e)
        {
            txeTel.Properties.Mask.EditMask = @"[0-9,\-(␣)^ ]*";
            txeTel.Properties.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.RegEx;

            txeFax.Properties.Mask.EditMask = @"[0-9,\-(␣)^ ]*";
            txeFax.Properties.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.RegEx;

            txeTax.Properties.Mask.EditMask = "[0-9]*";
            txeTax.Properties.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.RegEx;

            txeBranchNo.Properties.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.Numeric;
            txeBranchNo.Properties.Mask.EditMask = "00000";
            txeBranchNo.Properties.Mask.UseMaskAsDisplayFormat = true;

            txeBranchNoTH.Properties.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.Numeric;
            txeBranchNoTH.Properties.Mask.EditMask = "00000";
            txeBranchNoTH.Properties.Mask.UseMaskAsDisplayFormat = true;

            txeCode.Focus();
        }

        private void txeTel_KeyUp(object sender, KeyEventArgs e)
        {
            txeTelTH.Text = txeTel.Text;
        }

        private void txeFax_KeyUp(object sender, KeyEventArgs e)
        {
            txeFaxTH.Text = txeFax.Text;
        }

        private void txeTax_KeyUp(object sender, KeyEventArgs e)
        {
            txeTaxTH.Text = txeTax.Text;
        }

        private void txeBranchNo_KeyUp(object sender, KeyEventArgs e)
        {
            txeBranchNoTH.Text = txeBranchNo.Text;
        }

        private void txeCode_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeName.Focus();
            }
        }

        private void txeCode_Leave(object sender, EventArgs e)
        {
            if (txeCode.Text.Trim() != "")
            {
                txeCode.Text = txeCode.Text.ToUpper().Trim();
                bool chkDup = chkDuplicateNo();
                if (chkDup == true)
                {
                    txeName.Focus();
                }
                else
                {
                    txeCode.Text = "";
                    txeCode.Focus();
                    FUNC.msgWarning("Duplicate company code. !! Please Change.");

                }
            }
        }

        private void gvCompany_DoubleClick(object sender, EventArgs e)
        {
            DXMouseEventArgs ea = e as DXMouseEventArgs;
            GridView view = sender as GridView;
            GridHitInfo info = view.CalcHitInfo(ea.Location);
            if (info.InRow || info.InRowCell)
            {
                GridView gv = gvCompany;
                lblStatus.Text = "* Edit Company";
                lblStatus.ForeColor = Color.Red;

                txeID.Text = gv.GetFocusedRowCellValue("ID").ToString();
                txeCode.Text = gv.GetFocusedRowCellValue("Company Code").ToString();
                txeName.Text = gv.GetFocusedRowCellValue("Company Name").ToString();
                txeAddr1.Text = gv.GetFocusedRowCellValue("Address1").ToString();
                txeAddr2.Text = gv.GetFocusedRowCellValue("Address2").ToString();
                txeAddr3.Text = gv.GetFocusedRowCellValue("Address3").ToString();

                txeNameTH.Text = gv.GetFocusedRowCellValue("ชื่อบริษัท").ToString();
                txeAddr1TH.Text = gv.GetFocusedRowCellValue("ที่อยู่1").ToString();
                txeAddr2TH.Text = gv.GetFocusedRowCellValue("ที่อยู่2").ToString();
                txeAddr3TH.Text = gv.GetFocusedRowCellValue("ที่อยู่3").ToString();

                txeTel.Text = gv.GetFocusedRowCellValue("Telephone").ToString();
                txeFax.Text = gv.GetFocusedRowCellValue("Fax.").ToString();
                txeTax.Text = gv.GetFocusedRowCellValue("Tax ID.").ToString();
                txeBranchNo.Text = gv.GetFocusedRowCellValue("Branch No.").ToString();

                txeTelTH.Text = txeTel.Text;
                txeFaxTH.Text = txeFax.Text;
                txeTaxTH.Text = txeTax.Text;
                txeBranchNoTH.Text = txeBranchNo.Text;

                rgStatus.EditValue = Convert.ToInt32(gv.GetFocusedRowCellValue("State").ToString()); 

                string CreatedBy = gv.GetFocusedRowCellValue("CreatedBy").ToString() == null ? "" : gv.GetFocusedRowCellValue("CreatedBy").ToString();
                glueCREATE.EditValue = CreatedBy;
                txeCDATE.Text = gv.GetFocusedRowCellValue("CreatedDate").ToString();
            }


        }

        private void gvCompany_RowCellClick(object sender, RowCellClickEventArgs e)
        {
            if (gvCompany.IsFilterRow(e.RowHandle)) return;
        }

        private void gvCompany_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator) e.Info.DisplayText = (e.RowHandle + 1).ToString();
            gvCompany.IndicatorWidth = 40;
        }
    }
}