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
    public partial class F02 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private Functionality.Function FUNC = new Functionality.Function();
        private string dbBranch = "Branchs";
        public LogIn UserLogin { get; set; }

        public string ConnectionString { get; set; }
        string CONNECT_STRING = "";
        DatabaseConnect DBC;

        public F02()
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
            sbSQL.Append("SELECT TOP (1) ReadWriteStatus FROM FunctionAccess WHERE (OIDUser = '" + UserLogin.OIDUser + "') AND(FunctionNo = 'F02') ");
            int chkReadWrite = this.DBC.DBQuery(sbSQL).getInt();
            if (chkReadWrite == 0)
                ribbonPageGroup1.Visible = false;

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
            sbSQL.Append("SELECT Code AS [Company Code], EngName AS [Company Name (En)], THName AS [Company Name (Th)], OIDCOMPANY AS ID ");
            sbSQL.Append("FROM Company ");
            sbSQL.Append("ORDER BY ID ");
            new ObjDE.setGridLookUpEdit(glueCompany, sbSQL, "Company Code", "ID").getData();

            sbSQL.Clear();
            sbSQL.Append("SELECT 0 AS ID, 'Branch' AS Type ");
            sbSQL.Append("UNION ALL ");
            sbSQL.Append("SELECT 1 AS ID, 'Branch Sub Contract' AS Type ");
            sbSQL.Append("ORDER BY ID");
            new ObjDE.setGridLookUpEdit(glueBranchType, sbSQL, "Type", "ID").getData();

            LoadBranch();
        }

        private void LoadBranch()
        {
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT BN.OIDBranch AS [OID Branch], BN.Code AS [Branch No.], BN.Name AS [Branch Name], BN.OIDCOMPANY, CP.Code AS [Company Code], CP.EngName AS [Company Name (En)], CP.THName AS [Company Name (Th)], ");
            sbSQL.Append("       BN.BranchType AS [Branch Type], CASE WHEN BN.BranchType = 0 THEN 'Branch' ELSE CASE WHEN BN.BranchType = 1 THEN 'Branch Sub Contract' ELSE '' END END AS [Type], BN.CreateBy AS[Created By], BN.CreateDate AS[Created Date] ");
            sbSQL.Append("FROM   " + this.dbBranch + " AS BN LEFT OUTER JOIN ");
            sbSQL.Append("       Company AS CP ON BN.OIDCOMPANY = CP.OIDCOMPANY ");
            sbSQL.Append("WHERE (BN.OIDBranch <> '') ");
            if (glueCompany.Text.Trim() != "")
                sbSQL.Append("AND (BN.OIDCOMPANY = '" + glueCompany.EditValue.ToString() + "') ");
            sbSQL.Append("ORDER BY [OID Branch] ");
            new ObjDE.setGridControl(gcBranch, gvBranch, sbSQL).getData(false, false, false, true);
            gvBranch.Columns[0].Visible = false;
            gvBranch.Columns[3].Visible = false; //OIDCOMPANY
            gvBranch.Columns[7].Visible = false; //Branch Type

            gvBranch.Columns["Created By"].Visible = false;
            gvBranch.Columns["Created Date"].Visible = false;
        }

        private void NewData()
        {
            lblStatus.Text = "* Add Branch";
            lblStatus.ForeColor = Color.Green;

            txeID.Text = this.DBC.DBQuery("SELECT CASE WHEN ISNULL(MAX(OIDBranch), '') = '' THEN 1 ELSE MAX(OIDBranch) + 1 END AS NewNo FROM " + this.dbBranch).getString();
            glueCompany.EditValue = "";
            txeBranchNo.Text = "";
            txeBranchName.Text = "";
            glueBranchType.EditValue = "";

            glueCREATE.EditValue = UserLogin.OIDUser;
            txeCDATE.Text = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");

            ////txeID.Focus();
        }

        private void bbiNew_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            LoadData();
            NewData();
        }

        private void gvGarment_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {
            
        }


        private bool chkDuplicate()
        {
            bool chkDup = true;
            string Company = "";
            if (glueCompany.Text.Trim() != "")
            {
                Company = glueCompany.EditValue.ToString();
            }
            
            if (lblStatus.Text == "* Add Branch")
            {
                if (txeBranchNo.Text.Trim() != "" || txeBranchName.Text.Trim() != "")
                {
                    StringBuilder sbSQL = new StringBuilder();
                    if (txeBranchNo.Text.Trim() != "" && txeBranchNo.Text.Trim() != "00000" && chkDup == true)
                    {
                        sbSQL.Clear();
                        sbSQL.Append("SELECT TOP(1) Code FROM " + this.dbBranch + " WHERE (OIDCOMPANY = '" + Company + "') AND (Code = N'" + txeBranchNo.Text.Trim() + "') ");
                        string chkNo = this.DBC.DBQuery(sbSQL).getString();
                        if (chkNo != "")
                        {
                            txeBranchNo.Text = "";
                            txeBranchNo.Focus();
                            chkDup = false;
                            FUNC.msgWarning("Duplicate branch no. !! Please Change.");
                        }
                    }

                    if (txeBranchName.Text.Trim() != "" && chkDup == true)
                    {
                        sbSQL.Clear();
                        sbSQL.Append("SELECT TOP(1) Code FROM " + this.dbBranch + " WHERE (OIDCOMPANY = '" + Company + "') AND (Name = N'" + txeBranchName.Text.Trim() + "') ");
                        string chkNo = this.DBC.DBQuery(sbSQL).getString();
                        if (chkNo != "")
                        {
                            txeBranchName.Text = "";
                            txeBranchName.Focus();
                            chkDup = false;
                            FUNC.msgWarning("Duplicate branch name. !! Please Change.");
                        }
                    }

                }
            }
            else if (lblStatus.Text == "* Edit Branch")
            {
                if (txeBranchNo.Text.Trim() != "" || txeBranchName.Text.Trim() != "")
                {
                    StringBuilder sbSQL = new StringBuilder();
                    if (txeBranchNo.Text.Trim() != "" && txeBranchNo.Text.Trim() != "00000" && chkDup == true)
                    {
                        sbSQL.Clear();
                        sbSQL.Append("SELECT TOP(1) OIDBranch FROM " + this.dbBranch + " WHERE (OIDCOMPANY = '" + Company + "') AND (Code = N'" + txeBranchNo.Text.Trim() + "') ");
                        string chkNo = this.DBC.DBQuery(sbSQL).getString();
                        if (chkNo != "" && chkNo != txeID.Text.Trim())
                        {
                            txeBranchNo.Text = "";
                            txeBranchNo.Focus();
                            chkDup = false;
                            FUNC.msgWarning("Duplicate branch no. !! Please Change.");
                        }
                    }

                    if (txeBranchName.Text.Trim() != "" && chkDup == true)
                    {
                        sbSQL.Clear();
                        sbSQL.Append("SELECT TOP(1) OIDBranch FROM " + this.dbBranch + " WHERE (OIDCOMPANY = '" + Company + "') AND (Name = N'" + txeBranchName.Text.Trim() + "') ");
                        string chkNo = this.DBC.DBQuery(sbSQL).getString();
                        if (chkNo != "" && chkNo != txeID.Text.Trim())
                        {
                            txeBranchName.Text = "";
                            txeBranchName.Focus();
                            chkDup = false;
                            FUNC.msgWarning("Duplicate branch name. !! Please Change.");
                        }
                    }

                }
            }

            return chkDup;
        }

        private void bbiSave_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (glueCompany.Text.Trim() == "")
            {
                FUNC.msgWarning("Please select company.");
                glueCompany.Focus();
            }
            else if (txeBranchNo.Text.Trim() == "" || txeBranchNo.Text.Trim() == "00000")
            {
                FUNC.msgWarning("Please input branch no.");
                txeBranchNo.Focus();
            }
            else if (txeBranchName.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input branch name.");
                txeBranchName.Focus();
            }
            else if (glueBranchType.Text.Trim() == "")
            {
                FUNC.msgWarning("Please select branch type.");
                glueBranchType.Focus();
            }
            else
            {
                bool chkGMP = chkDuplicate();
                if (chkGMP == true)
                {
                    StringBuilder sbSQL = new StringBuilder();
                    string strCREATE = UserLogin.OIDUser.ToString() != "" ? UserLogin.OIDUser.ToString() : "0";

                    if (FUNC.msgQuiz("Confirm save data ?") == true)
                    {
                        if (lblStatus.Text == "* Add Branch")
                        {
                            sbSQL.Append("  INSERT INTO " + this.dbBranch + "(Code, Name, OIDCOMPANY, BranchType, CreateBy, CreateDate) ");
                            sbSQL.Append("  VALUES(N'" + txeBranchNo.Text.Trim() + "', N'" + txeBranchName.Text.Trim().Replace("'", "''") + "', '" + glueCompany.EditValue.ToString() + "', '" + glueBranchType.EditValue.ToString() + "', '" + strCREATE + "', GETDATE()) ");
                        }
                        else if (lblStatus.Text == "* Edit Branch")
                        {
                            sbSQL.Append("  UPDATE " + this.dbBranch + " SET ");
                            sbSQL.Append("      Code = N'" + txeBranchNo.Text.Trim() + "', ");
                            sbSQL.Append("      Name = N'" + txeBranchName.Text.Trim().Replace("'", "''") + "', ");
                            sbSQL.Append("      OIDCOMPANY = '" + glueCompany.EditValue.ToString() + "', ");
                            sbSQL.Append("      BranchType = '" + glueBranchType.EditValue.ToString() + "' ");
                            sbSQL.Append("  WHERE (OIDBranch = '" + txeID.Text.Trim() + "') ");
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
            }
        }

        private void bbiExcel_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            string pathFile = new ObjSet.Folder(@"C:\MDS\Export\").GetPath() + "BranchList_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
            gvBranch.ExportToXlsx(pathFile);
            System.Diagnostics.Process.Start(pathFile);
        }

        private void bbiPrintPreview_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            gcBranch.ShowPrintPreview();
        }

        private void bbiPrint_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            gcBranch.Print();
        }

        private void F02_Shown(object sender, EventArgs e)
        {
            txeBranchNo.Properties.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.Numeric;
            txeBranchNo.Properties.Mask.EditMask = "00000";
            txeBranchNo.Properties.Mask.UseMaskAsDisplayFormat = true;
        }

        private void glueCompany_EditValueChanged(object sender, EventArgs e)
        {
            if (glueCompany.Text.Trim() != "")
            {
                bool chkDup = chkDuplicate();
                if (chkDup == true)
                {
                    LoadBranch();
                    txeBranchNo.Focus();
                }
            }
        }

        private void txeBranchNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeBranchName.Focus();
            }
        }

        private void txeBranchNo_Leave(object sender, EventArgs e)
        {
            if (txeBranchNo.Text.Trim() != "")
            {
                txeBranchNo.Text = txeBranchNo.Text.Trim();
                bool chkDup = chkDuplicate();
                if (chkDup == true)
                {
                    txeBranchName.Focus();
                }
            }
        }

        private void glueCompany_Leave(object sender, EventArgs e)
        {
           
        }

        private void txeBranchName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                glueBranchType.Focus();
            }
        }

        private void txeBranchName_Leave(object sender, EventArgs e)
        {
            if (txeBranchName.Text.Trim() != "")
            {
                txeBranchName.Text = txeBranchName.Text.Trim();
                bool chkDup = chkDuplicate();
                if (chkDup == true)
                {
                    glueBranchType.Focus();
                }
            }
        }

        private void gvBranch_DoubleClick(object sender, EventArgs e)
        {
            DXMouseEventArgs ea = e as DXMouseEventArgs;
            GridView view = sender as GridView;
            GridHitInfo info = view.CalcHitInfo(ea.Location);
            if (info.InRow || info.InRowCell)
            {
                GridView gv = gvBranch;
                lblStatus.Text = "* Edit Branch";
                lblStatus.ForeColor = Color.Red;

                txeID.Text = gv.GetFocusedRowCellValue("OID Branch").ToString();
                glueCompany.EditValue = gv.GetFocusedRowCellValue("OIDCOMPANY").ToString();
                txeBranchNo.Text = gv.GetFocusedRowCellValue("Branch No.").ToString();
                txeBranchName.Text = gv.GetFocusedRowCellValue("Branch Name").ToString();
                glueBranchType.EditValue = gv.GetFocusedRowCellValue("Branch Type").ToString();

                string CreatedBy = gv.GetFocusedRowCellValue("Created By").ToString() == null ? "" : gv.GetFocusedRowCellValue("Created By").ToString();
                glueCREATE.EditValue = CreatedBy;
                txeCDATE.Text = gv.GetFocusedRowCellValue("Created Date").ToString();
            }

        }

        private void gvBranch_RowCellClick(object sender, RowCellClickEventArgs e)
        {
            if (gvBranch.IsFilterRow(e.RowHandle)) return;
        }

        private void gvBranch_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator) e.Info.DisplayText = (e.RowHandle + 1).ToString();
            gvBranch.IndicatorWidth = 40;
        }
    }
}