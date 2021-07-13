using System;
using System.Text;
using System.Windows.Forms;
using DevExpress.LookAndFeel;
using DevExpress.Utils.Extensions;
using DBConnect;
using System.Drawing;
using DevExpress.XtraGrid.Views.Grid;
using TheepClass;
using DevExpress.Utils;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;

namespace MDS.Master
{
    public partial class M10 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private Functionality.Function FUNC = new Functionality.Function();
        public LogIn UserLogin { get; set; }
        public int Company { get; set; }
        public string ConnectionString { get; set; }
        string CONNECT_STRING = "";
        DatabaseConnect DBC;

        public M10()
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
            sbSQL.Append("SELECT TOP (1) ReadWriteStatus FROM FunctionAccess WHERE (OIDUser = '" + UserLogin.OIDUser + "') AND(FunctionNo = 'M10') ");
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
            sbSQL.Append("SELECT OIDSIZE AS No, SizeNo, SizeName, CreatedBy, CreatedDate ");
            sbSQL.Append("FROM ProductSize ");
            sbSQL.Append("ORDER BY OIDSIZE ");
            new ObjDE.setGridControl(gcSize, gvSize, sbSQL).getData(false, false, false, true);
            gvSize.Columns[0].Visible = false;
            gvSize.Columns[3].Visible = false;
            gvSize.Columns[4].Visible = false;
        }

        private void NewData()
        {
            txeID.EditValue = this.DBC.DBQuery("SELECT CASE WHEN ISNULL(MAX(OIDSIZE), '') = '' THEN 1 ELSE MAX(OIDSIZE) + 1 END AS NewNo FROM ProductSize").getString();
            txeSizeNo.EditValue = "";
            txeSizeName.EditValue = "";

            lblStatus.Text = "* Add Size";
            lblStatus.ForeColor = Color.Green;

            glueCREATE.EditValue = UserLogin.OIDUser;
            txeDATE.EditValue = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
            txeSizeNo.Focus();
        }

        private void bbiNew_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            LoadData();
            NewData();
        }

        private void gvGarment_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {
            
        }

        private void txeSizeNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeSizeName.Focus();
            }
        }

        private void txeSizeNo_LostFocus(object sender, EventArgs e)
        {
            
        }

        private void txeSizeName_LostFocus(object sender, EventArgs e)
        {

        }

        private bool chkDuplicateNo()
        {
            bool chkDup = true;
            if (txeSizeNo.Text != "")
            {
                txeSizeNo.Text = txeSizeNo.Text.Trim();
                if (txeSizeNo.Text.Trim() != "" && lblStatus.Text == "* Add Size")
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT TOP(1) SizeNo FROM ProductSize WHERE (SizeNo = N'" + txeSizeNo.Text.Trim() + "') ");
                    if (this.DBC.DBQuery(sbSQL).getString() != "")
                    {
                        chkDup = false;
                    }
                }
                else if (txeSizeNo.Text.Trim() != "" && lblStatus.Text == "* Edit Size")
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT TOP(1) OIDSIZE ");
                    sbSQL.Append("FROM ProductSize ");
                    sbSQL.Append("WHERE (SizeNo = N'" + txeSizeNo.Text.Trim().Replace("'", "''") + "') ");
                    string strCHK = this.DBC.DBQuery(sbSQL).getString();
                    if (strCHK != "" && strCHK != txeID.Text.Trim())
                    {
                        chkDup = false;
                    }
                }
            }
            return chkDup;
        }

        private bool chkDuplicateName()
        {
            bool chkDup = true;
            if (txeSizeName.Text != "")
            {
                txeSizeName.Text = txeSizeName.Text.Trim();
                if (txeSizeName.Text.Trim() != "" && lblStatus.Text == "* Add Size")
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT TOP(1) SizeName FROM ProductSize WHERE (SizeName = N'" + txeSizeName.Text.Trim() + "') ");
                    if (this.DBC.DBQuery(sbSQL).getString() != "")
                    {
                        txeSizeName.Text = "";
                        txeSizeName.Focus();
                        chkDup = false;
                        FUNC.msgWarning("Duplicate size name. !! Please Change.");
                    }
                }
                else if (txeSizeName.Text.Trim() != "" && lblStatus.Text == "* Edit Size")
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT TOP(1) OIDSIZE ");
                    sbSQL.Append("FROM ProductSize ");
                    sbSQL.Append("WHERE (SizeName = N'" + txeSizeName.Text.Trim().Replace("'", "''") + "') ");
                    string strCHK = this.DBC.DBQuery(sbSQL).getString();

                    if (strCHK != "" && strCHK != txeID.Text.Trim())
                    {
                        txeSizeName.Text = "";
                        txeSizeName.Focus();
                        chkDup = false;
                        FUNC.msgWarning("Duplicate size name. !! Please Change.");
                    }
                }
            }
            return chkDup;
        }

        private void gvSize_RowStyle(object sender, RowStyleEventArgs e)
        {
            
        }

        private void bbiSave_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (txeSizeNo.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input size no.");
                txeSizeNo.Focus();
            }
            else if (txeSizeName.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input size name.");
                txeSizeName.Focus();
            }
            else
            {
                bool chkGMP = chkDuplicateNo();

                if (chkGMP == true)
                {
                    if (FUNC.msgQuiz("Confirm save data ?") == true)
                    {
                        StringBuilder sbSQL = new StringBuilder();

                        string strCREATE = UserLogin.OIDUser.ToString() != "" ? UserLogin.OIDUser.ToString() : "0";

                        if (lblStatus.Text == "* Add Size")
                        {
                            sbSQL.Append("  INSERT INTO ProductSize(SizeNo, SizeName, CreatedBy, CreatedDate) ");
                            sbSQL.Append("  VALUES(N'" + txeSizeNo.Text.Trim().Replace("'", "''") + "', N'" + txeSizeName.Text.Trim().Replace("'", "''") + "', '" + strCREATE + "', GETDATE()) ");
                        }
                        else if (lblStatus.Text == "* Edit Size")
                        {
                            sbSQL.Append("  UPDATE ProductSize SET ");
                            sbSQL.Append("      SizeNo = N'" + txeSizeNo.Text.Trim().Replace("'", "''") + "', SizeName = N'" + txeSizeName.Text.Trim().Replace("'", "''") + "' ");
                            sbSQL.Append("  WHERE (OIDSIZE = '" + txeID.Text.Trim() + "') ");
                        }

                        //sbSQL.Append("IF NOT EXISTS(SELECT OIDSIZE FROM ProductSize WHERE OIDSIZE = '" + txeID.Text.Trim() + "') ");
                        //sbSQL.Append(" BEGIN ");
                        //sbSQL.Append("  INSERT INTO ProductSize(SizeNo, SizeName, CreatedBy, CreatedDate) ");
                        //sbSQL.Append("  VALUES(N'" + txeSizeNo.Text.Trim().Replace("'", "''") + "', N'" + txeSizeName.Text.Trim().Replace("'", "''") + "', '" + strCREATE + "', GETDATE()) ");
                        //sbSQL.Append(" END ");
                        //sbSQL.Append("ELSE ");
                        //sbSQL.Append(" BEGIN ");
                        //sbSQL.Append("  UPDATE ProductSize SET ");
                        //sbSQL.Append("      SizeNo = N'" + txeSizeNo.Text.Trim().Replace("'", "''") + "', SizeName = N'" + txeSizeName.Text.Trim().Replace("'", "''") + "' ");
                        //sbSQL.Append("  WHERE (OIDSIZE = '" + txeID.Text.Trim() + "') ");
                        //sbSQL.Append(" END ");
                        //MessageBox.Show(sbSQL.ToString());
                        if (sbSQL.Length > 0)
                        {
                            try
                            {
                                bool chkSAVE = this.DBC.DBQuery(sbSQL).runSQL();
                                if (chkSAVE == true)
                                {

                                    bbiNew.PerformClick();
                                    FUNC.msgInfo("Save complete.");
                                }
                            }
                            catch (Exception)
                            { }
                        }
                    }
                }
                else
                {
                    txeSizeNo.Text = "";
                    txeSizeNo.Focus();
                    FUNC.msgWarning("Duplicate size no. !! Please Change.");
                }
            }
        }

        private void bbiExcel_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            string pathFile = new ObjSet.Folder(@"C:\MDS\Export\").GetPath() + "SizeList_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
            gvSize.ExportToXlsx(pathFile);
            System.Diagnostics.Process.Start(pathFile);
        }

        private void txeSizeName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeSizeNo.Focus();
            }
        }

        private void gvSize_RowClick(object sender, RowClickEventArgs e)
        {
            if (gvSize.IsFilterRow(e.RowHandle)) return;
            
        }

        private void gvSize_DoubleClick(object sender, EventArgs e)
        {
            DXMouseEventArgs ea = e as DXMouseEventArgs;
            GridView view = sender as GridView;
            GridHitInfo info = view.CalcHitInfo(ea.Location);
            if (info.InRow || info.InRowCell)
            {
                GridView gv = gvSize;
                lblStatus.Text = "* Edit Size";
                lblStatus.ForeColor = Color.Red;
                txeID.Text = gv.GetFocusedRowCellValue("No").ToString();
                txeSizeNo.Text = gv.GetFocusedRowCellValue("SizeNo").ToString();
                txeSizeName.Text = gv.GetFocusedRowCellValue("SizeName").ToString();

                string CreatedBy = gv.GetFocusedRowCellValue("CreatedBy").ToString() == null ? "" : gv.GetFocusedRowCellValue("CreatedBy").ToString();
                glueCREATE.EditValue = CreatedBy;
                txeDATE.Text = gv.GetFocusedRowCellValue("CreatedDate").ToString();
            }
        }

        private void bbiPrintPreview_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            gcSize.ShowPrintPreview();
        }

        private void bbiPrint_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            gcSize.Print();
        }

        private void txeSizeNo_Leave(object sender, EventArgs e)
        {
            txeSizeNo.Text = txeSizeNo.Text.ToUpper().Trim();
            bool chkDup = chkDuplicateNo();
            if (chkDup == true)
            {
                txeSizeName.Focus();
            }
            else
            {
                txeSizeNo.Text = "";
                txeSizeNo.Focus();
                FUNC.msgWarning("Duplicate size no. !! Please Change.");
                
            }
        }

        private void txeSizeName_Leave(object sender, EventArgs e)
        {
            txeSizeName.Text = txeSizeName.Text.ToUpper().Trim();
        }

        private void gvSize_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator) e.Info.DisplayText = (e.RowHandle + 1).ToString();
            gvSize.IndicatorWidth = 40;
        }

       
    }
}