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
    public partial class M11 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private Functionality.Function FUNC = new Functionality.Function();
        public LogIn UserLogin { get; set; }

        public string ConnectionString { get; set; }
        string CONNECT_STRING = "";
        DatabaseConnect DBC;

        public M11()
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
            sbSQL.Append("SELECT TOP (1) ReadWriteStatus FROM FunctionAccess WHERE (OIDUser = '" + UserLogin.OIDUser + "') AND(FunctionNo = 'M11') ");
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
            sbSQL.Append("SELECT PS.OIDSTYLE AS No, PS.StyleName, PS.OIDGCATEGORY AS CategoryID, GC.CategoryName, PS.CreatedBy, PS.CreatedDate ");
            sbSQL.Append("FROM   ProductStyle AS PS INNER JOIN ");
            sbSQL.Append("       GarmentCategory AS GC ON PS.OIDGCATEGORY = GC.OIDGCATEGORY ");
            sbSQL.Append("ORDER BY GC.CategoryName, PS.StyleName ");
            new ObjDE.setGridControl(gcStyle, gvStyle, sbSQL).getData(false, false, false, true);
            gvStyle.Columns[0].Visible = false;
            gvStyle.Columns[2].Visible = false;
            gvStyle.Columns[4].Visible = false;
            gvStyle.Columns[5].Visible = false;

            sbSQL.Clear();
            sbSQL.Append("SELECT CategoryName, OIDGCATEGORY AS ID ");
            sbSQL.Append("FROM GarmentCategory ");
            sbSQL.Append("ORDER BY CategoryName ");
            new ObjDE.setGridLookUpEdit(glueCategory, sbSQL, "CategoryName", "ID").getData(true);
            glueCategory.Properties.View.PopulateColumns(glueCategory.Properties.DataSource);
            glueCategory.Properties.View.Columns["ID"].Visible = false;
        }

        private void NewData()
        {
            txeID.EditValue = this.DBC.DBQuery("SELECT CASE WHEN ISNULL(MAX(OIDSTYLE), '') = '' THEN 1 ELSE MAX(OIDSTYLE) + 1 END AS NewNo FROM ProductStyle").getString();
            txeStyleNo.Text = "";
            glueCategory.EditValue = "";

            lblStatus.Text = "* Add Style";
            lblStatus.ForeColor = Color.Green;

            glueCREATE.EditValue = UserLogin.OIDUser;
            txeDATE.EditValue = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
            txeStyleNo.Focus();
        }

        private void bbiNew_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            LoadData();
            NewData();
        }

        private void gvGarment_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {
            
        }

        private bool chkDuplicateName()
        {
            bool chkDup = true;
            if (txeStyleNo.Text != "")
            {
                txeStyleNo.Text = txeStyleNo.Text.Trim();
                if (txeStyleNo.Text.Trim() != "" && lblStatus.Text == "* Add Style")
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT TOP(1) StyleName FROM ProductStyle WHERE (StyleName = N'" + txeStyleNo.Text.Trim() + "') ");
                    if (this.DBC.DBQuery(sbSQL).getString() != "")
                    {
                        chkDup = false;
                    }
                }
                else if (txeStyleNo.Text.Trim() != "" && lblStatus.Text == "* Edit Style")
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT TOP(1) OIDSTYLE ");
                    sbSQL.Append("FROM ProductStyle ");
                    sbSQL.Append("WHERE (StyleName = N'" + txeStyleNo.Text.Trim().Replace("'", "''") + "') ");
                    string strCHK = this.DBC.DBQuery(sbSQL).getString();

                    if (strCHK != "" && strCHK != txeID.Text.Trim())
                    {
                        chkDup = false;
                    }
                }
            }
            return chkDup;
        }

        private void txeStyleNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                glueCategory.Focus();
            }
        }

        private void txeStyleNo_LostFocus(object sender, EventArgs e)
        {
            
        }

        private void gvStyle_RowStyle(object sender, RowStyleEventArgs e)
        {
            
        }

        private void bbiSave_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (txeStyleNo.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input style no.");
                txeStyleNo.Focus();
            }
            else if (glueCategory.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input product category.");
                glueCategory.Focus();
            }
            else
            {
                bool chkDup = chkDuplicateName();
                if (chkDup == true)
                {
                    if (FUNC.msgQuiz("Confirm save data ?") == true)
                    {
                        StringBuilder sbSQL = new StringBuilder();

                        string strCREATE = UserLogin.OIDUser.ToString() != "" ? UserLogin.OIDUser.ToString() : "0";

                        if (lblStatus.Text == "* Add Style")
                        {
                            sbSQL.Append("  INSERT INTO ProductStyle(StyleName, OIDGCATEGORY, CreatedBy, CreatedDate) ");
                            sbSQL.Append("  VALUES(N'" + txeStyleNo.Text.Trim().Replace("'", "''") + "', '" + glueCategory.EditValue.ToString() + "', '" + strCREATE + "', GETDATE()) ");
                        }
                        else if (lblStatus.Text == "* Edit Style")
                        {
                            sbSQL.Append("  UPDATE ProductStyle SET ");
                            sbSQL.Append("      StyleName = N'" + txeStyleNo.Text.Trim().Replace("'", "''") + "', OIDGCATEGORY = '" + glueCategory.EditValue.ToString() + "' ");
                            sbSQL.Append("  WHERE (OIDSTYLE = '" + txeID.Text.Trim() + "') ");
                        }

                        //sbSQL.Append("IF NOT EXISTS(SELECT OIDSTYLE FROM ProductStyle WHERE OIDSTYLE = '" + txeID.Text.Trim() + "') ");
                        //sbSQL.Append(" BEGIN ");
                        //sbSQL.Append("  INSERT INTO ProductStyle(StyleName, OIDGCATEGORY, CreatedBy, CreatedDate) ");
                        //sbSQL.Append("  VALUES(N'" + txeStyleNo.Text.Trim().Replace("'", "''") + "', '" + glueCategory.EditValue.ToString() + "', '" + strCREATE + "', GETDATE()) ");
                        //sbSQL.Append(" END ");
                        //sbSQL.Append("ELSE ");
                        //sbSQL.Append(" BEGIN ");
                        //sbSQL.Append("  UPDATE ProductStyle SET ");
                        //sbSQL.Append("      StyleName = N'" + txeStyleNo.Text.Trim().Replace("'", "''") + "', OIDGCATEGORY = '" + glueCategory.EditValue.ToString() + "' ");
                        //sbSQL.Append("  WHERE (OIDSTYLE = '" + txeID.Text.Trim() + "') ");
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
                    txeStyleNo.Text = "";
                    txeStyleNo.Focus();
                    FUNC.msgWarning("Duplicate style name. !! Please Change.");
                }
            }
        }

        private void bbiExcel_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            string pathFile = new ObjSet.Folder(@"C:\MDS\Export\").GetPath() + "StyleList_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
            gvStyle.ExportToXlsx(pathFile);
            System.Diagnostics.Process.Start(pathFile);
        }

        private void gvStyle_RowClick(object sender, RowClickEventArgs e)
        {
            if (gvStyle.IsFilterRow(e.RowHandle)) return;
            
        }

        private void gvStyle_DoubleClick(object sender, EventArgs e)
        {
            DXMouseEventArgs ea = e as DXMouseEventArgs;
            GridView view = sender as GridView;
            GridHitInfo info = view.CalcHitInfo(ea.Location);
            if (info.InRow || info.InRowCell)
            {
                GridView gv = gvStyle;
                lblStatus.Text = "* Edit Style";
                lblStatus.ForeColor = Color.Red;

                txeID.Text = gv.GetFocusedRowCellValue("No").ToString();
                txeStyleNo.Text = gv.GetFocusedRowCellValue("StyleName").ToString();
                glueCategory.EditValue = gv.GetFocusedRowCellValue("CategoryID").ToString();

                string CreatedBy = gv.GetFocusedRowCellValue("CreatedBy").ToString() == null ? "" : gv.GetFocusedRowCellValue("CreatedBy").ToString();
                glueCREATE.EditValue = CreatedBy;
                txeDATE.Text = gv.GetFocusedRowCellValue("CreatedDate").ToString();
            }
        }

        private void bbiPrintPreview_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            gcStyle.ShowPrintPreview();
        }

        private void bbiPrint_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            gcStyle.Print();
        }

        private void txeStyleNo_Leave(object sender, EventArgs e)
        {
            if (txeStyleNo.Text.Trim() != "")
            {
                txeStyleNo.Text = txeStyleNo.Text.ToUpper().Trim();
                bool chkDup = chkDuplicateName();
                if (chkDup == false)
                {
                    txeStyleNo.Text = "";
                    txeStyleNo.Focus();
                    FUNC.msgWarning("Duplicate style name. !! Please Change.");  
                }
            }
        }

        private void glueCategory_EditValueChanged(object sender, EventArgs e)
        {
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT PS.OIDSTYLE AS No, PS.StyleName, PS.OIDGCATEGORY AS CategoryID, GC.CategoryName, PS.CreatedBy, PS.CreatedDate ");
            sbSQL.Append("FROM   ProductStyle AS PS INNER JOIN ");
            sbSQL.Append("       GarmentCategory AS GC ON PS.OIDGCATEGORY = GC.OIDGCATEGORY ");
            if (glueCategory.Text.Trim() != "")
                sbSQL.Append("WHERE (PS.OIDGCATEGORY = '" + glueCategory.EditValue.ToString() + "') ");
            sbSQL.Append("ORDER BY GC.CategoryName, PS.StyleName ");
            new ObjDE.setGridControl(gcStyle, gvStyle, sbSQL).getData(false, false, false, true);
            gvStyle.Columns[0].Visible = false;
            gvStyle.Columns[2].Visible = false;
            gvStyle.Columns[4].Visible = false;
            gvStyle.Columns[5].Visible = false;

            txeStyleNo.Focus();
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            var frm = new M11_01(this.DBC, UserLogin.OIDUser);
            frm.ShowDialog(this);
        }

        private void gvStyle_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator) e.Info.DisplayText = (e.RowHandle + 1).ToString();
            gvStyle.IndicatorWidth = 40;
        }


    }
}