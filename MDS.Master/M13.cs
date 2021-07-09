using System;
using System.Text;
using System.Windows.Forms;
using DevExpress.LookAndFeel;
using DevExpress.Utils.Extensions;
using DBConnect;
using System.Drawing;
using DevExpress.XtraPrinting;
using DevExpress.XtraGrid.Views.Grid;
using TheepClass;
using DevExpress.Utils;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;

namespace MDS.Master
{
    public partial class M13 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private Functionality.Function FUNC = new Functionality.Function();
        public LogIn UserLogin { get; set; }
        public string ConnectionString { get; set; }
        string CONNECT_STRING = "";
        DatabaseConnect DBC;

        public M13()
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
            sbSQL.Append("SELECT TOP (1) ReadWriteStatus FROM FunctionAccess WHERE (OIDUser = '" + UserLogin.OIDUser + "') AND(FunctionNo = 'M13') ");
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
            sbSQL.Append("SELECT OIDUNIT AS No, UnitName, CreatedBy, CreatedDate ");
            sbSQL.Append("FROM Unit ");
            sbSQL.Append("ORDER BY OIDUNIT ");
            new ObjDE.setGridControl(gcUnit, gvUnit, sbSQL).getData(false, false, false, true);
            gvUnit.Columns[0].Visible = false;
            gvUnit.Columns[2].Visible = false;
            gvUnit.Columns[3].Visible = false;
        }

        private void NewData()
        {
            txeUnit.Text = "";
            lblStatus.Text = "* Add Unit";
            lblStatus.ForeColor = Color.Green;

            txeID.Text = this.DBC.DBQuery("SELECT CASE WHEN ISNULL(MAX(OIDUNIT), '') = '' THEN 1 ELSE MAX(OIDUNIT) + 1 END AS NewNo FROM Unit").getString();

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

        private bool chkDuplicate()
        {
            bool chkDup = true;
            if (txeUnit.Text != "")
            {
                txeUnit.Text = txeUnit.Text.Trim();
                if (txeUnit.Text.Trim() != "" && lblStatus.Text == "* Add Unit")
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT TOP(1) UnitName FROM Unit WHERE (UnitName = N'" + txeUnit.Text.Trim().Replace("'", "''") + "') ");
                    if (this.DBC.DBQuery(sbSQL).getString() != "")
                    {
                        chkDup = false;
                    }
                }
                else if (txeUnit.Text.Trim() != "" && lblStatus.Text == "* Edit Unit")
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT TOP(1) OIDUNIT ");
                    sbSQL.Append("FROM Unit ");
                    sbSQL.Append("WHERE (UnitName = N'" + txeUnit.Text.Trim().Replace("'", "''") + "') ");
                    string strCHK = this.DBC.DBQuery(sbSQL).getString();
                    if (strCHK != "" && strCHK != txeID.Text.Trim())
                    {
                        chkDup = false;
                    }
                }
            }
            return chkDup;
        }

        private void txeUnit_LostFocus(object sender, EventArgs e)
        {
            
        }

        private void txeUnit_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeID.Focus();
            }
        }

        private void gvUnit_RowStyle(object sender, RowStyleEventArgs e)
        {
            
        }

        private void bbiExcel_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            string pathFile = new ObjSet.Folder(@"C:\MDS\Export\").GetPath() + "UnitList_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
            gvUnit.ExportToXlsx(pathFile);
            System.Diagnostics.Process.Start(pathFile);
        }

        private void bbiSave_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (txeUnit.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input unit name.");
                txeUnit.Focus();
            }
            else
            {
                txeUnit.Text = txeUnit.Text.Trim();
                bool chkUNIT = chkDuplicate();

                if (chkUNIT == true)
                {
                    if (FUNC.msgQuiz("Confirm save data ?") == true)
                    {
                        StringBuilder sbSQL = new StringBuilder();
                        string strCREATE = UserLogin.OIDUser.ToString() != "" ? UserLogin.OIDUser.ToString() : "0";

                        if (lblStatus.Text == "* Add Unit")
                        {
                            sbSQL.Append("  INSERT INTO Unit(UnitName, CreatedBy, CreatedDate) ");
                            sbSQL.Append("  VALUES(N'" + txeUnit.Text.Trim().Replace("'", "''") + "', '" + strCREATE + "', GETDATE()) ");
                        }
                        else if (lblStatus.Text == "* Edit Unit")
                        {
                            sbSQL.Append("  UPDATE Unit SET ");
                            sbSQL.Append("      UnitName = N'" + txeUnit.Text.Trim().Replace("'", "''") + "' ");
                            sbSQL.Append("  WHERE(OIDUNIT = '" + txeID.Text.Trim() + "') ");
                        }

                        //sbSQL.Append("IF NOT EXISTS(SELECT OIDUNIT FROM Unit WHERE OIDUNIT = N'" + txeID.Text.Trim() + "') ");
                        //sbSQL.Append(" BEGIN ");
                        //sbSQL.Append("  INSERT INTO Unit(UnitName, CreatedBy, CreatedDate) ");
                        //sbSQL.Append("  VALUES(N'" + txeUnit.Text.Trim().Replace("'", "''") + "', '" + strCREATE + "', GETDATE()) ");
                        //sbSQL.Append(" END ");
                        //sbSQL.Append("ELSE ");
                        //sbSQL.Append(" BEGIN ");
                        //sbSQL.Append("  UPDATE Unit SET ");
                        //sbSQL.Append("      UnitName = N'" + txeUnit.Text.Trim().Replace("'", "''") + "' ");
                        //sbSQL.Append("  WHERE(OIDUNIT = '" + txeID.Text.Trim() + "') ");
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
                    txeUnit.Text = "";
                    txeUnit.Focus();
                    FUNC.msgWarning("Duplicate unit. !! Please Change.");
                }
            }
        }

        private void gvUnit_RowClick(object sender, RowClickEventArgs e)
        {
            if (gvUnit.IsFilterRow(e.RowHandle)) return;
            
        }

        private void gvUnit_DoubleClick(object sender, EventArgs e)
        {
            DXMouseEventArgs ea = e as DXMouseEventArgs;
            GridView view = sender as GridView;
            GridHitInfo info = view.CalcHitInfo(ea.Location);
            if (info.InRow || info.InRowCell)
            {
                GridView gv = gvUnit;
                lblStatus.Text = "* Edit Unit";
                lblStatus.ForeColor = Color.Red;

                txeID.Text = gv.GetFocusedRowCellValue("No").ToString();
                txeUnit.Text = gv.GetFocusedRowCellValue("UnitName").ToString();

                string CreatedBy = gv.GetFocusedRowCellValue("CreatedBy").ToString() == null ? "" : gv.GetFocusedRowCellValue("CreatedBy").ToString();
                glueCREATE.EditValue = CreatedBy;

                txeDATE.Text = gv.GetFocusedRowCellValue("CreatedDate").ToString();
            }
        }

        private void bbiPrintPreview_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            gcUnit.ShowPrintPreview();
        }

        private void bbiPrint_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            gcUnit.Print();
        }

        private void txeUnit_Leave(object sender, EventArgs e)
        {
            if (txeUnit.Text.Trim() != "")
            {
                bool chkDup = chkDuplicate();
                if (chkDup == false)
                {
                    txeUnit.Text = "";
                    txeUnit.Focus();
                    FUNC.msgWarning("Duplicate unit. !! Please Change.");
                }
            }
        }

        private void gvUnit_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator) e.Info.DisplayText = (e.RowHandle + 1).ToString();
            gvUnit.IndicatorWidth = 40;
        }

       
    }
}