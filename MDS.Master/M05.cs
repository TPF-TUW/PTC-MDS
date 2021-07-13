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
    public partial class M05 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private Functionality.Function FUNC = new Functionality.Function();
        public LogIn UserLogin { get; set; }
        public int Company { get; set; }
        public string ConnectionString { get; set; }
        string CONNECT_STRING = "";
        DatabaseConnect DBC;


        public M05()
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
            sbSQL.Append("SELECT TOP (1) ReadWriteStatus FROM FunctionAccess WHERE (OIDUser = '" + UserLogin.OIDUser + "') AND(FunctionNo = 'M05') ");
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
            sbSQL.Append("SELECT OIDCURR AS No, Currency, CreateBy, CreateDate ");
            sbSQL.Append("FROM Currency ");
            sbSQL.Append("ORDER BY OIDCURR, Currency ");
            new ObjDE.setGridControl(gcCurrency, gvCurrency, sbSQL).getData(false, false, true, true);

            gvCurrency.Columns[0].Visible = false;
            gvCurrency.Columns[2].Visible = false;
            gvCurrency.Columns[3].Visible = false;

        }

        private void NewData()
        {
            txeCurrency.Text = "";
            lblStatus.Text = "* Add Currency";
            lblStatus.ForeColor = Color.Green;

            txeID.Text = this.DBC.DBQuery("SELECT CASE WHEN ISNULL(MAX(OIDCURR), '') = '' THEN 1 ELSE MAX(OIDCURR) + 1 END AS NewNo FROM Currency").getString();

            glueCREATE.EditValue = UserLogin.OIDUser;
            txeDATE.Text = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");

            //txeID.Focus();
        }

        private void bbiNew_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            LoadData();
            NewData();
        }

        private void gvCurrency_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {
            
            
        }

        private void bbiSave_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (txeCurrency.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input currency name.");
                txeCurrency.Focus();
            }
            else
            {
                txeCurrency.Text = txeCurrency.Text.ToUpper().Trim();
                bool chkCURR = chkDuplicate();

                if (chkCURR == true)
                {
                    if (FUNC.msgQuiz("Confirm save data ?") == true)
                    {
                        StringBuilder sbSQL = new StringBuilder();
                        string strCREATE = UserLogin.OIDUser.ToString() != "" ? UserLogin.OIDUser.ToString() : "0";

                        if (lblStatus.Text == "* Add Currency")
                        {
                            sbSQL.Append("  INSERT INTO Currency(Currency, CreateBy, CreateDate) ");
                            sbSQL.Append("  VALUES(N'" + txeCurrency.Text.Trim().Replace("'", "''") + "', '" + strCREATE + "', GETDATE()) ");
                        }
                        else if (lblStatus.Text == "* Edit Currency")
                        {
                            sbSQL.Append("  UPDATE Currency SET ");
                            sbSQL.Append("      Currency = N'" + txeCurrency.Text.Trim().Replace("'", "''") + "' ");
                            sbSQL.Append("  WHERE(OIDCURR = '" + txeID.Text.Trim() + "') ");
                        }

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
                    txeCurrency.Text = "";
                    txeCurrency.Focus();
                    FUNC.msgWarning("Duplicate currency. !! Please Change.");
                }
            }

        }

        private void txeCurrency_LostFocus(object sender, EventArgs e)
        {
            
        }

        private bool chkDuplicate()
        {
            bool chkDup = true;
            if (txeCurrency.Text != "")
            {
                txeCurrency.Text = txeCurrency.Text.ToUpper().Trim();
                if (txeCurrency.Text.Trim() != "" && lblStatus.Text == "* Add Currency")
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT TOP(1) Currency FROM Currency WHERE (Currency = N'" + txeCurrency.Text.Trim().Replace("'", "''") + "') ");
                    if (this.DBC.DBQuery(sbSQL).getString() != "")
                    {
                        chkDup = false;
                    }
                }
                else if (txeCurrency.Text.Trim() != "" && lblStatus.Text == "* Edit Currency")
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT TOP(1) OIDCURR ");
                    sbSQL.Append("FROM Currency ");
                    sbSQL.Append("WHERE (Currency = N'" + txeCurrency.Text.Trim().Replace("'", "''") + "') ");
                    string strCHK = this.DBC.DBQuery(sbSQL).getString();
                    if (strCHK != "" && strCHK != txeID.Text.Trim())
                    {
                        chkDup = false;
                    }
                }
            }
            return chkDup;
        }


        private void bbiExcel_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            string pathFile = new ObjSet.Folder(@"C:\MDS\Export\").GetPath() + "CurrencyList_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
            gvCurrency.ExportToXlsx(pathFile);
            System.Diagnostics.Process.Start(pathFile);
        }

        private void gvCurrency_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            
        }

        private void txeCurrency_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeID.Focus();
            }
        }

        private void gvCurrency_RowClick(object sender, RowClickEventArgs e)
        {
            if (gvCurrency.IsFilterRow(e.RowHandle)) return;
            //lblStatus.Text = "* Edit Currency";
            //lblStatus.ForeColor = Color.Red;

            //txeID.Text = gvCurrency.GetFocusedRowCellValue("No").ToString();
            //txeCurrency.Text = gvCurrency.GetFocusedRowCellValue("Currency").ToString();

            //glueCREATE.EditValue = gvCurrency.GetFocusedRowCellValue("CreateBy").ToString();
            //txeDATE.Text = gvCurrency.GetFocusedRowCellValue("CreateDate").ToString();
        }

        private void gvCurrency_DoubleClick(object sender, EventArgs e)
        {
            DXMouseEventArgs ea = e as DXMouseEventArgs;
            GridView view = sender as GridView;
            GridHitInfo info = view.CalcHitInfo(ea.Location);
            if (info.InRow || info.InRowCell)
            {
                GridView gv = gvCurrency;
                lblStatus.Text = "* Edit Currency";
                lblStatus.ForeColor = Color.Red;

                txeID.Text = gv.GetFocusedRowCellValue("No").ToString();
                txeCurrency.Text = gv.GetFocusedRowCellValue("Currency").ToString();

                glueCREATE.EditValue = gv.GetFocusedRowCellValue("CreateBy").ToString();
                txeDATE.Text = gv.GetFocusedRowCellValue("CreateDate").ToString();
            }

        }

        private void bbiPrintPreview_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            gcCurrency.ShowPrintPreview();
        }

        private void bbiPrint_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            gcCurrency.Print();
        }

        private void txeCurrency_Leave(object sender, EventArgs e)
        {
            if (txeCurrency.Text.Trim() != "")
            {
                bool chkDup = chkDuplicate();
                if (chkDup == false)
                {
                    txeCurrency.Text = "";
                    txeCurrency.Focus();
                    FUNC.msgWarning("Duplicate currency. !! Please Change.");

                }
            }
        }

        private void gvCurrency_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator) e.Info.DisplayText = (e.RowHandle + 1).ToString();
            gvCurrency.IndicatorWidth = 40;
        }

        private void ribbonControl_Click(object sender, EventArgs e)
        {

        }
    }
}