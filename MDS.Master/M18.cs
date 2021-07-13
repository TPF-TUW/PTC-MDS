using System;
using System.Text;
using DBConnect;
using System.Windows.Forms;
using DevExpress.LookAndFeel;
using DevExpress.Utils.Extensions;
using System.Drawing;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraEditors;
using DevExpress.XtraLayout.Utils;
using TheepClass;
using System.Globalization;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Data;
using System.Text.RegularExpressions;
using System.IO;
using DevExpress.Utils;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;

namespace MDS.Master
{
    public partial class M18 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private Functionality.Function FUNC = new Functionality.Function();
        public LogIn UserLogin { get; set; }
        public int Company { get; set; }
        public string ConnectionString { get; set; }
        string CONNECT_STRING = "";
        DatabaseConnect DBC;

        public M18()
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
            sbSQL.Append("SELECT TOP (1) ReadWriteStatus FROM FunctionAccess WHERE (OIDUser = '" + UserLogin.OIDUser + "') AND(FunctionNo = 'M18') ");
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
            sbSQL.Clear();
            sbSQL.Append("SELECT SeasonNo AS [Season No.], SeasonName AS [Season Name] ");
            sbSQL.Append("FROM Season ");
            sbSQL.Append("ORDER BY OIDSEASON");
            new ObjDE.setGridLookUpEdit(glueSeason, sbSQL, "Season No.", "Season No.").getData();

            sbSQL.Clear();
            sbSQL.Append("SELECT Code, Name AS Customer, OIDCUST AS ID ");
            sbSQL.Append("FROM Customer ");
            sbSQL.Append("ORDER BY Code ");
            new ObjDE.setSearchLookUpEdit(slueCustomer, sbSQL, "Customer", "ID").getData();
            slueCustomer.Properties.View.PopulateColumns(slueCustomer.Properties.DataSource);
            slueCustomer.Properties.View.Columns["ID"].Visible = false;

            sbSQL.Clear();
            sbSQL.Append("SELECT PS.StyleName, GC.CategoryName, PS.OIDSTYLE AS ID ");
            sbSQL.Append("FROM   ProductStyle AS PS INNER JOIN ");
            sbSQL.Append("       GarmentCategory AS GC ON PS.OIDGCATEGORY = GC.OIDGCATEGORY ");
            sbSQL.Append("ORDER BY PS.StyleName ");
            new ObjDE.setSearchLookUpEdit(slueStyle, sbSQL, "StyleName", "ID").getData();
            slueStyle.Properties.View.PopulateColumns(slueStyle.Properties.DataSource);
            slueStyle.Properties.View.Columns["ID"].Visible = false;

            LoadItemCustomer();
        }

        private void LoadItemCustomer()
        {
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT IC.OIDCUST, CUS.Name AS Customer, IC.ItemCode, IC.ItemName, IC.StyleNo, IC.OIDSTYLE, PS.StyleName, IC.Season, IC.FabricWidth, IC.FBComposition, IC.OIDCSITEM AS ID ");
            sbSQL.Append("FROM   ItemCustomer AS IC INNER JOIN ");
            sbSQL.Append("       Customer AS CUS ON IC.OIDCUST = CUS.OIDCUST LEFT OUTER JOIN ");
            sbSQL.Append("       ProductStyle AS PS ON IC.OIDSTYLE = PS.OIDSTYLE ");
            sbSQL.Append("WHERE  (IC.OIDCSITEM > 0) ");
            if(slueCustomer.Text.Trim() != "")
                sbSQL.Append("AND  (IC.OIDCUST = '" + slueCustomer.EditValue.ToString() + "') ");
            if(slueStyle.Text.Trim() != "")
                sbSQL.Append("AND  (IC.OIDSTYLE = '" + slueStyle.EditValue.ToString() + "') ");
            if (glueSeason.Text.Trim() != "")
                sbSQL.Append("AND  (IC.Season = N'" + speSeason.Value.ToString() + glueSeason.EditValue.ToString() + "') ");
            sbSQL.Append("ORDER BY IC.OIDCUST, IC.ItemCode ");
            new ObjDE.setGridControl(gcCustItem, gvCustItem, sbSQL).getData(false, false, false, true);
            gvCustItem.Columns["ID"].Visible = false;
            gvCustItem.Columns["OIDCUST"].Visible = false;
            gvCustItem.Columns["OIDSTYLE"].Visible = false;
        }

        private void NewData()
        {
            lblStatus.Text = "* Add Item";
            lblStatus.ForeColor = Color.Green;

            txeID.Text = this.DBC.DBQuery("SELECT CASE WHEN ISNULL(MAX(OIDCSITEM), '') = '' THEN 1 ELSE MAX(OIDCSITEM) + 1 END AS NewNo FROM ItemCustomer").getString();

            if (Convert.ToInt32(DateTime.Now.ToString("yyyy")) > 2500)
                speSeason.Value = Convert.ToInt32(DateTime.Now.ToString("yyyy")) - 543;
            else
                speSeason.Value = Convert.ToInt32(DateTime.Now.ToString("yyyy"));

            glueSeason.EditValue = "";
            slueCustomer.EditValue = "";
            slueStyle.EditValue = "";
            txeItemCode.Text = "";
            txeItemName.Text = "";
            txeFabricWidth.Text = "";
            txeFBComposition.Text = "";
            txeStyleNo.Text = "";
            txeStyleCode.Text = "";

            glueCREATE.EditValue = UserLogin.OIDUser;
            txeCDATE.Text = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");

            slueCustomer.Focus();
        }

        private void bbiNew_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            LoadData();
            NewData();
        }

        private void bbiSave_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (slueCustomer.Text.Trim() == "")
            {
                FUNC.msgWarning("Please select customer");
                slueCustomer.Focus();
            }
            else if (txeItemCode.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input item code.");
                txeItemCode.Focus();
            }
            else if (txeItemName.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input item name.");
                txeItemName.Focus();
            }
            else
            {
                bool chkDup = chkDuplicate();
                if (chkDup == true)
                {
                    if (FUNC.msgQuiz("Confirm save data ?") == true)
                    {
                        StringBuilder sbSQL = new StringBuilder();

                        string OIDSTYLE = slueStyle.Text.Trim() != "" ? "'" + slueStyle.EditValue.ToString() + "'" : "NULL";
                        string Season = glueSeason.Text.Trim() != "" ? speSeason.Value.ToString() + glueSeason.EditValue.ToString() : "";


                        if (lblStatus.Text == "* Add Item")
                        {
                            sbSQL.Append("  INSERT INTO ItemCustomer(OIDCUST, ItemCode, ItemName, OIDSTYLE, Season, FabricWidth, FBComposition, StyleNo) ");
                            sbSQL.Append("  VALUES(N'" + slueCustomer.EditValue.ToString() + "', N'" + txeItemCode.Text.Trim().Replace("'", "''") + "', N'" + txeItemName.Text.Trim().Replace("'", "''") + "', " + OIDSTYLE + ", N'" + Season + "', N'" + txeFabricWidth.Text.Trim().Replace("'", "''") + "', N'" + txeFBComposition.Text.Trim().Replace("'", "''") + "', N'" + txeStyleNo.Text.Trim() + txeStyleCode.Text.Trim() + "') ");
                        }
                        else if (lblStatus.Text == "* Edit Item")
                        {
                            sbSQL.Append("  UPDATE ItemCustomer SET ");
                            sbSQL.Append("      OIDCUST=N'" + slueCustomer.EditValue.ToString() + "', ");
                            sbSQL.Append("      ItemCode=N'" + txeItemCode.Text.Trim().Replace("'", "''") + "', ");
                            sbSQL.Append("      ItemName=N'" + txeItemName.Text.Trim().Replace("'", "''") + "', ");
                            sbSQL.Append("      OIDSTYLE=" + OIDSTYLE + ", ");
                            sbSQL.Append("      Season=N'" + Season + "', ");
                            sbSQL.Append("      FabricWidth=N'" + txeFabricWidth.Text.Trim().Replace("'", "''") + "', ");
                            sbSQL.Append("      FBComposition=N'" + txeFBComposition.Text.Trim().Replace("'", "''") + "', ");
                            sbSQL.Append("      StyleNo=N'" + txeStyleNo.Text.Trim() + txeStyleCode.Text.Trim() + "' ");
                            sbSQL.Append("  WHERE (OIDCSITEM = '" + txeID.Text.Trim() + "') ");
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
                
            }
        }

        private void bbiExcel_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            string pathFile = new ObjSet.Folder(@"C:\MDS\Export\").GetPath() + "CustomerItemList_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
            gvCustItem.ExportToXlsx(pathFile);
            System.Diagnostics.Process.Start(pathFile);
        }

        private void bbiPrintPreview_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            gcCustItem.ShowPrintPreview();
        }

        private void bbiPrint_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            gcCustItem.Print();
        }

        private void slueStyle_EditValueChanged(object sender, EventArgs e)
        {
            txeStyleCode.Text = slueStyle.Text;
            LoadItemCustomer();
            speSeason.Focus();
        }

        private void bbiRefresh_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            LoadItemCustomer();
        }

        private void gvCustItem_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator) e.Info.DisplayText = (e.RowHandle + 1).ToString();
            gvCustItem.IndicatorWidth = 40;
        }

        private void gvCustItem_DoubleClick(object sender, EventArgs e)
        {
            DXMouseEventArgs ea = e as DXMouseEventArgs;
            GridView view = sender as GridView;
            GridHitInfo info = view.CalcHitInfo(ea.Location);
            if (info.InRow || info.InRowCell)
            {
                GridView gv = gvCustItem;
                lblStatus.Text = "* Edit Item";
                lblStatus.ForeColor = Color.Red;

                string Season = gv.GetFocusedRowCellValue("Season").ToString();
                speSeason.Value = Convert.ToInt32(Regex.Match(Season, @"\d+([,\.]\d+)?").Value);
                glueSeason.EditValue = Season.Replace(Regex.Match(Season, @"\d+([,\.]\d+)?").Value, "");
                slueCustomer.EditValue = gv.GetFocusedRowCellValue("OIDCUST").ToString();
                slueStyle.EditValue = gv.GetFocusedRowCellValue("OIDSTYLE").ToString();
                txeItemCode.Text = gv.GetFocusedRowCellValue("ItemCode").ToString();
                txeItemName.Text = gv.GetFocusedRowCellValue("ItemName").ToString();
                txeFabricWidth.Text = gv.GetFocusedRowCellValue("FabricWidth").ToString();
                txeFBComposition.Text = gv.GetFocusedRowCellValue("FBComposition").ToString();

                string StyleNo = gv.GetFocusedRowCellValue("StyleNo").ToString();
                txeStyleNo.Text = Regex.Match(StyleNo, @"\d+([,\.]\d+)?").Value.ToString();
                txeStyleCode.Text = StyleNo.Replace(Regex.Match(StyleNo, @"\d+([,\.]\d+)?").Value, "");
            }

        }

        private void txeItemCode_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txeItemName.Focus();
        }

        private void txeItemCode_Leave(object sender, EventArgs e)
        {
            if (txeItemCode.Text.Trim() != "")
            {
                txeItemCode.Text = txeItemCode.Text.ToUpper().Trim();
                bool chkDup = chkDuplicate();
                if (chkDup == true)
                    txeItemName.Focus();
            }
        }

        private bool chkDuplicate()
        {
            bool chkDup = true;
            if (txeItemCode.Text.Trim() != "")
            {
                txeItemCode.Text = txeItemCode.Text.ToUpper().Trim().Replace("'", "''");
                string CUST = slueCustomer.Text.Trim() != "" ? slueCustomer.EditValue.ToString() : "";
                if (lblStatus.Text == "* Add Item")
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT TOP(1) ItemCode FROM ItemCustomer WHERE (OIDCUST = '" + CUST + "') AND (ItemCode = N'" + txeItemCode.Text + "') ");
                    if (this.DBC.DBQuery(sbSQL).getString() != "")
                    {
                        txeItemCode.Text = "";
                        txeItemCode.Focus();
                        FUNC.msgWarning("Duplicate item code. !! Please Change.");
                        chkDup = false;
                    }
                }
                else if (lblStatus.Text == "* Edit Item")
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT TOP(1) OIDCSITEM ");
                    sbSQL.Append("FROM ItemCustomer ");
                    sbSQL.Append("WHERE (OIDCUST = '" + CUST + "') AND (ItemCode = N'" + txeItemCode.Text + "') ");
                    string strCHK = this.DBC.DBQuery(sbSQL).getString();
                    if (strCHK != "" && strCHK != txeID.Text.Trim())
                    {
                        txeItemCode.Text = "";
                        txeItemCode.Focus();
                        FUNC.msgWarning("Duplicate item code. !! Please Change.");
                        chkDup = false;
                    }
                }
            }
            return chkDup;
        }

        private void slueCustomer_EditValueChanged(object sender, EventArgs e)
        {
            bool chkDup = chkDuplicate();
            if (chkDup == true)
            {
                LoadItemCustomer();
                slueStyle.Focus();
            }
        }

        private void glueSeason_EditValueChanged(object sender, EventArgs e)
        {
            LoadItemCustomer();
            txeItemCode.Focus();
        }

        private void txeItemName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txeFabricWidth.Focus();
        }

        private void txeFabricWidth_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txeFBComposition.Focus();
        }

        private void txeFBComposition_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txeStyleNo.Focus();
        }
    }
}