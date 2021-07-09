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
    public partial class M03 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private Functionality.Function FUNC = new Functionality.Function();
        public LogIn UserLogin { get; set; }

        public string ConnectionString { get; set; }
        string CONNECT_STRING = "";
        DatabaseConnect DBC;


        public M03()
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
            sbSQL.Append("SELECT TOP (1) ReadWriteStatus FROM FunctionAccess WHERE (OIDUser = '" + UserLogin.OIDUser + "') AND(FunctionNo = 'M03') ");
            int chkReadWrite = this.DBC.DBQuery(sbSQL).getInt();
            if (chkReadWrite == 0)
                ribbonPageGroup1.Visible = false;

            sbSQL.Clear();
            sbSQL.Append("SELECT FullName, OIDUSER FROM Users ORDER BY OIDUSER ");
            new ObjDE.setGridLookUpEdit(glueCREATE, sbSQL, "FullName", "OIDUSER").getData();

            glueCREATE.EditValue = UserLogin.OIDUser;

            bbiNew.PerformClick();
            LoadData();
            cbeColorType.EditValue = "";
        }

        private void LoadData()
        {
            LoadType();

            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT '0' AS ID, 'Finished Goods' AS ColorType ");
            sbSQL.Append("UNION ALL ");
            sbSQL.Append("SELECT '1' AS ID, 'Fabric' AS ColorType ");
            sbSQL.Append("UNION ALL ");
            sbSQL.Append("SELECT '2' AS ID, 'Accessory' AS ColorType ");
            sbSQL.Append("UNION ALL ");
            sbSQL.Append("SELECT '3' AS ID, 'Packaging' AS ColorType ");
            new ObjDE.setGridLookUpEdit(cbeColorType, sbSQL, "ColorType", "ID").getData();
        }

        private void LoadType(string ColorType = "")
        {
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT OIDCOLOR AS No, ColorNo, ColorName, ColorType, CASE WHEN ColorType=0 THEN 'Finished Goods' ELSE CASE WHEN ColorType=1 THEN 'Fabric' ELSE CASE WHEN ColorType=2 THEN 'Accessory' ELSE CASE WHEN ColorType=3 THEN 'Packaging' ELSE '' END END END END AS ColorTypeName, CreatedBy, CreatedDate ");
            sbSQL.Append("FROM ProductColor ");
            if (ColorType != "")
            {
                sbSQL.Append("WHERE (ColorType = '" + ColorType + "') ");
            }
            sbSQL.Append("ORDER BY ColorType, ColorName, OIDCOLOR ");
            new ObjDE.setGridControl(gcColor, gvColor, sbSQL).getData(false, false, false, true);
            gvColor.Columns[0].Visible = false;
            gvColor.Columns[3].Visible = false;
            gvColor.Columns[5].Visible = false;
            gvColor.Columns[6].Visible = false;
        }

        private void NewData()
        {
            txeColorID.EditValue = this.DBC.DBQuery("SELECT CASE WHEN ISNULL(MAX(OIDCOLOR), '') = '' THEN 1 ELSE MAX(OIDCOLOR) + 1 END AS NewNo FROM ProductColor").getString();
            txeColorNo.EditValue = "";
            txeColorName.EditValue = "";

            string ColorType = "";
            if (cbeColorType.Text.Trim() != "")
            {
                ColorType = cbeColorType.EditValue.ToString();
            }
            LoadType(ColorType);

            lblStatus.Text = "* Add Color";
            lblStatus.ForeColor = Color.Green;

            glueCREATE.EditValue = UserLogin.OIDUser;
            txeDATE.EditValue = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
        }

        private void gvColor_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {
            
        }

        private void bbiNew_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            //LoadData();
            NewData();
        }

        private void bbiSave_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (txeColorNo.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input color no.");
                txeColorNo.Focus();
            }
            else if (txeColorName.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input color name.");
                txeColorName.Focus();
            }
            else if (cbeColorType.Text.Trim() == "")
            {
                FUNC.msgWarning("Please select color type.");
                cbeColorType.Focus();
            }
            else
            {
                bool chkDup = chkDuplicateNo();
                if (chkDup == true)
                {
                    if (FUNC.msgQuiz("Confirm save data ?") == true)
                    {
                        StringBuilder sbSQL = new StringBuilder();
                        //CalendarMaster

                        int ComType = Convert.ToInt32(cbeColorType.EditValue.ToString());

                        string strCREATE = UserLogin.OIDUser.ToString() != "" ? UserLogin.OIDUser.ToString() : "0";

                        if (lblStatus.Text == "* Add Color")
                        {
                            sbSQL.Append("  INSERT INTO ProductColor(ColorNo, ColorName, ColorType, CreatedBy, CreatedDate) ");
                            sbSQL.Append("  VALUES(N'" + txeColorNo.Text.Trim().Replace("'", "''") + "', N'" + txeColorName.Text.Trim().Replace("'", "''") + "', '" + ComType.ToString() + "', '" + strCREATE + "', GETDATE()) ");
                        }
                        else if (lblStatus.Text == "* Edit Color")
                        {
                            sbSQL.Append("  UPDATE ProductColor SET ");
                            sbSQL.Append("      ColorNo = N'" + txeColorNo.Text.Trim().Replace("'", "''") + "', ColorName = N'" + txeColorName.Text.Trim().Replace("'", "''") + "', ColorType = '" + ComType.ToString() + "' ");
                            sbSQL.Append("  WHERE (OIDCOLOR = '" + txeColorID.Text.Trim() + "') ");
                        }

                        //sbSQL.Append("IF NOT EXISTS(SELECT OIDCOLOR FROM ProductColor WHERE OIDCOLOR = '" + txeColorID.Text.Trim() + "') ");
                        //sbSQL.Append(" BEGIN ");
                        //sbSQL.Append("  INSERT INTO ProductColor(ColorNo, ColorName, ColorType, CreatedBy, CreatedDate) ");
                        //sbSQL.Append("  VALUES(N'" + txeColorNo.Text.Trim().Replace("'", "''") + "', N'" + txeColorName.Text.Trim().Replace("'", "''") + "', '" + ComType.ToString() + "', '" + strCREATE + "', GETDATE()) ");
                        //sbSQL.Append(" END ");
                        //sbSQL.Append("ELSE ");
                        //sbSQL.Append(" BEGIN ");
                        //sbSQL.Append("  UPDATE ProductColor SET ");
                        //sbSQL.Append("      ColorNo = N'" + txeColorNo.Text.Trim().Replace("'", "''") + "', ColorName = N'" + txeColorName.Text.Trim().Replace("'", "''") + "', ColorType = '" + ComType.ToString() + "' ");
                        //sbSQL.Append("  WHERE (OIDCOLOR = '" + txeColorID.Text.Trim() + "') ");
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
                    txeColorNo.Text = "";
                    txeColorNo.Focus();
                    FUNC.msgWarning("Duplicate color no. !! Please Change.");
                }
            }
        }

        private void txeColorNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeColorName.Focus();
            }
        }

        private void txeColorNo_LostFocus(object sender, EventArgs e)
        {
            

        }

        private void txeColorName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeColorNo.Focus();
            }
        }

        private void txeColorName_LostFocus(object sender, EventArgs e)
        {
           
        }

        private bool chkDuplicateNo()
        {
            bool chkDup = true;
            if (txeColorNo.Text != "")
            {
                txeColorNo.Text = txeColorNo.Text.Trim();
                string ColorType = cbeColorType.Text.Trim() != "" ? cbeColorType.EditValue.ToString() : "0";

                if (txeColorNo.Text.Trim() != "" && lblStatus.Text == "* Add Color")
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT TOP(1) ColorNo FROM ProductColor WHERE (ColorType = '" + ColorType + "') AND (ColorNo = N'" + txeColorNo.Text.Trim().Trim().Replace("'", "''") + "') ");
                    if (this.DBC.DBQuery(sbSQL).getString() != "")
                    {
                        chkDup = false;
                    }
                }
                else if (txeColorNo.Text.Trim() != "" && lblStatus.Text == "* Edit Color")
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT TOP(1) OIDCOLOR ");
                    sbSQL.Append("FROM ProductColor ");
                    sbSQL.Append("WHERE (ColorType = '" + ColorType + "') AND (ColorNo = N'" + txeColorNo.Text.Trim().Trim().Replace("'", "''") + "') ");
                    string strCHK = this.DBC.DBQuery(sbSQL).getString();
                    if (strCHK != "" && strCHK != txeColorID.Text.Trim())
                    {
                        //FUNC.msgWarning("Duplicate color no. !! Please Change.");
                        //txeColorNo.Text = "";
                        //txeColorNo.Focus();
                        chkDup = false;
                    }
                }
            }
            return chkDup;
        }

        private bool chkDuplicateName()
        {
            bool chkDup = true;
            if (txeColorName.Text != "")
            {
                txeColorName.Text = txeColorName.Text.Trim();
                string ColorType = cbeColorType.Text.Trim() != "" ? cbeColorType.EditValue.ToString() : "0";
                if (txeColorName.Text.Trim() != "" && lblStatus.Text == "* Add Color")
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT TOP(1) ColorName FROM ProductColor WHERE (ColorType = '" + ColorType + "') AND (ColorName = N'" + txeColorName.Text.Trim().Replace("'", "''") + "') ");
                    if (this.DBC.DBQuery(sbSQL).getString() != "")
                    {
                        FUNC.msgWarning("Duplicate color name. !! Please Change.");
                        txeColorName.Text = "";
                        txeColorName.Focus();
                        chkDup = false;
                    }
                }
                else if (txeColorName.Text.Trim() != "" && lblStatus.Text == "* Edit Color")
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT TOP(1) OIDCOLOR ");
                    sbSQL.Append("FROM ProductColor ");
                    sbSQL.Append("WHERE (ColorType = '" + ColorType + "') AND (ColorName = N'" + txeColorName.Text.Trim().Replace("'", "''") + "') ");
                    string strCHK = this.DBC.DBQuery(sbSQL).getString();
                    if (strCHK != "" && strCHK != txeColorID.Text.Trim())
                    {
                        FUNC.msgWarning("Duplicate color name. !! Please Change.");
                        txeColorName.Text = "";
                        txeColorName.Focus();
                        chkDup = false;
                    }
                }
            }
            return chkDup;
        }


        private void gvColor_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            
        }

        private void gvColor_RowClick(object sender, RowClickEventArgs e)
        {
            if (gvColor.IsFilterRow(e.RowHandle)) return;
            //lblStatus.Text = "* Edit Color";
            //lblStatus.ForeColor = Color.Red;
            //txeColorID.EditValue = gvColor.GetFocusedRowCellValue("No").ToString();
            //txeColorNo.EditValue = gvColor.GetFocusedRowCellValue("ColorNo").ToString();
            //txeColorName.EditValue = gvColor.GetFocusedRowCellValue("ColorName").ToString();
            //cbeColorType.EditValue = gvColor.GetFocusedRowCellValue("ColorType").ToString();

            //glueCREATE.EditValue = gvColor.GetFocusedRowCellValue("CreatedBy").ToString();
            //txeDATE.EditValue = gvColor.GetFocusedRowCellValue("CreatedDate").ToString();
        }

        private void gvColor_DoubleClick(object sender, EventArgs e)
        {
            DXMouseEventArgs ea = e as DXMouseEventArgs;
            GridView view = sender as GridView;
            GridHitInfo info = view.CalcHitInfo(ea.Location);
            if (info.InRow || info.InRowCell)
            {
                GridView gv = gvColor;
                lblStatus.Text = "* Edit Color";
                lblStatus.ForeColor = Color.Red;

                txeColorID.EditValue = gv.GetFocusedRowCellValue("No").ToString();
                txeColorNo.EditValue = gv.GetFocusedRowCellValue("ColorNo").ToString();
                txeColorName.EditValue = gv.GetFocusedRowCellValue("ColorName").ToString();
                cbeColorType.EditValue = gv.GetFocusedRowCellValue("ColorType").ToString();

                glueCREATE.EditValue = gv.GetFocusedRowCellValue("CreatedBy").ToString();
                txeDATE.EditValue = gv.GetFocusedRowCellValue("CreatedDate").ToString();
            }
        }

        private void bbiPrintPreview_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            gcColor.ShowPrintPreview();
        }

        private void bbiPrint_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            gcColor.Print();
        }

        private void bbiRefresh_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            LoadData();
            NewData();
            cbeColorType.EditValue = "";
        }

        private void cbeColorType_EditValueChanged(object sender, EventArgs e)
        {
            string ColorType = "";
            if (cbeColorType.Text.Trim() != "")
            {
                ColorType = cbeColorType.EditValue.ToString();
            }
            LoadType(ColorType);
        }

        private void txeColorNo_Leave(object sender, EventArgs e)
        {
            if (txeColorNo.Text.Trim() != "")
            {
                txeColorNo.Text = txeColorNo.Text.ToUpper().Trim();
                bool chkDup = chkDuplicateNo();
                if (chkDup == true)
                {
                    txeColorName.Focus();
                }
                else
                {
                    txeColorNo.Text = "";
                    txeColorNo.Focus();
                    FUNC.msgWarning("Duplicate color no. !! Please Change.");
                    
                }
            }
        }

        private void txeColorName_Leave(object sender, EventArgs e)
        {
            if (txeColorName.Text.Trim() != "")
            {
                txeColorName.Text = txeColorName.Text.ToUpper().Trim();
            }


        }

        private void gvColor_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator) e.Info.DisplayText = (e.RowHandle + 1).ToString();
            gvColor.IndicatorWidth = 40;
        }

       
    }
}