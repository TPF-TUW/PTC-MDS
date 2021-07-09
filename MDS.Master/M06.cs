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
    public partial class M06 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private Functionality.Function FUNC = new Functionality.Function();
        public LogIn UserLogin { get; set; }
        public string ConnectionString { get; set; }
        string CONNECT_STRING = "";
        DatabaseConnect DBC;

        public M06()
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
            sbSQL.Append("SELECT TOP (1) ReadWriteStatus FROM FunctionAccess WHERE (OIDUser = '" + UserLogin.OIDUser + "') AND(FunctionNo = 'M06') ");
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
            sbSQL.Append("SELECT OIDGParts AS No, GarmentParts, CreatedBy, CreatedDate ");
            sbSQL.Append("FROM GarmentParts ");
            sbSQL.Append("ORDER BY OIDGParts, GarmentParts ");
            new ObjDE.setGridControl(gcGarment, gvGarment, sbSQL).getData(false, false, true, true);
            gvGarment.Columns[0].Visible = false;
            gvGarment.Columns[2].Visible = false;
            gvGarment.Columns[3].Visible = false;
        }

        private void NewData()
        {
            txeGarment.Text = "";
            lblStatus.Text = "* Add Garment";
            lblStatus.ForeColor = Color.Green;

            txeID.Text = this.DBC.DBQuery("SELECT CASE WHEN ISNULL(MAX(OIDGParts), '') = '' THEN 1 ELSE MAX(OIDGParts) + 1 END AS NewNo FROM GarmentParts").getString();

            glueCREATE.EditValue = UserLogin.OIDUser;
            txeDATE.Text = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");

            //txeID.Focus();
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
            if (txeGarment.Text != "")
            {
                txeGarment.Text = txeGarment.Text.Trim();
                if (txeGarment.Text.Trim() != "" && lblStatus.Text == "* Add Garment")
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT TOP(1) GarmentParts FROM GarmentParts WHERE (GarmentParts = N'" + txeGarment.Text.Trim().Replace("'", "''") + "') ");
                    if (this.DBC.DBQuery(sbSQL).getString() != "")
                    {
                        chkDup = false;
                    }
                }
                else if (txeGarment.Text.Trim() != "" && lblStatus.Text == "* Edit Garment")
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT TOP(1) OIDGParts ");
                    sbSQL.Append("FROM GarmentParts ");
                    sbSQL.Append("WHERE (GarmentParts = N'" + txeGarment.Text.Trim().Replace("'", "''") + "') ");
                    string strCHK = this.DBC.DBQuery(sbSQL).getString();
                    if (strCHK != "" && strCHK != txeID.Text.Trim())
                    {
                        chkDup = false;
                    }
                }
            }
            return chkDup;
        }

        private void txeGarment_Leave(object sender, EventArgs e)
        {
            if (txeGarment.Text.Trim() != "")
            {
                bool chkDup = chkDuplicate();
                if (chkDup == false)
                {
                    txeGarment.Text = "";
                    txeGarment.Focus();
                    FUNC.msgWarning("Duplicate garment parts. !! Please Change.");
                }
            }
        }

        private void bbiSave_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (txeGarment.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input garment parts.");
                txeGarment.Focus();
            }
            else
            {
                txeGarment.Text = txeGarment.Text.Trim();
                bool chkGMP = chkDuplicate();

                if (chkGMP == true)
                {
                    if (FUNC.msgQuiz("Confirm save data ?") == true)
                    {
                        StringBuilder sbSQL = new StringBuilder();
                        string strCREATE = UserLogin.OIDUser.ToString() != "" ? UserLogin.OIDUser.ToString() : "0";

                        if (lblStatus.Text == "* Add Garment")
                        {
                            sbSQL.Append("  INSERT INTO GarmentParts(GarmentParts, CreatedBy, CreatedDate) ");
                            sbSQL.Append("  VALUES(N'" + txeGarment.Text.Trim().Replace("'", "''") + "', '" + strCREATE + "', GETDATE()) ");
                        }
                        else if (lblStatus.Text == "* Edit Garment")
                        {
                            sbSQL.Append("  UPDATE GarmentParts SET ");
                            sbSQL.Append("      GarmentParts = N'" + txeGarment.Text.Trim().Replace("'", "''") + "' ");
                            sbSQL.Append("  WHERE(OIDGParts = '" + txeID.Text.Trim() + "') ");
                        }

                        //sbSQL.Append("IF NOT EXISTS(SELECT OIDGParts FROM GarmentParts WHERE OIDGParts = N'" + txeID.Text.Trim() + "') ");
                        //sbSQL.Append(" BEGIN ");
                        //sbSQL.Append("  INSERT INTO GarmentParts(GarmentParts, CreatedBy, CreatedDate) ");
                        //sbSQL.Append("  VALUES(N'" + txeGarment.Text.Trim().Replace("'", "''") + "', '" + strCREATE + "', GETDATE()) ");
                        //sbSQL.Append(" END ");
                        //sbSQL.Append("ELSE ");
                        //sbSQL.Append(" BEGIN ");
                        //sbSQL.Append("  UPDATE GarmentParts SET ");
                        //sbSQL.Append("      GarmentParts = N'" + txeGarment.Text.Trim().Replace("'", "''") + "' ");
                        //sbSQL.Append("  WHERE(OIDGParts = '" + txeID.Text.Trim() + "') ");
                        //sbSQL.Append(" END ");
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
                    txeGarment.Text = "";
                    txeGarment.Focus();
                    FUNC.msgWarning("Duplicate garment parts. !! Please Change.");
                }
            }
        }

        private void bbiExcel_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            string pathFile = new ObjSet.Folder(@"C:\MDS\Export\").GetPath() + "GarmentPartsList_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
            gvGarment.ExportToXlsx(pathFile);
            System.Diagnostics.Process.Start(pathFile);
        }

        private void gvGarment_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            
        }

        private void gvGarment_RowClick(object sender, RowClickEventArgs e)
        {
            if (gvGarment.IsFilterRow(e.RowHandle)) return;
            //lblStatus.Text = "* Edit Garment";
            //lblStatus.ForeColor = Color.Red;

            //txeID.Text = gvGarment.GetFocusedRowCellValue("No").ToString();
            //txeGarment.Text = gvGarment.GetFocusedRowCellValue("GarmentParts").ToString();

            //glueCREATE.EditValue = gvGarment.GetFocusedRowCellValue("CreatedBy").ToString();
            //txeDATE.Text = gvGarment.GetFocusedRowCellValue("CreatedDate").ToString();
        }

        private void gvGarment_DoubleClick(object sender, EventArgs e)
        {
            DXMouseEventArgs ea = e as DXMouseEventArgs;
            GridView view = sender as GridView;
            GridHitInfo info = view.CalcHitInfo(ea.Location);
            if (info.InRow || info.InRowCell)
            {
                GridView gv = gvGarment;
                lblStatus.Text = "* Edit Garment";
                lblStatus.ForeColor = Color.Red;

                txeID.Text = gv.GetFocusedRowCellValue("No").ToString();
                txeGarment.Text = gv.GetFocusedRowCellValue("GarmentParts").ToString();

                glueCREATE.EditValue = gv.GetFocusedRowCellValue("CreatedBy").ToString();
                txeDATE.Text = gv.GetFocusedRowCellValue("CreatedDate").ToString();
            }
        }

        private void bbiPrintPreview_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            gcGarment.ShowPrintPreview();
        }

        private void bbiPrint_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            gcGarment.Print();
        }

        private void txeGarment_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeID.Focus();
            }
        }

        private void gvGarment_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator) e.Info.DisplayText = (e.RowHandle + 1).ToString();
            gvGarment.IndicatorWidth = 40;
        }

       
    }
}