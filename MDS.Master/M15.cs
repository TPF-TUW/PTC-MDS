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
using DevExpress.Utils;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;

namespace MDS.Master
{
    public partial class M15 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private Functionality.Function FUNC = new Functionality.Function();
        public LogIn UserLogin { get; set; }
        public int Company { get; set; }
        public string ConnectionString { get; set; }
        string CONNECT_STRING = "";
        DatabaseConnect DBC;

        public M15()
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
            sbSQL.Append("SELECT TOP (1) ReadWriteStatus FROM FunctionAccess WHERE (OIDUser = '" + UserLogin.OIDUser + "') AND(FunctionNo = 'M15') ");
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
            sbSQL.Append("SELECT OIDPayment AS No, Name, Description, DuedateCalculation, Status, CreatedBy, CreatedDate, CASE WHEN Status = 0 THEN 'Non Active' ELSE 'Active' END AS StatusName ");
            sbSQL.Append("FROM PaymentTerm ");
            sbSQL.Append("ORDER BY OIDPayment ");
            new ObjDE.setGridControl(gcPTerm, gvPTerm, sbSQL).getData(false, false, false, true);
            gvPTerm.Columns[0].Visible = false;
            gvPTerm.Columns[4].Visible = false;
            gvPTerm.Columns[5].Visible = false;
            gvPTerm.Columns[6].Visible = false;
        }

        private void NewData()
        {
            txeName.Text = "";
            lblStatus.Text = "* Add Payment Term";
            lblStatus.ForeColor = Color.Green;

            txeID.Text = this.DBC.DBQuery("SELECT CASE WHEN ISNULL(MAX(OIDPayment), '') = '' THEN 1 ELSE MAX(OIDPayment) + 1 END AS NewNo FROM PaymentTerm").getString();
            txeDescription.Text = "";
            txeDueDate.Text = "";
            rgStatus.EditValue = 1;

            glueCREATE.EditValue = UserLogin.OIDUser;
            txeDATE.Text = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");

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

        private void selectStatus(int value)
        {
            rgStatus.EditValue = value;
        }

        private bool chkDuplicate()
        {
            bool chkDup = true;
            if (txeName.Text != "")
            {
                txeName.Text = txeName.Text.Trim();
                if (txeName.Text.Trim() != "" && lblStatus.Text == "* Add Payment Term")
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT TOP(1) Name FROM PaymentTerm WHERE (Name = N'" + txeName.Text.Trim() + "') ");
                    if (this.DBC.DBQuery(sbSQL).getString() != "")
                    {
                        chkDup = false;
                    }
                }
                else if (txeName.Text.Trim() != "" && lblStatus.Text == "* Edit Payment Term")
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT TOP(1) OIDPayment ");
                    sbSQL.Append("FROM PaymentTerm ");
                    sbSQL.Append("WHERE (Name = N'" + txeName.Text.Trim() + "') ");
                    string strCHK = this.DBC.DBQuery(sbSQL).getString();
                    if (strCHK != "" && strCHK != txeID.Text.Trim())
                    {
                        chkDup = false;
                    }
                }
            }
            return chkDup;
        }

        private void txeName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeDescription.Focus();
            }
        }

        private void txeName_LostFocus(object sender, EventArgs e)
        {
            
        }

        private void txeDescription_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeDueDate.Focus();
            }
        }

        private void txeDueDate_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                rgStatus.Focus();
            }
        }

        private void gvPTerm_RowStyle(object sender, RowStyleEventArgs e)
        {
            
        }

        private void bbiSave_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (txeName.Text.Trim() == "")
            {
                FUNC.msgWarning("Please name.");
                txeName.Focus();
            }
            else if (txeDescription.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input description.");
                txeDescription.Focus();
            }
            else
            {
                if (FUNC.msgQuiz("Confirm save data ?") == true)
                {
                    StringBuilder sbSQL = new StringBuilder();
                    string strCREATE = UserLogin.OIDUser.ToString() != "" ? UserLogin.OIDUser.ToString() : "0";

                    bool chkGMP = chkDuplicate();
                    if (chkGMP == true)
                    {
                        string Status = rgStatus.EditValue.ToString();

                       // MessageBox.Show(Status);
                        if (lblStatus.Text == "* Add Payment Term")
                        {
                            sbSQL.Append("  INSERT INTO PaymentTerm(Name, Description, DueDateCalculation, Status, CreatedBy, CreatedDate) ");
                            sbSQL.Append("  VALUES(N'" + txeName.Text.Trim().Replace("'", "''") + "', N'" + txeDescription.Text.Trim().Replace("'", "''") + "', N'" + txeDueDate.Text.Trim().Replace("'", "''") + "', " + Status + ", '" + strCREATE + "', GETDATE()) ");
                        }
                        else if (lblStatus.Text == "* Edit Payment Term")
                        {
                            sbSQL.Append("  UPDATE PaymentTerm SET ");
                            sbSQL.Append("      Name=N'" + txeName.Text.Trim().Replace("'", "''") + "', ");
                            sbSQL.Append("      Description=N'" + txeDescription.Text.Trim().Replace("'", "''") + "', ");
                            sbSQL.Append("      DueDateCalculation=N'" + txeDueDate.Text.Trim().Replace("'", "''") + "', ");
                            sbSQL.Append("      Status=" + Status + " ");
                            sbSQL.Append("  WHERE(OIDPayment = '" + txeID.Text.Trim() + "') ");
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
                    else
                    {
                        txeName.Text = "";
                        txeName.Focus();
                        FUNC.msgWarning("Duplicate payment term. !! Please Change.");
                    }
                }
            }
        }

        private void bbiExcel_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            string pathFile = new ObjSet.Folder(@"C:\MDS\Export\").GetPath() + "PaymentTermList_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
            gvPTerm.ExportToXlsx(pathFile);
            System.Diagnostics.Process.Start(pathFile);
        }

        private void gvPTerm_RowClick(object sender, RowClickEventArgs e)
        {
            if (gvPTerm.IsFilterRow(e.RowHandle)) return;
        }

        private void gvPTerm_DoubleClick(object sender, EventArgs e)
        {
            DXMouseEventArgs ea = e as DXMouseEventArgs;
            GridView view = sender as GridView;
            GridHitInfo info = view.CalcHitInfo(ea.Location);
            if (info.InRow || info.InRowCell)
            {
                GridView gv = gvPTerm;
                lblStatus.Text = "* Edit Payment Term";
                lblStatus.ForeColor = Color.Red;

                txeID.Text = gv.GetFocusedRowCellValue("No").ToString();
                txeName.Text = gv.GetFocusedRowCellValue("Name").ToString();
                txeDescription.Text = gv.GetFocusedRowCellValue("Description").ToString();
                txeDueDate.Text = gv.GetFocusedRowCellValue("DuedateCalculation").ToString();

                int status = -1;
                if (gv.GetFocusedRowCellValue("Status").ToString() != "")
                {
                    status = Convert.ToInt32(gv.GetFocusedRowCellValue("Status").ToString());
                }

                selectStatus(status);

                string CreatedBy = gv.GetFocusedRowCellValue("CreatedBy").ToString() == null ? "" : gv.GetFocusedRowCellValue("CreatedBy").ToString();
                glueCREATE.EditValue = CreatedBy;
                txeDATE.Text = gv.GetFocusedRowCellValue("CreatedDate").ToString();
            }
        }

        private void bbiPrintPreview_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            gcPTerm.ShowPrintPreview();
        }

        private void bbiPrint_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            gcPTerm.Print();
        }

        private void txeName_Leave(object sender, EventArgs e)
        {
            if (txeName.Text.Trim() != "")
            {
                txeName.Text = txeName.Text.ToUpper().Trim();
                bool chkDup = chkDuplicate();
                if (chkDup == false)
                {
                    txeName.Text = "";
                    txeName.Focus();
                    FUNC.msgWarning("Duplicate payment term. !! Please Change.");
                }
                else
                {
                    txeDescription.Focus();
                }
            }
        }

        private void gvPTerm_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator) e.Info.DisplayText = (e.RowHandle + 1).ToString();
            gvPTerm.IndicatorWidth = 40;
        }

       
    }
}