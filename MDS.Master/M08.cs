using System;
using System.Text;
using System.Windows.Forms;
using DevExpress.LookAndFeel;
using DevExpress.Utils.Extensions;
using DBConnect;
using System.Drawing;
using System.Data;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraEditors;
using TheepClass;
using DevExpress.Utils;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;

namespace MDS.Master
{
    public partial class M08 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private Functionality.Function FUNC = new Functionality.Function();
        private string selCode = "";
        public LogIn UserLogin { get; set; }

        public string ConnectionString { get; set; }
        string CONNECT_STRING = "";
        DatabaseConnect DBC;


        public M08()
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
            sbSQL.Append("SELECT TOP (1) ReadWriteStatus FROM FunctionAccess WHERE (OIDUser = '" + UserLogin.OIDUser + "') AND(FunctionNo = 'M08') ");
            int chkReadWrite = this.DBC.DBQuery(sbSQL).getInt();
            if (chkReadWrite == 0)
                ribbonPageGroup1.Visible = false;

            sbSQL.Clear();
            sbSQL.Append("SELECT FullName, OIDUSER FROM Users ORDER BY OIDUSER ");
            new ObjDE.setGridLookUpEdit(glueCREATE, sbSQL, "FullName", "OIDUSER").getData();

            glueCREATE.EditValue = UserLogin.OIDUser;

            //glueLineName.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard;
            //glueLineName.Properties.AcceptEditorTextAsNewValue = DevExpress.Utils.DefaultBoolean.True;

            bbiNew.PerformClick();
        }

        private void LoadData()
        {
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT LN.OIDLINE AS ID, LN.LINENAME, B.Name AS Branch, LN.Branch AS BranchID ");
            sbSQL.Append("FROM   LineNumber AS LN INNER JOIN ");
            sbSQL.Append("       Branchs AS B ON LN.Branch = B.OIDBranch ");
            sbSQL.Append("ORDER BY LN.LINENAME ");
            new ObjDE.setSearchLookUpEdit(glueLineName, sbSQL, "LINENAME", "ID").getData(true);
            glueLineName.Properties.View.PopulateColumns(glueLineName.Properties.DataSource);
            glueLineName.Properties.View.Columns["ID"].Visible = false;
            glueLineName.Properties.View.Columns["BranchID"].Visible = false;

            sbSQL.Clear();
            sbSQL.Append("SELECT OIDUser AS ID, UserName, FullName ");
            sbSQL.Append("FROM [User] ");
            sbSQL.Append("ORDER BY UserName ");
            new ObjDE.setSearchLookUpEdit(slueInCharge, sbSQL, "UserName", "ID").getData(true);
            slueInCharge.Properties.View.PopulateColumns(slueInCharge.Properties.DataSource);
            slueInCharge.Properties.View.Columns["ID"].Visible = false;

            sbSQL.Clear();
            sbSQL.Append("SELECT B.Code, B.Name AS Branch, C.Code AS CompnayCode, C.EngName AS CompanyName, ");
            sbSQL.Append("       CASE WHEN B.BranchType = 0 THEN 'Branch' ELSE CASE WHEN B.BranchType = 1 THEN 'Branch Sub Contract' ELSE '' END END AS Type, B.OIDBranch AS ID ");
            sbSQL.Append("FROM   Branchs AS B INNER JOIN ");
            sbSQL.Append("       Company AS C ON B.OIDCOMPANY = C.OIDCOMPANY ");
            sbSQL.Append("ORDER BY B.Code ");
            new ObjDE.setSearchLookUpEdit(glueBranch, sbSQL, "Branch", "ID").getData(true);
            glueBranch.Properties.View.PopulateColumns(glueBranch.Properties.DataSource);
            glueBranch.Properties.View.Columns["ID"].Visible = false;
            glueBranch.Properties.View.Columns["Code"].Visible = false;

            sbSQL.Clear();
            sbSQL.Append("SELECT OIDCUST AS ID, Code, ShortName, Name ");
            sbSQL.Append("FROM Customer ");
            sbSQL.Append("ORDER BY ShortName ");
            new ObjDE.setSearchLookUpEdit(slueCustomer, sbSQL, "ShortName", "ID").getData(true);
            slueCustomer.Properties.View.PopulateColumns(slueCustomer.Properties.DataSource);
            slueCustomer.Properties.View.Columns["ID"].Visible = false;

            sbSQL.Clear();
            sbSQL.Append("SELECT OIDGCATEGORY, CategoryName ");
            sbSQL.Append("FROM GarmentCategory ");
            sbSQL.Append("ORDER BY OIDGCATEGORY ");
            DataTable drCategory = this.DBC.DBQuery(sbSQL).getDataTable();
            clbCategory.ValueMember = "OIDGCATEGORY";
            clbCategory.DisplayMember = "CategoryName";
            clbCategory.DataSource = drCategory;

            LoadLineCategory();

        }

        private void NewData()
        {
            glueLineName.EditValue = "";
            lblStatus.Text = "* Add Line";
            lblStatus.ForeColor = Color.Green;

            //txeID.Text = this.DBC.DBQuery("SELECT CASE WHEN ISNULL(MAX(OIDGParts), '') = '' THEN 1 ELSE MAX(OIDGParts) + 1 END AS NewNo FROM GarmentParts").getString();
            txeID.Text = "";

            slueInCharge.EditValue = "";
            glueBranch.EditValue = "";
            slueCustomer.EditValue = "";

            glueCREATE.EditValue = UserLogin.OIDUser;
            txeCDATE.Text = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
            selCode = "";
            //txeID.Focus();
        }

        private void LoadLineCategory()
        {
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT PL.OIDLine AS LineID, LN.LINENAME AS LineName, PL.OIDUSER AS InChangeID, US.UserName AS InChange, PL.Branch AS BranchID, BN.Name AS Branch, PL.OIDCUST AS CustomerID, CUS.ShortName AS Customer, PL.OIDCATEGORY AS CategoryID, GC.CategoryName, PL.CreatedBy, PL.CreatedDate ");
            sbSQL.Append("FROM ProductionLine AS PL INNER JOIN ");
            sbSQL.Append("     LineNumber AS LN ON PL.OIDLine = LN.OIDLine INNER JOIN ");
            sbSQL.Append("     [User] AS US ON PL.OIDUSER = US.OIDUser INNER JOIN ");
            sbSQL.Append("     Customer AS CUS ON PL.OIDCUST = CUS.OIDCUST INNER JOIN ");
            sbSQL.Append("     Branchs AS BN ON PL.Branch = BN.OIDBranch INNER JOIN ");
            sbSQL.Append("     GarmentCategory AS GC ON PL.OIDCATEGORY = GC.OIDGCATEGORY ");
            sbSQL.Append("WHERE (PL.OIDLine <> '') ");
            if (glueLineName.Text.Trim() != "")
                sbSQL.Append("AND (PL.OIDLine='" + glueLineName.EditValue.ToString() + "') ");
            if (glueBranch.Text.Trim() != "")
                sbSQL.Append("AND (PL.Branch='" + glueBranch.EditValue.ToString() + "') ");
            if (slueCustomer.Text.Trim() != "")
                sbSQL.Append("AND (PL.OIDCUST='" + slueCustomer.EditValue.ToString() + "') ");
            //if (txeID.Text.Trim() != "" && glueBranch.Text.Trim() != "" && slueCustomer.Text.Trim() != "")
            //{
            //    sbSQL.Append("WHERE (PL.OIDLine='" + txeID.Text.Trim() + "') AND (PL.Branch='" + glueBranch.EditValue.ToString() + "') AND (PL.OIDCUST='" + slueCustomer.EditValue.ToString() + "') ");
            //}
            sbSQL.Append("ORDER BY LN.LINENAME, PL.Branch, PL.OIDCUST, PL.OIDCATEGORY ");
            //MessageBox.Show(sbSQL.ToString());
            new ObjDE.setGridControl(gcLine, gvLine, sbSQL).getData(false, false, false, true);
            gvLine.Columns["LineID"].Visible = false;
            gvLine.Columns["InChangeID"].Visible = false;
            gvLine.Columns["BranchID"].Visible = false;
            gvLine.Columns["CustomerID"].Visible = false;
            gvLine.Columns["CategoryID"].Visible = false;
            gvLine.Columns["CreatedBy"].Visible = false;
            gvLine.Columns["CreatedDate"].Visible = false;
        }

        private void bbiNew_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            LoadData();
            NewData();
        }

        private void gvGarment_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {
            
        }

        private void glueLineName_EditValueChanged(object sender, EventArgs e)
        {
            CheckLine();
            LoadCategory();
            //if (glueLineName.Text.Trim() == "")
            //{
            //    lblStatus.Text = "* Add Line";
            //    lblStatus.ForeColor = Color.Green;
            //    txeID.Text = "";
            //}

            //if (glueLineName.Text.Trim() != "" && glueLineName.Text.ToUpper().Trim() != selCode)
            //{
            //    glueLineName.Text = glueLineName.Text.ToUpper().Trim();
            //    selCode = glueLineName.Text;
            //    LoadCode(glueLineName.Text);
            //    //MessageBox.Show(glueCode.Text);
            //}
        }

        private void glueLineName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                slueInCharge.Focus();
            }
        }

        private void glueLineName_LostFocus(object sender, EventArgs e)
        {
            //if (glueLineName.Text.Trim() == "")
            //{
            //    lblStatus.Text = "* Add Line";
            //    lblStatus.ForeColor = Color.Green;
            //    txeID.Text = "";
            //}

            //if (glueLineName.Text.Trim() != "" && glueLineName.Text.ToUpper().Trim() != selCode)
            //{
            //    glueLineName.Text = glueLineName.Text.ToUpper().Trim();
            //    selCode = glueLineName.Text;
            //    LoadCode(glueLineName.Text);
            //    //MessageBox.Show(glueCode.Text);
            //}
        }

        private void CheckLine()
        {
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT OIDGCATEGORY, CategoryName ");
            sbSQL.Append("FROM GarmentCategory ");
            sbSQL.Append("ORDER BY OIDGCATEGORY ");
            DataTable drCategory = this.DBC.DBQuery(sbSQL).getDataTable();
            clbCategory.ValueMember = "OIDGCATEGORY";
            clbCategory.DisplayMember = "CategoryName";
            clbCategory.DataSource = drCategory;

            //sbSQL.Clear();
            //sbSQL.Append("SELECT OIDLINE ");
            //sbSQL.Append("FROM LineNumber ");
            //sbSQL.Append("WHERE (LINENAME = N'" + glueLineName.Text.Trim() + "') AND (Branch = '" + glueBranch.EditValue.ToString() + "') ");
            txeID.Text = glueLineName.EditValue.ToString();
            if (txeID.Text.Trim() == "")
            {
                lblStatus.Text = "* Add Line";
                lblStatus.ForeColor = Color.Green;
                LoadCategory();
            }
            else
            {
                lblStatus.Text = "* Edit Line";
                lblStatus.ForeColor = Color.Red;
                LoadCategory();
            }
        }

        private void LoadCategory()
        {
            //Clear Check Category
            for (int i = 0; i < clbCategory.ItemCount; i++)
            {
                clbCategory.SetItemCheckState(i, CheckState.Unchecked);
            }

            if (glueBranch.Text.Trim() != "" && glueLineName.Text.Trim() != "" && slueCustomer.Text.Trim() != "")
            {
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT OIDCATEGORY ");
                sbSQL.Append("FROM ProductionLine ");
                sbSQL.Append("WHERE (OIDLine = '" + txeID.Text.Trim() + "') ");
                sbSQL.Append("AND (Branch = '" + glueBranch.EditValue.ToString() + "') ");
                sbSQL.Append("AND (OIDCUST = '" + slueCustomer.EditValue.ToString() + "') ");
                sbSQL.Append("ORDER BY OIDCATEGORY ");
                //MessageBox.Show(sbSQL.ToString());
                DataTable dtQC = this.DBC.DBQuery(sbSQL).getDataTable();

                foreach (DataRow row in dtQC.Rows)
                {
                    for (int i = 0; i < clbCategory.ItemCount; i++)
                    {
                        if (row["OIDCATEGORY"].ToString() == clbCategory.GetItemValue(i).ToString())
                        {
                            clbCategory.SetItemCheckState(i, CheckState.Checked);
                            break;
                        }
                    }
                }
            }

            LoadLineCategory();
        }

        private void LoadCode(string strCODE)
        {

            string BRANCH = "";
            if (glueLineName.Properties.View.GetFocusedRowCellValue("BranchID") != null)
            {
                BRANCH = glueLineName.Properties.View.GetFocusedRowCellValue("BranchID").ToString();
            }
            glueBranch.EditValue = BRANCH;

            CheckLine();

        }

        private void glueLineName_Closed(object sender, DevExpress.XtraEditors.Controls.ClosedEventArgs e)
        {
            glueLineName.Focus();
            slueInCharge.Focus();
        }

        private void glueLineName_ProcessNewValue(object sender, DevExpress.XtraEditors.Controls.ProcessNewValueEventArgs e)
        {
            GridLookUpEdit gridLookup = sender as GridLookUpEdit;
            if (e.DisplayValue == null) return;
            string newValue = e.DisplayValue.ToString();
            if (newValue == String.Empty) return;
        }

        private void glueBranch_EditValueChanged(object sender, EventArgs e)
        {
            if (glueBranch.Text.Trim() != "")
            {
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT LN.OIDLINE AS ID, LN.LINENAME, B.Name AS Branch, LN.Branch AS BranchID ");
                sbSQL.Append("FROM   LineNumber AS LN INNER JOIN ");
                sbSQL.Append("       Branchs AS B ON LN.Branch = B.OIDBranch ");
                sbSQL.Append("WHERE (B.OIDBranch = '" + glueBranch.EditValue.ToString() + "') ");
                sbSQL.Append("ORDER BY LN.OIDLINE ");
                new ObjDE.setSearchLookUpEdit(glueLineName, sbSQL, "LINENAME", "ID").getData(true);
            }
            CheckLine();
            LoadCategory();

        }

        private void slueCustomer_EditValueChanged(object sender, EventArgs e)
        {
            LoadCategory();
        }

        private void gvLine_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            
        }

        private void bbiSave_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (glueLineName.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input line name.");
                glueLineName.Focus();
            }
            else if (slueInCharge.Text.Trim() == "")
            {
                FUNC.msgWarning("Please select in-charge.");
                slueInCharge.Focus();
            }
            else if (glueBranch.Text.Trim() == "")
            {
                FUNC.msgWarning("Please select branch.");
                glueBranch.Focus();
            }
            else if (slueCustomer.Text.Trim() == "")
            {
                FUNC.msgWarning("Please select customer.");
                slueCustomer.Focus();
            }
            else
            {
                if (FUNC.msgQuiz("Confirm save data ?") == true)
                {
                    StringBuilder sbSQL = new StringBuilder();

                    string strCREATE = UserLogin.OIDUser.ToString() != "" ? UserLogin.OIDUser.ToString() : "0";

                    //******** save LineNumber table ************
                    //string LINENAME = glueLineName.Text.ToUpper().Trim();
                    string LINENAME = glueLineName.EditValue.ToString();
                    string BRANCHID = glueBranch.EditValue.ToString();

                    if (lblStatus.Text == "* Add Line")
                    {
                        sbSQL.Append(" INSERT INTO LineNumber(LINENAME, Branch) ");
                        sbSQL.Append("  VALUES(N'" + LINENAME + "', '" + BRANCHID + "') ");
                    }
                    else if (lblStatus.Text == "* Edit Line")
                    {
                        sbSQL.Append(" UPDATE LineNumber SET ");
                        sbSQL.Append("  LINENAME = N'" + LINENAME + "', Branch = '" + BRANCHID + "' ");
                        sbSQL.Append(" WHERE (OIDLINE = '" + txeID.Text.Trim() + "') ");
                    }

                    if (sbSQL.Length > 0)
                    {
                        try
                        {
                            bool saveLine = this.DBC.DBQuery(sbSQL).runSQL();
                            if (saveLine == true)
                            {
                                sbSQL.Clear();
                                sbSQL.Append("SELECT OIDLINE FROM LineNumber WHERE (LINENAME = N'" + LINENAME + "') AND (Branch = '" + BRANCHID + "') ");
                                string LINEID = this.DBC.DBQuery(sbSQL).getString();

                                //******** save ProductionLine table ********
                                sbSQL.Clear();
                                string strCATEGORY = "";
                                int iCQC = 0;
                                foreach (DataRowView item in clbCategory.CheckedItems)
                                {
                                    if (iCQC != 0)
                                    {
                                        strCATEGORY += ", ";
                                    }
                                    strCATEGORY += "'" + item["OIDGCATEGORY"].ToString() + "'";
                                    sbSQL.Append("IF NOT EXISTS(SELECT OIDPDLINE FROM ProductionLine WHERE (OIDLine = '" + LINEID + "') AND (Branch = '" + BRANCHID + "') AND (OIDCUST = '" + slueCustomer.EditValue.ToString() + "') AND (OIDCATEGORY = '" + item["OIDGCATEGORY"].ToString() + "')) ");
                                    sbSQL.Append(" BEGIN ");
                                    sbSQL.Append("  INSERT INTO ProductionLine(OIDLine, OIDUSER, Branch, OIDCUST, OIDCATEGORY, CreatedBy, CreatedDate) ");
                                    sbSQL.Append("  VALUES('" + LINEID + "', '" + slueInCharge.EditValue.ToString() + "', '" + BRANCHID + "', '" + slueCustomer.EditValue.ToString() + "', '" + item["OIDGCATEGORY"].ToString() + "', '" + strCREATE + "', GETDATE()) ");
                                    sbSQL.Append(" END ");
                                    iCQC++;
                                }

                                if (strCATEGORY == "")
                                {
                                    sbSQL.Append("DELETE FROM ProductionLine WHERE (OIDLine = '" + LINEID + "') AND (Branch = '" + BRANCHID + "') AND (OIDCUST = '" + slueCustomer.EditValue.ToString() + "')  ");
                                }
                                else
                                {
                                    sbSQL.Append("DELETE FROM ProductionLine WHERE (OIDLine = '" + LINEID + "') AND (Branch = '" + BRANCHID + "') AND (OIDCUST = '" + slueCustomer.EditValue.ToString() + "') AND (OIDCATEGORY NOT IN (" + strCATEGORY + "))  ");
                                }

                                if (sbSQL.Length > 0)
                                {
                                    try
                                    {
                                        bool chkSAVECAT = this.DBC.DBQuery(sbSQL).runSQL();
                                        if (chkSAVECAT == true)
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
                        catch (Exception)
                        { }
                    }

                }
            }
        }

        private void bbiExcel_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            string pathFile = new ObjSet.Folder(@"C:\MDS\Export\").GetPath() + "LineList_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
            gvLine.ExportToXlsx(pathFile);
            System.Diagnostics.Process.Start(pathFile);
        }

        private void ribbonControl_Click(object sender, EventArgs e)
        {

        }

        private void gvLine_RowClick(object sender, RowClickEventArgs e)
        {
            if (gvLine.IsFilterRow(e.RowHandle)) return;
            //lblStatus.Text = "* Edit Line";
            //lblStatus.ForeColor = Color.Red;

            //string strLineName = gvLine.GetFocusedRowCellValue("LineName").ToString();
            //string strLineID = gvLine.GetFocusedRowCellValue("LineID").ToString();
            //string strInChargeID = gvLine.GetFocusedRowCellValue("InChangeID").ToString();
            //string strBranchID = gvLine.GetFocusedRowCellValue("BranchID").ToString();
            //string CusID = gvLine.GetFocusedRowCellValue("CustomerID").ToString();

            //glueLineName.EditValue = strLineName;
            //txeID.Text = strLineID;
            //slueInCharge.EditValue = strInChargeID;
            //glueBranch.EditValue = strBranchID;
            //slueCustomer.EditValue = CusID;

            //glueCREATE.EditValue = gvLine.GetFocusedRowCellValue("CreatedBy").ToString();
            //txeCDATE.Text = gvLine.GetFocusedRowCellValue("CreatedDate").ToString();
        }

        private void gvLine_DoubleClick(object sender, EventArgs e)
        {
            DXMouseEventArgs ea = e as DXMouseEventArgs;
            GridView view = sender as GridView;
            GridHitInfo info = view.CalcHitInfo(ea.Location);
            if (info.InRow || info.InRowCell)
            {
                GridView gv = gvLine;
                lblStatus.Text = "* Edit Line";
                lblStatus.ForeColor = Color.Red;

                string strLineName = gv.GetFocusedRowCellValue("LineName").ToString();
                string strLineID = gv.GetFocusedRowCellValue("LineID").ToString();
                string strInChargeID = gv.GetFocusedRowCellValue("InChangeID").ToString();
                string strBranchID = gv.GetFocusedRowCellValue("BranchID").ToString();
                string CusID = gv.GetFocusedRowCellValue("CustomerID").ToString();
                //MessageBox.Show(strLineID);
                
                txeID.Text = strLineID;
                glueBranch.EditValue = strBranchID;
                glueLineName.EditValue = strLineID;
                slueInCharge.EditValue = strInChargeID;
                slueCustomer.EditValue = CusID;

                glueCREATE.EditValue = gv.GetFocusedRowCellValue("CreatedBy").ToString();
                txeCDATE.Text = gv.GetFocusedRowCellValue("CreatedDate").ToString();
            }

        }

        private void glueLineName_Leave(object sender, EventArgs e)
        {
            if (glueLineName.Text.Trim() == "")
            {
                lblStatus.Text = "* Add Line";
                lblStatus.ForeColor = Color.Green;
                txeID.Text = "";
            }

            if (glueLineName.Text.Trim() != "" && glueLineName.Text.ToUpper().Trim() != selCode)
            {
                //glueLineName.Text = glueLineName.Text.ToUpper().Trim();
                selCode = glueLineName.Text.ToUpper().Trim();
                LoadCode(glueLineName.EditValue.ToString());
                //MessageBox.Show(glueCode.Text);
            }
        }

        private void gvLine_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator) e.Info.DisplayText = (e.RowHandle + 1).ToString();
            gvLine.IndicatorWidth = 40;
        }

        private void sbADD_Click(object sender, EventArgs e)
        {
            if (glueBranch.Text.Trim() == "")
            {
                FUNC.msgWarning("Please select branch before.\nกรุณาเลือกสาขาก่อน");
                glueBranch.Focus();
            }
            else
            {
                if (lblStatus.Text == "* Add Line")
                {
                    var frm = new M08_01(this.DBC, UserLogin.OIDUser, glueBranch.EditValue.ToString(), glueBranch.Text, "");
                    frm.ShowDialog(this);
                }
                else if (lblStatus.Text == "* Edit Line")
                {
                    var frm = new M08_01(this.DBC, UserLogin.OIDUser, glueBranch.EditValue.ToString(), glueBranch.Text, txeID.Text.Trim());
                    frm.ShowDialog(this);
                }

                
            }
        }
    }
}