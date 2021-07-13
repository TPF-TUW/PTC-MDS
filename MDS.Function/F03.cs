using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Collections.Generic;
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
    public partial class F03 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private Functionality.Function FUNC = new Functionality.Function();
        private string dbDP = "Departments";
        private string dbBranch = "Branchs";
        List<DepartmentType> departmentTypes;
        public LogIn UserLogin { get; set; }
        public int Company { get; set; }

        public string ConnectionString { get; set; }
        string CONNECT_STRING = "";
        DatabaseConnect DBC;

        string tmpCode = "";
        string tmpName = "";

        public F03()
        {
            InitializeComponent();
            UserLookAndFeel.Default.StyleChanged += MyStyleChanged;
            departmentTypes = new List<DepartmentType>();
            departmentTypes.Add(new DepartmentType { name = "Admin", value = 0 });
            departmentTypes.Add(new DepartmentType { name = "Packing", value = 1 });
            departmentTypes.Add(new DepartmentType { name = "NeedleRoom", value= 2 });
            departmentTypes.Add(new DepartmentType { name = "Warehouse", value= 3 });
            departmentTypes.Add(new DepartmentType { name = "StoreFabric", value = 4 });
            departmentTypes.Add(new DepartmentType { name = "StoreAccessory", value = 5 });
            departmentTypes.Add(new DepartmentType { name = "Delivery", value = 6 });
            departmentTypes.Add(new DepartmentType { name = "FQA", value = 7 });
            departmentTypes.Add(new DepartmentType { name = "CMT", value = 8 });
            departmentTypes.Add(new DepartmentType { name = "Sales", value = 9 });
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
            sbSQL.Append("SELECT TOP (1) ReadWriteStatus FROM FunctionAccess WHERE (OIDUser = '" + UserLogin.OIDUser + "') AND(FunctionNo = 'F03') ");
            int chkReadWrite = this.DBC.DBQuery(sbSQL).getInt();
            if (chkReadWrite == 0)
                ribbonPageGroup1.Visible = false;

            sbSQL.Clear();
            sbSQL.Append("SELECT FullName, OIDUSER FROM Users ORDER BY OIDUSER ");
            new ObjDE.setGridLookUpEdit(glueCREATE, sbSQL, "FullName", "OIDUSER").getData();

            glueCREATE.EditValue = UserLogin.OIDUser;

            glueDPType.Properties.DataSource = departmentTypes;
            glueDPType.Properties.DisplayMember = "name";
            glueDPType.Properties.ValueMember = "value";

            
            NewData();
            LoadData();
        }

        private void LoadDEPT()
        {
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT DP.OIDDEPT AS ID, DP.Code AS [Department Code], DP.Name AS [Department Name], DP.DepartmentType AS [Department Type ID], DT.Name AS [Department Type], DP.OIDCOMPANY AS [Company ID], CP.Code AS [Company Code], ");
            sbSQL.Append("       CP.EngName AS[Company Name(En)], CP.THName AS[Company Name(Th)], DP.OIDBRANCH AS [Branch ID], BN.Code AS[Branch Code], BN.Name AS[Branch Name], DP.Status AS [Status ID], CASE WHEN DP.Status = 0 THEN 'Non Active' ELSE CASE WHEN DP.Status = 1 THEN 'Active' ELSE '' END END AS Status, DP.CreatedBy, DP.CreatedDate ");
            sbSQL.Append("FROM   " + this.dbDP + " AS DP LEFT OUTER JOIN ");
            sbSQL.Append("       Company AS CP ON DP.OIDCOMPANY = CP.OIDCOMPANY LEFT OUTER JOIN ");
            sbSQL.Append("       " + this.dbBranch + " AS BN ON DP.OIDBRANCH = BN.OIDBranch LEFT OUTER JOIN ");
            sbSQL.Append("       DepartmentType AS DT ON DP.DepartmentType = DT.Code ");
            sbSQL.Append("WHERE (DP.Code <> N'') ");
            if (glueCompany.Text.Trim() != "")
            {
                sbSQL.Append("AND (DP.OIDCOMPANY = '" + glueCompany.EditValue.ToString() + "') ");
            }
            if (glueBranch.Text.Trim() != "")
            {
                sbSQL.Append("AND (DP.OIDBRANCH = '" + glueBranch.EditValue.ToString() + "') ");
            }
            if (glueDPType.Text.Trim() != "")
            {
                sbSQL.Append("AND (DP.DepartmentType = '" + glueDPType.EditValue.ToString() + "') ");
            }
            sbSQL.Append("ORDER BY[Company ID], [Branch ID], ID ");
            new ObjDE.setGridControl(gcDP, gvDP, sbSQL).getData(false, false, false, true);
            gvDP.Columns[0].Visible = false;
            gvDP.Columns[3].Visible = false; //Department Type ID
            gvDP.Columns[5].Visible = false; //Company ID
            gvDP.Columns[9].Visible = false; //Branch ID
            gvDP.Columns[12].Visible = false; //Status ID

            gvDP.Columns["CreatedBy"].Visible = false;
            gvDP.Columns["CreatedDate"].Visible = false;
        }

        private void LoadData()
        {
            StringBuilder sbSQL = new StringBuilder();

            sbSQL.Clear();
            sbSQL.Append("SELECT Code AS [Company Code], EngName AS [Company Name (En)], THName AS [Company Name (Th)], OIDCOMPANY AS ID ");
            sbSQL.Append("FROM Company ");
            sbSQL.Append("WHERE (OIDCOMPANY = '" + this.Company + "') ");
            sbSQL.Append("ORDER BY OIDCOMPANY ");
            new ObjDE.setGridLookUpEdit(glueCompany, sbSQL, "Company Code", "ID").getData();
            glueCompany.Properties.View.PopulateColumns(glueCompany.Properties.DataSource);
            glueCompany.Properties.View.Columns["ID"].Visible = false;
            glueCompany.EditValue = this.Company;
            glueCompany.Enabled = false;
            glueCompany.BackColor = Color.White;
            glueCompany.ForeColor = Color.Navy;

            sbSQL.Clear();
            sbSQL.Append("SELECT Name AS [Department Type], Code AS ID ");
            sbSQL.Append("FROM DepartmentType ");
            sbSQL.Append("ORDER BY Code ");
            new ObjDE.setGridLookUpEdit(glueDPType, sbSQL, "Department Type", "ID").getData();
            glueDPType.Properties.View.PopulateColumns(glueDPType.Properties.DataSource);
            glueDPType.Properties.View.Columns["ID"].Visible = false;

            LoadDEPT();
        }

        private void NewData()
        {
            tmpCode = "";
            tmpName = "";
            txeName.Text = "";
            lblStatus.Text = "* Add Department";
            lblStatus.ForeColor = Color.Green;

            txeID.Text = this.DBC.DBQuery("SELECT CASE WHEN ISNULL(MAX(OIDDEPT), '') = '' THEN 1 ELSE MAX(OIDDEPT) + 1 END AS NewNo FROM " + this.dbDP).getString();
            txeCode.Text = "";
            txeName.Text = "";
            rgStatus.EditValue = 1;

            glueDPType.EditValue = "";
            glueCompany.EditValue = "";
            glueBranch.Properties.DataSource = null;

            glueCREATE.EditValue = UserLogin.OIDUser;
            txeCDATE.Text = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");

            glueCompany.Focus();
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

            string Branch = "";
            if (glueBranch.Text.Trim() != "")
            {
                Branch = glueBranch.EditValue.ToString();
            }

            string DPType = "";
            if (glueDPType.Text.Trim() != "")
            {
                DPType = glueDPType.EditValue.ToString();
            }

            if (lblStatus.Text == "* Add Department")
            {
                if (txeCode.Text.Trim() != "" || txeName.Text.Trim() != "")
                {
                     StringBuilder sbSQL = new StringBuilder();
                     if (txeCode.Text.Trim() != "" && chkDup == true)
                     {
                        if (txeCode.Text.Trim() != tmpCode)
                        {
                            sbSQL.Clear();
                            sbSQL.Append("SELECT TOP(1) Code FROM " + this.dbDP + " WHERE (OIDCOMPANY = '" + Company + "') AND (OIDBRANCH = '" + Branch + "') AND (DepartmentType = '" + DPType + "') AND (Code = N'" + txeCode.Text.Trim() + "') ");
                            string chkNo = this.DBC.DBQuery(sbSQL).getString();
                            if (chkNo != "")
                            {
                                txeCode.Text = "";
                                txeCode.Focus();
                                chkDup = false;
                                FUNC.msgWarning("Duplicate department code. !! Please Change.");
                            }
                            tmpCode = txeCode.Text.Trim();
                        }
                     }

                    if (txeName.Text.Trim() != "" && chkDup == true)
                    {
                        if (txeName.Text.Trim() != tmpName)
                        {
                            sbSQL.Clear();
                            sbSQL.Append("SELECT TOP(1) Code FROM " + this.dbDP + " WHERE (OIDCOMPANY = '" + Company + "') AND (OIDBRANCH = '" + Branch + "') AND (DepartmentType = '" + DPType + "') AND (Name = N'" + txeName.Text.Trim() + "') ");
                            string chkNo = this.DBC.DBQuery(sbSQL).getString();
                            if (chkNo != "")
                            {
                                txeName.Text = "";
                                txeName.Focus();
                                chkDup = false;
                                FUNC.msgWarning("Duplicate department name. !! Please Change.");
                            }
                            tmpName = txeName.Text.Trim();
                        }
                    }
                }
            }
            else if (lblStatus.Text == "* Edit Department")
            {
                if (txeCode.Text.Trim() != "" || txeName.Text.Trim() != "")
                {
                    StringBuilder sbSQL = new StringBuilder();
                    if (txeCode.Text.Trim() != "" && chkDup == true)
                    {
                        if (txeCode.Text.Trim() != tmpCode)
                        {
                            sbSQL.Clear();
                            sbSQL.Append("SELECT TOP(1) OIDDEPT FROM " + this.dbDP + " WHERE (OIDCOMPANY = '" + Company + "') AND (OIDBRANCH = '" + Branch + "') AND (DepartmentType = '" + DPType + "') AND (Code = N'" + txeCode.Text.Trim() + "') ");
                            //MessageBox.Show(sbSQL.ToString());
                            string chkNo = this.DBC.DBQuery(sbSQL).getString();
                            //MessageBox.Show(chkNo);
                            if (chkNo != "" && chkNo != txeID.Text.Trim())
                            {
                                txeCode.Text = "";
                                txeCode.Focus();
                                chkDup = false;
                                FUNC.msgWarning("Duplicate department code. !! Please Change.");
                            }
                            tmpCode = txeCode.Text.Trim();
                        }
                    }

                    if (txeName.Text.Trim() != "" && chkDup == true)
                    {
                        if (txeName.Text.Trim() != tmpName)
                        {
                            sbSQL.Clear();
                            sbSQL.Append("SELECT TOP(1) OIDDEPT FROM " + this.dbDP + " WHERE (OIDCOMPANY = '" + Company + "') AND (OIDBRANCH = '" + Branch + "') AND (DepartmentType = '" + DPType + "') AND (Name = N'" + txeName.Text.Trim() + "') ");
                            string chkNo = this.DBC.DBQuery(sbSQL).getString();
                            if (chkNo != "" && chkNo != txeID.Text.Trim())
                            {
                                txeName.Text = "";
                                txeName.Focus();
                                chkDup = false;
                                FUNC.msgWarning("Duplicate department name. !! Please Change.");
                            }
                            tmpName = txeName.Text.Trim();
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
            else if (glueBranch.Text.Trim() == "")
            {
                FUNC.msgWarning("Please select branch.");
                glueBranch.Focus();
            }
            else if (glueDPType.Text.Trim() == "")
            {
                FUNC.msgWarning("Please select department type.");
                glueDPType.Focus();
            }
            else if (txeCode.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input code.");
                txeCode.Focus();
            }
            else if (txeName.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input name.");
                txeName.Focus();
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
                        string Status = "NULL";
                        if (rgStatus.SelectedIndex != -1)
                        {
                            Status = rgStatus.Properties.Items[rgStatus.SelectedIndex].Value.ToString();
                        }

                        if (lblStatus.Text == "* Add Department")
                        {
                            sbSQL.Append("  INSERT INTO " + this.dbDP + "(Code, Name, DepartmentType, OIDCOMPANY, OIDBRANCH, Status, CreatedBy, CreatedDate) ");
                            sbSQL.Append("  VALUES(N'" + txeCode.Text.Trim().Replace("'", "''") + "', N'" + txeName.Text.Trim().Replace("'", "''") + "', '" + glueDPType.EditValue.ToString() + "', '" + glueCompany.EditValue.ToString() + "', '" + glueBranch.EditValue.ToString() + "', " + Status + ", '" + strCREATE + "', GETDATE()) ");
                        }
                        else if (lblStatus.Text == "* Edit Department")
                        {
                            sbSQL.Append("  UPDATE " + this.dbDP + " SET ");
                            sbSQL.Append("      Code=N'" + txeCode.Text.Trim().Replace("'", "''") + "', ");
                            sbSQL.Append("      Name=N'" + txeName.Text.Trim().Replace("'", "''") + "', ");
                            sbSQL.Append("      DepartmentType='" + glueDPType.EditValue.ToString() + "', ");
                            sbSQL.Append("      OIDCOMPANY='" + glueCompany.EditValue.ToString() + "', ");
                            sbSQL.Append("      OIDBRANCH='" + glueBranch.EditValue.ToString() + "', ");
                            sbSQL.Append("      Status=" + Status + " ");
                            sbSQL.Append("  WHERE (OIDDEPT = '" + txeID.Text.Trim() + "') ");
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
            string pathFile = new ObjSet.Folder(@"C:\MDS\Export\").GetPath() + "DepartmentList_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
            gvDP.ExportToXlsx(pathFile);
            System.Diagnostics.Process.Start(pathFile);
        }

        private void gvPTerm_RowClick(object sender, RowClickEventArgs e)
        {

        }

        private void bbiPrintPreview_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            gcDP.ShowPrintPreview();
        }

        private void bbiPrint_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            gcDP.Print();
        }

        private void F03_Shown(object sender, EventArgs e)
        {
            glueCompany.Focus();
        }


        private void gvDP_DoubleClick(object sender, EventArgs e)
        {
            DXMouseEventArgs ea = e as DXMouseEventArgs;
            GridView view = sender as GridView;
            GridHitInfo info = view.CalcHitInfo(ea.Location);
            if (info.InRow || info.InRowCell)
            {
                GridView gv = gvDP;
                lblStatus.Text = "* Edit Department";
                lblStatus.ForeColor = Color.Red;

                //string msg = "ID: " + gv.GetFocusedRowCellValue("ID").ToString() + "\n";
                //msg += "Company ID: " + gv.GetFocusedRowCellValue("Company ID").ToString() + "\n";
                //msg += "Branch ID: " + gv.GetFocusedRowCellValue("Branch ID").ToString() + "\n";
                //msg += "Department Type ID: " + gv.GetFocusedRowCellValue("Department Type ID").ToString() + "\n";
                //msg += "Department Code: " + gv.GetFocusedRowCellValue("Department Code").ToString() + "\n";
                //msg += "Department Name: " + gv.GetFocusedRowCellValue("Department Name").ToString() + "\n";
                //MessageBox.Show(msg);

                txeID.Text = gv.GetFocusedRowCellValue("ID").ToString();
                txeCode.Text = gv.GetFocusedRowCellValue("Department Code").ToString();
                txeName.Text = gv.GetFocusedRowCellValue("Department Name").ToString();

                tmpCode = txeCode.Text.Trim();
                tmpName = txeName.Text.Trim();

                rgStatus.EditValue = Convert.ToInt32(gv.GetFocusedRowCellValue("Status ID").ToString());
                glueDPType.EditValue = gv.GetFocusedRowCellValue("Department Type ID").ToString();
                
                glueCompany.EditValue = gv.GetFocusedRowCellValue("Company ID").ToString();
                glueBranch.EditValue = gv.GetFocusedRowCellValue("Branch ID").ToString();
                
               
                string CreatedBy = gv.GetFocusedRowCellValue("CreatedBy").ToString() == null ? "" : gv.GetFocusedRowCellValue("CreatedBy").ToString();
                glueCREATE.EditValue = CreatedBy;
                txeCDATE.Text = gv.GetFocusedRowCellValue("CreatedDate").ToString();
            }

        }

        private void gvDP_RowCellClick(object sender, RowCellClickEventArgs e)
        {
            if (gvDP.IsFilterRow(e.RowHandle)) return;
        }

        private void glueCompany_EditValueChanged(object sender, EventArgs e)
        {
            glueBranch.Properties.DataSource = null;


            if (glueCompany.Text.Trim() != "")
            {
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT Code AS [Branch Code], Name AS [Branch Name], CASE WHEN BranchType = 0 THEN 'Branch' ELSE CASE WHEN BranchType = 1 THEN 'Branch Sub Contract' ELSE '' END END AS [Branch Type], OIDBranch AS ID ");
                sbSQL.Append("FROM  " + this.dbBranch + " ");
                sbSQL.Append("WHERE (OIDCOMPANY = '" + glueCompany.EditValue.ToString() + "') ");
                sbSQL.Append("ORDER BY [Branch Code] ");

                new ObjDE.setGridLookUpEdit(glueBranch, sbSQL, "Branch Code", "ID").getData();
                glueBranch.Properties.View.PopulateColumns(glueBranch.Properties.DataSource);
                glueBranch.Properties.View.Columns["ID"].Visible = false;

                bool chkDup = chkDuplicate();
                if (chkDup == true)
                {
                    glueBranch.Focus();
                }
                
            }

            LoadDEPT();
        }

        private void glueBranch_EditValueChanged(object sender, EventArgs e)
        {
            if (glueBranch.Text.Trim() != "")
            {
                bool chkDup = chkDuplicate();
                if (chkDup == true)
                {
                    glueDPType.Focus();
                }
            }
            LoadDEPT();
        }


        private void glueDPType_EditValueChanged(object sender, EventArgs e)
        {
            if (glueDPType.Text.Trim() != "")
            {
                bool chkDup = chkDuplicate();
                if (chkDup == true)
                {
                    txeCode.Focus();
                }
            }
            LoadDEPT();
        }


        private void txeCode_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeName.Focus();
            }
        }

        private void txeCode_Leave(object sender, EventArgs e)
        {
            if (txeCode.Text.Trim() != "")
            {
                txeCode.Text = txeCode.Text.Trim();
                bool chkDup = chkDuplicate();
                if (chkDup == true)
                {
                    txeName.Focus();
                }
            }
        }

        private void txeName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                rgStatus.Focus();
            }
        }

        private void txeName_Leave(object sender, EventArgs e)
        {
            if (txeName.Text.Trim() != "")
            {
                txeName.Text = txeName.Text.Trim();
                bool chkDup = chkDuplicate();
                if (chkDup == true)
                {
                    rgStatus.Focus();
                }
            }
        }

        private void gvDP_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator) e.Info.DisplayText = (e.RowHandle + 1).ToString();
            gvDP.IndicatorWidth = 40;
        }
    }

    public class DepartmentType
    {
        public string name { get;set; }
        public int value { get; set; }
    }
}