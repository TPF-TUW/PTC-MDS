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
using System.IO;
using DevExpress.XtraEditors.DXErrorProvider;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraGrid.Views.Grid.Drawing;
using DevExpress.XtraGrid.Skins;
using DevExpress.Utils.Drawing;
using System.Data;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraEditors.ViewInfo;
using System.Collections;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraEditors.Controls;
using TheepClass;


namespace MDS.Function
{
    public partial class F05 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private Functionality.Function FUNC = new Functionality.Function();
        StringBuilder sbGROUP = new StringBuilder();
        ArrayList arlFunc;
        DataTable dtFuncName;
        public LogIn UserLogin { get; set; }
        public int Company { get; set; }

        public string ConnectionString { get; set; }
        string CONNECT_STRING = "";
        DatabaseConnect DBC;


        public F05()
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
            sbSQL.Append("SELECT TOP (1) ReadWriteStatus FROM FunctionAccess WHERE (OIDUser = '" + UserLogin.OIDUser + "') AND(FunctionNo = 'F05') ");
            int chkReadWrite = this.DBC.DBQuery(sbSQL).getInt();
            if (chkReadWrite == 0)
                ribbonPageGroup1.Visible = false;

            sbGROUP.Append("SELECT '0' AS [ID], 'Master' AS [Group Function] ");
            sbGROUP.Append("UNION ALL ");
            sbGROUP.Append("SELECT '1' AS [ID], 'Development' AS [Group Function] ");
            sbGROUP.Append("UNION ALL ");
            sbGROUP.Append("SELECT '2' AS [ID], 'MPS : Master Production Schedule' AS [Group Function] ");
            sbGROUP.Append("UNION ALL ");
            sbGROUP.Append("SELECT '3' AS [ID], 'MRP : Material Resource Planning' AS [Group Function] ");
            sbGROUP.Append("UNION ALL ");
            sbGROUP.Append("SELECT '4' AS [ID], 'Shipment' AS [Group Function] ");
            sbGROUP.Append("UNION ALL ");
            sbGROUP.Append("SELECT '5' AS [ID], 'EXIMs' AS [Group Function] ");
            sbGROUP.Append("UNION ALL ");
            sbGROUP.Append("SELECT '6' AS [ID], 'Administrator' AS [Group Function] ");

            arlFunc = new ArrayList();
            arlFunc.Add("");
            //Administrator
            arlFunc.Add("M01");
            arlFunc.Add("F02");
            arlFunc.Add("F01");
            arlFunc.Add("M05");
            arlFunc.Add("F03");
            arlFunc.Add("F04");
            arlFunc.Add("F05");

            //Master
            arlFunc.Add("M02");
            arlFunc.Add("M03");
            arlFunc.Add("M04");
            arlFunc.Add("M14");
            arlFunc.Add("M18");
            arlFunc.Add("M06");
            arlFunc.Add("M07");
            arlFunc.Add("M08");
            arlFunc.Add("M09");
            arlFunc.Add("M10");
            arlFunc.Add("M11");
            arlFunc.Add("M12");
            arlFunc.Add("M13");
            arlFunc.Add("M15");
            arlFunc.Add("M16");
            arlFunc.Add("M17");

            //Develop
            arlFunc.Add("DEV01");
            arlFunc.Add("DEV02");
            arlFunc.Add("DEV03");
            arlFunc.Add("DEV04");
            arlFunc.Add("DEV05");

            //MPS
            arlFunc.Add("MPS01");
            arlFunc.Add("MPS02");

            NewData();
            LoadData();
        }

        private void LoadData()
        {
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT Code AS [Company Code], EngName AS[Name (En)], THName AS[Name (Th)], OIDCOMPANY AS ID ");
            sbSQL.Append("FROM Company ");
            sbSQL.Append("ORDER BY OIDCOMPANY ");
            new ObjDE.setGridLookUpEdit(glueCompany, sbSQL, "Company Code", "ID").getData();
            glueCompany.Properties.View.PopulateColumns(glueCompany.Properties.DataSource);
            glueCompany.Properties.View.Columns["ID"].Visible = false;

            glueCompany.EditValue = this.Company;
            glueCompany.Enabled = false;
            glueCompany.BackColor = Color.White;
            glueCompany.ForeColor = Color.Navy;

            glueBranch.Properties.DataSource = null;
            sbSQL.Clear();
            sbSQL.Append("SELECT Code, Name AS [Branch Name], OIDBranch AS ID ");
            sbSQL.Append("FROM  Branchs ");
            sbSQL.Append("WHERE (OIDCOMPANY = '" + this.Company + "') ");
            sbSQL.Append("ORDER BY Code ");
            new ObjDE.setGridLookUpEdit(glueBranch, sbSQL, "Branch Name", "ID").getData();
            glueBranch.Properties.View.PopulateColumns(glueBranch.Properties.DataSource);
            glueBranch.Properties.View.Columns["ID"].Visible = false;

            glueDepartment.Properties.DataSource = null;

            LoadFunction();
            
            sbSQL.Clear();
            sbSQL.Append("SELECT U.OIDUSER AS [User ID], U.UserName, U.FullName, U.Password, U.PasswordFirstLogon AS [First Logon], U.OIDDEPT AS [Department ID], D.Code AS [Department Code], D.Name AS [Department Name],  ");
            sbSQL.Append("       U.OIDCompany AS[Company ID], C.Code AS[Company Code], C.EngName AS[Company Name], U.OIDBranch AS[Branch ID], B.Code AS[Branch Code], B.Name AS[Branch Name], U.Status AS [StatusID], CASE WHEN U.Status = 0 THEN 'Non Active' ELSE CASE WHEN U.Status = 1 THEN 'Active' ELSE '' END END AS Status ");
            sbSQL.Append("FROM   Users AS U LEFT OUTER JOIN ");
            sbSQL.Append("       Departments AS D ON U.OIDDEPT = D.OIDDEPT LEFT OUTER JOIN ");
            sbSQL.Append("       Company AS C ON U.OIDCompany = C.OIDCOMPANY LEFT OUTER JOIN ");
            sbSQL.Append("       Branchs AS B ON U.OIDBranch = B.OIDBranch ");
            sbSQL.Append("WHERE  (U.OIDCompany = '" + this.Company + "') AND (U.UserName <> 'admin') ");
            sbSQL.Append("ORDER BY[Company ID], [Department ID], B.OIDBranch, [User ID] ");
            new ObjDE.setGridControl(gcUser, gvUser, sbSQL).getData(false, false, false, true);
            gvUser.Columns["User ID"].Visible = false;
            gvUser.Columns["Department ID"].Visible = false;
            gvUser.Columns["Company ID"].Visible = false;
            gvUser.Columns["Branch ID"].Visible = false;
            gvUser.Columns["StatusID"].Visible = false;

            txeUserName.Focus();
        }

        private void LoadFunction(string UserID = "")
        {
            StringBuilder sbSQL = new StringBuilder();
            if (lblStatus.Text == "* Edit User")
            {
                sbSQL.Append("SELECT GF.[Group Function] AS [Group], FL.FunctionNo, FL.FunctionName, MAX(FL.Version) AS Version, FL.GroupFunction, ISNULL(FA.AllowDenyStatus, 0) AS [Permission], ISNULL(FA.ReadWriteStatus, 0) AS [ReadWrite]  ");
                sbSQL.Append("FROM   FunctionList AS FL LEFT OUTER JOIN ");
                sbSQL.Append("       (" + sbGROUP + ") AS GF ON FL.GroupFunction = GF.ID LEFT OUTER JOIN ");
                sbSQL.Append("       FunctionAccess AS FA ON FL.FunctionNo = FA.FunctionNo AND FA.OIDUser = '" + UserID + "' ");
                sbSQL.Append("GROUP BY FL.GroupFunction, GF.ID, GF.[Group Function], FL.FunctionNo, FL.FunctionName, FA.AllowDenyStatus, FA.ReadWriteStatus ");
                sbSQL.Append("ORDER BY GF.ID, FL.FunctionNo ");
            }
            else
            {
                sbSQL.Append("SELECT GF.[Group Function] AS [Group], FL.FunctionNo, FL.FunctionName, MAX(FL.Version) AS Version, FL.GroupFunction, 0 AS [Permission], 0 AS [ReadWrite]  ");
                sbSQL.Append("FROM   FunctionList AS FL LEFT OUTER JOIN ");
                sbSQL.Append("       (" + sbGROUP + ") AS GF ON FL.GroupFunction = GF.ID ");
                sbSQL.Append("GROUP BY FL.GroupFunction, GF.ID, GF.[Group Function], FL.FunctionNo, FL.FunctionName ");
                sbSQL.Append("ORDER BY GF.ID, FL.FunctionNo ");
            }
            new ObjDE.setGridControl(gcFunction, gvFunction, sbSQL).getData(false, false, false, true);

            gvFunction.Columns[0].Group();
            gvFunction.Columns["GroupFunction"].Visible = false;
            gvFunction.Columns["Version"].Visible = false;

            dtFuncName = this.DBC.DBQuery(sbSQL).getDataTable();
        }

        private void NewData()
        {
            lblStatus.Text = "* Add User";
            lblStatus.ForeColor = Color.Green;

            txeID.Text = this.DBC.DBQuery("SELECT CASE WHEN ISNULL(MAX(OIDUSER), '') = '' THEN 1 ELSE MAX(OIDUSER) + 1 END AS NewNo FROM Users").getString();
            txeUserName.Text = "";
            txeFullName.Text = "";
            txtPassword.Text = "";
            glueDepartment.EditValue = "";
            glueCompany.EditValue = "";
            glueBranch.EditValue = "";

            rgStatus.EditValue = 1;
        }

        private void bbiNew_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            LoadData();
            NewData();
        }

        private bool chkDuplicate()
        {
            bool chkDup = true;
            if (txeUserName.Text != "")
            {
                txeUserName.Text = txeUserName.Text.Trim();
                if (txeUserName.Text.Trim() != "" && lblStatus.Text == "* Add User")
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT TOP(1) UserName FROM Users WHERE (UserName = N'" + txeUserName.Text.Trim().Trim().Replace("'", "''") + "') ");
                    if (this.DBC.DBQuery(sbSQL).getString() != "")
                    {
                        txeUserName.Text = "";
                        txeUserName.Focus();
                        FUNC.msgWarning("Duplicate user name. !! Please Change.");
                        chkDup = false;
                    }
                }
                else if (txeUserName.Text.Trim() != "" && lblStatus.Text == "* Edit User")
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT TOP(1) OIDUSER ");
                    sbSQL.Append("FROM Users ");
                    sbSQL.Append("WHERE (UserName = N'" + txeUserName.Text.Trim().Trim().Replace("'", "''") + "') ");
                    string strCHK = this.DBC.DBQuery(sbSQL).getString();
                    if (strCHK != "" && strCHK != txeID.Text.Trim())
                    {
                        txeUserName.Text = "";
                        txeUserName.Focus();
                        FUNC.msgWarning("Duplicate user name. !! Please Change.");
                        chkDup = false;
                    }
                }
            }
            return chkDup;
        }

        private void bbiSave_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            gvFunction.CloseEditor();
            gvFunction.UpdateCurrentRow();

            if (txeUserName.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input user name.");
                txeUserName.Focus();
            }
            else if (txeFullName.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input full name.");
                txeFullName.Focus();
            }
            else if (txtPassword.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input password.");
                txtPassword.Focus();
            }
            else if (glueCompany.Text.Trim() == "")
            {
                FUNC.msgWarning("Please select company.");
                glueCompany.Focus();
            }
            else if (glueBranch.Text.Trim() == "")
            {
                FUNC.msgWarning("Please select branch.");
                glueBranch.Focus();
            }
            else if (glueDepartment.Text.Trim() == "")
            {
                FUNC.msgWarning("Please select department.");
                glueDepartment.Focus();
            }
            else
            {
                if (FUNC.msgQuiz("Confirm save data ?") == true)
                {
                    StringBuilder sbSQL = new StringBuilder();

                    bool chkGMP = chkDuplicate();
                    if (chkGMP == true)
                    {
                        string Status = "NULL";
                        if (rgStatus.SelectedIndex != -1)
                        {
                            Status = rgStatus.Properties.Items[rgStatus.SelectedIndex].Value.ToString();
                        }

                        //*** User ****
                        if (lblStatus.Text == "* Add User")
                        {
                            sbSQL.Append("  INSERT INTO Users(UserName, FullName, Password, PasswordFirstLogon, OIDCompany, OIDBranch, OIDDEPT, Status) ");
                            sbSQL.Append("  VALUES(N'" + txeUserName.Text.Trim().Replace("'", "''") + "', N'" + txeFullName.Text.Trim().Replace("'", "''") + "', N'" + txtPassword.Text.Trim().Replace("'", "''") + "', N'', '" + glueCompany.EditValue.ToString() + "', '" + glueBranch.EditValue.ToString() + "', '" + glueDepartment.EditValue.ToString() + "', " + Status + ")   ");
                        }
                        else if (lblStatus.Text == "* Edit User")
                        {
                            sbSQL.Append("  UPDATE Users SET ");
                            sbSQL.Append("      UserName=N'" + txeUserName.Text.Trim().Replace("'", "''") + "', ");
                            sbSQL.Append("      FullName=N'" + txeFullName.Text.Trim().Replace("'", "''") + "', ");
                            sbSQL.Append("      Password=N'" + txtPassword.Text.Trim().Replace("'", "''") + "', ");
                            sbSQL.Append("      OIDCompany = '" + glueCompany.EditValue.ToString() + "', ");
                            sbSQL.Append("      OIDBranch = '" + glueBranch.EditValue.ToString() + "', ");
                            sbSQL.Append("      OIDDEPT = '" + glueDepartment.EditValue.ToString() + "', ");
                            sbSQL.Append("      Status=" + Status + " ");
                            sbSQL.Append("  WHERE(OIDUSER = '" + txeID.Text.Trim() + "')  ");
                        }

                        //MessageBox.Show(sbSQL.ToString());
                        if (sbSQL.Length > 0)
                        {
                            try
                            {
                                bool chkSAVE = this.DBC.DBQuery(sbSQL).runSQL();
                                if (chkSAVE == true)
                                {
                                    sbSQL.Clear();
                                    //*** FunctionAccess ****
                                    string UID = "";
                                    if (lblStatus.Text == "* Add User")
                                        UID = this.DBC.DBQuery("SELECT OIDUSER FROM Users WHERE (UserName=N'" + txeUserName.Text.Trim().Replace("'", "''") + "') ").getString();
                                    else if (lblStatus.Text == "* Edit User")
                                        UID = txeID.Text.Trim();

                                    DataTable dtFUNC = (DataTable)gcFunction.DataSource;
                                    if (dtFUNC.Rows.Count > 0)
                                    {
                                        foreach (DataRow row in dtFUNC.Rows)
                                        {
                                            int Permission = Convert.ToInt32(row["Permission"].ToString());
                                            if (Permission == 0) //Deny
                                            {
                                                sbSQL.Append("DELETE FROM FunctionAccess WHERE (OIDUser = '" + UID + "') AND (FunctionNo = '" + row["FunctionNo"].ToString() + "')  ");
                                            }
                                            else if(Permission == 1) //Allow
                                            {
                                                sbSQL.Append("IF NOT EXISTS(SELECT OIDAccess FROM FunctionAccess WHERE (OIDUser = '" + UID + "') AND (FunctionNo = '" + row["FunctionNo"].ToString() + "')) ");
                                                sbSQL.Append(" BEGIN ");
                                                sbSQL.Append("  INSERT INTO FunctionAccess(OIDUser, FunctionNo, ReadWriteStatus, AllowDenyStatus) ");
                                                sbSQL.Append("   VALUES('" + UID + "', '" + row["FunctionNo"].ToString() + "', '" + row["ReadWrite"].ToString() + "', '" + row["Permission"].ToString() + "') ");
                                                sbSQL.Append(" END ");
                                                sbSQL.Append(" BEGIN ");
                                                sbSQL.Append("  UPDATE FunctionAccess SET ");
                                                sbSQL.Append("    ReadWriteStatus='" + row["ReadWrite"].ToString() + "' ");
                                                sbSQL.Append("  WHERE (OIDUser = '" + UID + "') AND (FunctionNo = '" + row["FunctionNo"].ToString() + "') ");
                                                sbSQL.Append(" END   ");
                                            }
                                        }
                                    }


                                    if (sbSQL.Length > 0)
                                    {
                                        chkSAVE = this.DBC.DBQuery(sbSQL).runSQL();
                                        if (chkSAVE == true)
                                        {
                                            FUNC.msgInfo("Save complete.");
                                            bbiNew.PerformClick();
                                        }
                                    }
                                    else //No function
                                    {
                                        FUNC.msgInfo("Save complete.");
                                        bbiNew.PerformClick();
                                    }
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
            string pathFile = new ObjSet.Folder(@"C:\MDS\Export\").GetPath() + "UserList_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
            gvUser.ExportToXlsx(pathFile);
            System.Diagnostics.Process.Start(pathFile);
        }


        private void bbiPrintPreview_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            gcUser.ShowPrintPreview();
        }

        private void bbiPrint_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            gcUser.Print();
        }


        int CalcGroupCount(int groupRowHandle)
        {
            int groupIndex = Math.Abs(groupRowHandle) - 1;
            return gvFunction.DataController.GroupInfo[groupIndex].ChildControllerRowCount;
        }

        private int FindIndexImgFunc(string FuncCode)
        {
            int retImg = 0;
            for (int i = 0; i < arlFunc.Count; i++)
            {
                if (arlFunc[i].ToString() == FuncCode)
                {
                    retImg = i;
                    break;
                }
            }
            return retImg;
        }

        private void SetImage(RowCellCustomDrawEventArgs e, Image image)
        {
            GridCellInfo gci = e.Cell as GridCellInfo;
            TextEditViewInfo info = gci.ViewInfo as TextEditViewInfo;
            info.ContextImage = image;
            info.CalcViewInfo(e.Graphics);
        }

        private void gvFunction_EndGrouping(object sender, EventArgs e)
        {
            gvFunction.ExpandAllGroups();
            gvFunction.Columns["Group"].Visible = false;
        }

        private void gvFunction_CustomColumnDisplayText(object sender, CustomColumnDisplayTextEventArgs e)
        {
            if (e.IsForGroupRow)
            {
                e.DisplayText = string.Format("{0} ({1})", e.DisplayText, CalcGroupCount(e.GroupRowHandle));
            }
        }

        private void gvFunction_CustomDrawCell(object sender, RowCellCustomDrawEventArgs e)
        {
            if (e.RowHandle < 0)
            {
                return;
            }

            if (e.Column.FieldName == "FunctionName")
            {
                if (e.CellValue == null)
                {
                    return;
                }

                string FuncCode = "";
                foreach (DataRow drFunc in dtFuncName.Rows)
                {
                    if (drFunc["FunctionName"].ToString() == e.CellValue.ToString())
                    {
                        FuncCode = drFunc["FunctionNo"].ToString();
                        break;
                    }
                }

                SetImage(e, icFunction.Images[FindIndexImgFunc(FuncCode)]);
            }
        }

        private void gvUser_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator) e.Info.DisplayText = (e.RowHandle + 1).ToString();
            gvUser.IndicatorWidth = 40;
        }

        private void glueCompany_EditValueChanged(object sender, EventArgs e)
        {
            if (glueCompany.Text.Trim() != "")
            {
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT Code, Name AS [Branch Name], OIDBranch AS ID ");
                sbSQL.Append("FROM  Branchs ");
                sbSQL.Append("WHERE (OIDCOMPANY = '" + glueCompany.EditValue.ToString() + "') ");
                sbSQL.Append("ORDER BY Code ");
                new ObjDE.setGridLookUpEdit(glueBranch, sbSQL, "Branch Name", "ID").getData();
                glueBranch.Properties.View.PopulateColumns(glueBranch.Properties.DataSource);
                glueBranch.Properties.View.Columns["ID"].Visible = false;

                glueDepartment.Properties.DataSource = null;
                glueBranch.Focus();
            }
        }

        private void glueBranch_EditValueChanged(object sender, EventArgs e)
        {
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT DP.Code, DP.Name AS Department, DT.Name AS Type, DP.OIDDEPT AS ID ");
            sbSQL.Append("FROM   Departments AS DP INNER JOIN ");
            sbSQL.Append("       DepartmentType AS DT ON DP.DepartmentType = DT.Code ");
            sbSQL.Append("WHERE (DP.OIDCOMPANY = '" + this.Company + "') AND(DP.OIDBRANCH = '" + glueBranch.EditValue.ToString() + "') ");
            sbSQL.Append("ORDER BY DT.Code, DP.Code ");
            new ObjDE.setGridLookUpEdit(glueDepartment, sbSQL, "Department", "ID").getData();
            glueDepartment.Properties.View.PopulateColumns(glueDepartment.Properties.DataSource);
            glueDepartment.Properties.View.Columns["ID"].Visible = false;
            glueDepartment.Focus();
        }

        private void txeUserName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeFullName.Focus();
            }
        }

        private void txeUserName_Leave(object sender, EventArgs e)
        {
            if (txeUserName.Text.Trim() != "")
            {
                txeUserName.Text = txeUserName.Text.Trim();
                bool chkDup = chkDuplicate();
                if (chkDup == true)
                {
                    txeFullName.Focus();
                }
            }
        }

        private void txeFullName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtPassword.Focus();
            }
        }

        private void txtPassword_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                glueCompany.Focus();
            }
        }

        private void glueDepartment_EditValueChanged(object sender, EventArgs e)
        {
            rgStatus.Focus();
        }

        private void gvUser_DoubleClick(object sender, EventArgs e)
        {
            GridView view = (GridView)sender;
            Point pt = view.GridControl.PointToClient(Control.MousePosition);
            DevExpress.XtraGrid.Views.Grid.ViewInfo.GridHitInfo info = view.CalcHitInfo(pt);
            if (info.InRow || info.InRowCell)
            {
                DataTable dtCP = (DataTable)gcUser.DataSource;
                if (dtCP.Rows.Count > 0)
                {
                    lblStatus.Text = "* Edit User";
                    lblStatus.ForeColor = Color.Red;

                    DataRow drCP = dtCP.Rows[info.RowHandle];
                    txeID.Text = drCP["User ID"].ToString();
                    
                    txeUserName.Text = drCP["UserName"].ToString();
                    txeFullName.Text = drCP["FullName"].ToString();
                    txtPassword.Text = drCP["Password"].ToString();
                    glueCompany.EditValue = drCP["Company ID"].ToString();
                    glueBranch.EditValue = drCP["Branch ID"].ToString();
                    glueDepartment.EditValue = drCP["Department ID"].ToString();
                    rgStatus.EditValue = Convert.ToInt32(drCP["StatusID"].ToString());

                    LoadFunction(txeID.Text.Trim());
                }
            }
        }
    }
}