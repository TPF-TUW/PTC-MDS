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
using DevExpress.XtraTreeList.Nodes;
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
using TheepClass;

namespace MDS.Function
{
    public partial class F04 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private Functionality.Function FUNC = new Functionality.Function();
        bool loadDrives = false;
        StringBuilder sbGROUP = new StringBuilder();

        private RepositoryItemTextEdit edit;
        ArrayList arlFunc;

        DataTable dtFuncName;
        private string pathDrive = @"\\172.16.0.190\MDS_Project\MDS\EXE\";
        public LogIn UserLogin { get; set; }
        public string ConnectionString { get; set; }
        string CONNECT_STRING = "";
        DatabaseConnect DBC;


        public F04()
        {
            InitializeComponent();
            UserLookAndFeel.Default.StyleChanged += MyStyleChanged;
            edit = new RepositoryItemTextEdit();
        }

        private void MyStyleChanged(object sender, EventArgs e)
        {
            UserLookAndFeel userLookAndFeel = (UserLookAndFeel)sender;
            cUtility.SaveRegistry(@"Software\MDS", "SkinName", userLookAndFeel.SkinName);
            cUtility.SaveRegistry(@"Software\MDS", "SkinPalette", userLookAndFeel.ActiveSvgPaletteName);
        }

        private void XtraForm1_Load(object sender, EventArgs e)
        {//***** SET CONNECT DB ********
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

            NewData();
            LoadData();

            
        }

        private void LoadData()
        {
            //_Painter = new GridSkinElementsPainter(gvFolder);
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT GF.[Group Function] AS [Group], FL.FunctionNo AS [Function No.], FL.FunctionName, MAX(FL.Version) AS Version, FL.GroupFunction ");
            sbSQL.Append("FROM   FunctionList AS FL LEFT OUTER JOIN ");
            sbSQL.Append("       (" + sbGROUP + ") AS GF ON FL.GroupFunction = GF.ID ");
            sbSQL.Append("GROUP BY FL.GroupFunction, GF.[Group Function], FL.FunctionNo, FL.FunctionName ");
            sbSQL.Append("ORDER BY FL.GroupFunction, FL.FunctionName ");
            new ObjDE.setGridControl(gcFolder, gvFolder, sbSQL).getData(false, false, false, true);

            gvFolder.Columns[0].Group();
            gvFolder.Columns["GroupFunction"].Visible = false;

            dtFuncName = this.DBC.DBQuery(sbSQL).getDataTable();

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
            arlFunc.Add("M06");
            arlFunc.Add("M07");
            arlFunc.Add("M08");
            arlFunc.Add("M09");
            arlFunc.Add("M10");
            arlFunc.Add("M11");
            arlFunc.Add("M12");
            arlFunc.Add("M13");
            arlFunc.Add("M14");
            arlFunc.Add("M15");

            //Develop
            arlFunc.Add("DEV01");
            arlFunc.Add("DEV02");
            arlFunc.Add("DEV03");


            gvFolder.Columns["FunctionName"].ColumnEdit = edit;
            new ObjDE.setGridLookUpEdit(glueGroup, sbGROUP, "Group Function", "ID").getData();

            LoadAllFunction();
        }

        private void NewData()
        {
            //**** LOAD DIRECTORY *********
            loadDrives = false;
            tlDirectory.DataSource = new object();
            tlDirectory.StateImageList = ImageCollection1;
            tlDirectory.OptionsBehavior.Editable = false;
            //tlDirectory.OptionsBehavior.AllowRecursiveNodeChecking = True
            tlDirectory.OptionsView.EnableAppearanceEvenRow = true;
            //tlDirectory.OptionsView.ShowCheckBoxes = true;
            tlDirectory.Nodes.FirstNode.Expanded = true;
            //**** END LOAD DIRECTORY *********

            lblStatus.Text = "* Add Function";
            lblStatus.ForeColor = Color.Green;

            txeID.Text = this.DBC.DBQuery("SELECT CASE WHEN ISNULL(MAX(OIDFunction), '') = '' THEN 1 ELSE MAX(OIDFunction) + 1 END AS NewNo FROM FunctionList").getString();
            txeCode.Text = "";
            txeName.Text = "";
            txeVersion.Text = "1.0.0.0";
            rgStatus.EditValue = 1;
            //txeName.ReadOnly = false;
            glueCREATE.EditValue = UserLogin.OIDUser;
            txeDATE.Text = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");

            //////txeID.Focus();
        }

        private void LoadAllFunction()
        {
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT FL.GroupFunction AS [GroupID], GF.[Group Function] AS [Group], FL.FunctionNo AS [Function No.], FL.OIDFunction AS ID, FL.FunctionName, FL.Version, FL.Status AS StatusID, CASE WHEN FL.Status = 0 THEN 'Non Active' ELSE CASE WHEN FL.Status = 1 THEN 'Active' ELSE '' END END AS Status, FL.CreateBy AS [Created By], CONVERT(VARCHAR(10), FL.CreateDate, 103) AS [Created Date], FL.PathFile ");
            sbSQL.Append("FROM   FunctionList AS FL LEFT OUTER JOIN ");
            sbSQL.Append("       (" + sbGROUP + ") AS GF ON FL.GroupFunction = GF.ID ");
            sbSQL.Append("WHERE (FL.FunctionNo <> N'')  ");
            if (glueGroup.Text.Trim() != "")
            {
                sbSQL.Append("AND (FL.GroupFunction = '" + glueGroup.EditValue.ToString() + "') ");
            }
            //if (txeCode.Text.Trim() != "")
            //{
            //    sbSQL.Append("AND (FL.FunctionNo = N'" + txeCode.Text.Trim() + "') ");
            //}
            sbSQL.Append("ORDER BY [Group], FL.FunctionNo, FL.Version ");
            new ObjDE.setGridControl(gcFunction, gvFunction, sbSQL).getData(false, false, false, true);
            gvFunction.Columns[1].Group();
            gvFunction.Columns[2].Group();
            gvFunction.Columns["ID"].Visible = false;
            gvFunction.Columns["GroupID"].Visible = false;
            gvFunction.Columns["StatusID"].Visible = false;
            gvFunction.Columns["PathFile"].Visible = false;

            gvFunction.Columns["Created By"].Visible = false;
            gvFunction.Columns["Created Date"].Visible = false;

            gvFunction.ExpandAllGroups();
        }

        private void bbiNew_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            LoadData();
            NewData();
        }

        private void bbiSave_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            txeID.Focus();
            if (glueGroup.Text.Trim() == "")
            {
                FUNC.msgWarning("Please select function group.");
                glueGroup.Focus();
            }
            else if (txeCode.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input function code.");
                txeCode.Focus();
            }
            else if (txeName.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input function name.");
                txeName.Focus();
            }
            else if (txeVersion.ErrorText != "")
            {
                FUNC.msgWarning("Please input version.");
                txeVersion.Focus();
            }
            else if (txeBrowse.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input function file.");
                txeBrowse.Focus();
            }
            else
            {   
                if (FUNC.msgQuiz("Confirm save data ?") == true)
                {
                    StringBuilder sbSQL = new StringBuilder();
                    string strCREATE = UserLogin.OIDUser.ToString() != "" ? UserLogin.OIDUser.ToString() : "0";

                    string Status = "NULL";
                    if (rgStatus.SelectedIndex != -1)
                    {
                        Status = rgStatus.Properties.Items[rgStatus.SelectedIndex].Value.ToString();
                    }


                    if (lblStatus.Text == "* Add Function")
                    {
                        string newFileName = new ObjSet.Folder(pathDrive + txeCode.Text.ToUpper().Trim()).GetPath() + txeCode.Text.ToUpper().Trim() + "_" + txeVersion.Text.Trim();
                        //CopyFile
                        if (txeBrowse.Text.Trim() != "")
                        {
                            System.IO.FileInfo fi = new System.IO.FileInfo(txeBrowse.Text);
                            string extn = fi.Extension;
                            newFileName += extn;
                            System.IO.File.Copy(txeBrowse.Text.Trim(), newFileName);
                        }

                        sbSQL.Append("  INSERT INTO FunctionList(FunctionNo, FunctionName, Version, GroupFunction, Status, PathFile, CreateBy, CreateDate) ");
                        sbSQL.Append("  VALUES(N'" + txeCode.Text.Trim().Replace("'", "''") + "', N'" + txeName.Text.Trim().Replace("'", "''") + "', N'" + txeVersion.Text.Trim() + "', '" + glueGroup.EditValue.ToString() + "', " + Status + ", '" + newFileName + "', '" + strCREATE + "', GETDATE()) ");
                    }
                    else if (lblStatus.Text == "* Edit Function")
                    {
                        StringBuilder sbCHK_FILE = new StringBuilder();
                        sbCHK_FILE.Append("SELECT PathFile FROM FunctionList WHERE (OIDFunction = '" + txeID.Text.Trim() + "') ");
                        string chkFILE = this.DBC.DBQuery(sbCHK_FILE).getString();
                        string newFileName = "";
                        if (chkFILE != txeBrowse.Text.Trim())
                        {
                            newFileName = new ObjSet.Folder(pathDrive + txeCode.Text.ToUpper().Trim()).GetPath() + txeCode.Text.ToUpper().Trim() + "_" + txeVersion.Text.Trim();
                            //CopyFile
                            System.IO.FileInfo fi = new System.IO.FileInfo(txeBrowse.Text);
                            string extn = fi.Extension;
                            newFileName += extn;
                            System.IO.File.Copy(txeBrowse.Text.Trim(), newFileName);
                        }

                        sbSQL.Append("  UPDATE FunctionList SET ");
                        sbSQL.Append("      FunctionNo=N'" + txeCode.Text.ToUpper().Trim().Replace("'", "''") + "', ");
                        sbSQL.Append("      FunctionName=N'" + txeName.Text.Trim().Replace("'", "''") + "', ");
                        sbSQL.Append("      Version=N'" + txeVersion.Text.Trim() + "', ");
                        sbSQL.Append("      GroupFunction='" + glueGroup.EditValue.ToString() + "', ");
                        if (newFileName != "")
                        {
                            sbSQL.Append("  PathFile='" + newFileName + "', ");
                        }
                        sbSQL.Append("      Status=" + Status + " ");
                        sbSQL.Append("  WHERE (OIDFunction = '" + txeID.Text.Trim() + "') ");
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

        private void bbiExcel_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            string pathFile = new ObjSet.Folder(@"C:\MDS\Export\").GetPath() + "FunctionList_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
            gvFunction.ExportToXlsx(pathFile);
            System.Diagnostics.Process.Start(pathFile);
        }

        private void bbiPrintPreview_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            gcFunction.ShowPrintPreview();
        }

        private void bbiPrint_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            gcFunction.Print();
        }

        private bool IsFile(DirectoryInfo info)
        {
            return (info.Attributes & FileAttributes.Directory) == 0;
        }
       
        private void tlDirectory_VirtualTreeGetChildNodes(object sender, DevExpress.XtraTreeList.VirtualTreeGetChildNodesInfo e)
        {
            if (!loadDrives)
            { // create drives
                //string[] root = Directory.GetLogicalDrives();
                string[] root = { pathDrive };
                e.Children = root;
                loadDrives = true;
            }
            else
            {
                try
                {
                    string path = (string)e.Node;
                    if (Directory.Exists(path))
                    {
                        string[] dirs = Directory.GetDirectories(path);
                        string[] files = Directory.GetFiles(path, "*.exe");
                        string[] arr = new string[dirs.Length + files.Length];
                        dirs.CopyTo(arr, 0);
                        files.CopyTo(arr, dirs.Length);
                        e.Children = arr;
                    }
                    else e.Children = new object[] { };
                }
                catch { e.Children = new object[] { }; }
            }
            
        }

        private void tlDirectory_VirtualTreeGetCellValue(object sender, DevExpress.XtraTreeList.VirtualTreeGetCellValueInfo e)
        {
            DirectoryInfo di = new DirectoryInfo((string)e.Node);
            if (e.Column == colName)
                e.CellData = di.Name;
            if (e.Column == colType)
            {
                if (!IsFile(di))
                    e.CellData = "Folder";
                else
                    e.CellData = "File";
            }
            if (e.Column == colSize)
            {
                if (IsFile(di))
                {
                    e.CellData = new FileInfo((string)e.Node).Length;
                }
                else e.CellData = null;
            }
            if (e.Column == colDate)
            {
                if (IsFile(di))
                {
                    e.CellData = new FileInfo((string)e.Node).CreationTime;
                }
                else e.CellData = null;
            }

        }

        private void tlDirectory_GetStateImage(object sender, DevExpress.XtraTreeList.GetStateImageEventArgs e)
        {
            if (e.Node.GetDisplayText("Type") == "Folder")
            {
                e.NodeImageIndex = e.Node.Expanded ? 1 : 0;
            } else if (e.Node.GetDisplayText("Type") == "File")
            {
                e.NodeImageIndex = 2;
            } else
            {
                e.NodeImageIndex = 3;
            }
        }

        private void gvFunction_CustomColumnGroup(object sender, DevExpress.XtraGrid.Views.Base.CustomColumnSortEventArgs e)
        {

        }

        private void gvFolder_CustomColumnDisplayText(object sender, DevExpress.XtraGrid.Views.Base.CustomColumnDisplayTextEventArgs e)
        {
            if (e.IsForGroupRow)
            {
                e.DisplayText = string.Format("{0} ({1})", e.DisplayText, CalcGroupCount(e.GroupRowHandle));
            }
        }

        int CalcGroupCount(int groupRowHandle)
        {
            int groupIndex = Math.Abs(groupRowHandle) - 1;
            return gvFolder.DataController.GroupInfo[groupIndex].ChildControllerRowCount;
        }

        private void gvFolder_EndGrouping(object sender, EventArgs e)
        {
            gvFolder.ExpandAllGroups();
            gvFolder.Columns["Group"].Visible = false;
        }

        private void gvFolder_CustomDrawColumnHeader(object sender, ColumnHeaderCustomDrawEventArgs e)
        {
          
        }

        private void gvFolder_CustomDrawGroupRow(object sender, RowObjectCustomDrawEventArgs e)
        {
           
        }

        private void gvFolder_CustomDrawCell(object sender, RowCellCustomDrawEventArgs e)
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
                        FuncCode = drFunc["Function No."].ToString();
                        break;
                    }
                }

                SetImage(e, icFunction.Images[FindIndexImgFunc(FuncCode)]);
            }
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
                        FuncCode = drFunc["Function No."].ToString();
                        break;
                    }
                }

                SetImage(e, icFunction.Images[FindIndexImgFunc(FuncCode)]);
            }
        }

        private void glueGroup_EditValueChanged(object sender, EventArgs e)
        {
            LoadAllFunction();
            txeCode.Focus();
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
                txeCode.Text = txeCode.Text.ToUpper().Trim();
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT FunctionName, MAX(Version) AS Version ");
                sbSQL.Append("FROM FunctionList ");
                sbSQL.Append("WHERE (FunctionNo = N'" + txeCode.Text.Trim() + "') ");
                sbSQL.Append("GROUP BY FunctionName ");
                string[] arrFunc = this.DBC.DBQuery(sbSQL).getMultipleValue();
                if (arrFunc.Length > 0)
                {
                    txeName.Text = arrFunc[0];
                    string strVer = arrFunc[1];
                    if (lblStatus.Text == "* Add Function")
                    {
                        if (strVer == "")
                        {
                            strVer = "1.0.0.0";
                        }
                        else
                        {
                            string[] spV = strVer.Split('.');
                            if (spV.Length > 0)
                            {
                                strVer = spV[0] + "." + spV[1] + "." + spV[2] + "." + (Convert.ToInt32(spV[3]) + 1).ToString();
                            }
                        }
                    }
                    txeVersion.Text = strVer;
                    //txeName.ReadOnly = true;
                    txeVersion.Focus();
                }
                else
                {
                    if (txeName.Text.Trim() != "")
                    {
                        sbSQL.Clear();
                        sbSQL.Append("SELECT FunctionNo FROM FunctionList WHERE (FunctionName = N'" + txeName.Text.Trim() + "') ");
                        string strNo = this.DBC.DBQuery(sbSQL).getString();
                        if (strNo != "" && strNo.ToUpper().Trim() != txeCode.Text.Trim())
                        {
                            txeName.Text = "";
                            //txeName.ReadOnly = false;
                            txeVersion.Text = "1.0.0.0";
                            txeName.Focus();
                        }
                    }
                    else
                    {
                        txeName.Text = "";
                        txeVersion.Text = "1.0.0.0";
                        //txeName.ReadOnly = false;
                        txeName.Focus();
                    }
                }
            }
            else
            {
                //txeName.ReadOnly = false;
                txeName.Focus();
            }
        }

        private void txeName_KeyDown_1(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeVersion.Focus();
            }
        }

        private void txeName_Leave(object sender, EventArgs e)
        {
            if (txeName.Text.Trim() != "")
            {
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT FunctionNo FROM FunctionList WHERE (FunctionName = N'" + txeName.Text.Trim() + "') ");
                string strNo = this.DBC.DBQuery(sbSQL).getString();
                if (strNo != "")
                {
                    if (txeCode.Text.Trim() != "" && strNo.ToUpper().Trim() != txeCode.Text.ToUpper().Trim())
                    {
                        txeName.Text = "";
                        txeName.Focus();
                        FUNC.msgWarning("Duplicate function name. !! Please Change.");
                    }
                    else
                    {
                        txeVersion.Focus();
                    }
                }
                else
                {
                    txeVersion.Focus();
                }

            }
        }

        private void txeVersion_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                rgStatus.Focus();
            }
        }

        private void txeVersion_Leave(object sender, EventArgs e)
        {
            if (txeVersion.Text.Trim() != "")
            {
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT OIDFunction FROM FunctionList WHERE (FunctionNo = N'" + txeCode.Text.Trim() + "') AND (FunctionName = N'" + txeName.Text.Trim() + "') AND (Version = N'" + txeVersion.Text.Trim() + "') ");
                string strVER = this.DBC.DBQuery(sbSQL).getString();
                if (strVER != "")
                {
                    if (lblStatus.Text == "* Add Function")
                    {
                        txeVersion.Text = "";
                        txeVersion.Focus();
                        FUNC.msgWarning("Duplicate version. !! Please Change.");
                    }
                    else if (lblStatus.Text == "* Edit Function" && strVER != txeID.Text.Trim())
                    {
                        txeVersion.Text = "";
                        txeVersion.Focus();
                        FUNC.msgWarning("Duplicate version. !! Please Change.");
                    }
                    
                }
            }
        }

        private void gvFunction_DoubleClick(object sender, EventArgs e)
        {
            GridView view = (GridView)sender;
            Point pt = view.GridControl.PointToClient(Control.MousePosition);
            DevExpress.XtraGrid.Views.Grid.ViewInfo.GridHitInfo info = view.CalcHitInfo(pt);
            if (info.InRow || info.InRowCell)
            {
                DataTable dtCP = (DataTable)gcFunction.DataSource;
                if (dtCP.Rows.Count > 0)
                {
                    lblStatus.Text = "* Edit Function";
                    lblStatus.ForeColor = Color.Red;
       
                    DataRow drCP = dtCP.Rows[info.RowHandle];
                    txeID.Text = drCP["ID"].ToString();
                    glueGroup.EditValue = drCP["GroupID"].ToString();
                    txeCode.Text = drCP["Function No."].ToString();
                    txeName.Text = drCP["FunctionName"].ToString();
                    txeVersion.Text = drCP["Version"].ToString();
                    txeBrowse.Text = drCP["PathFile"].ToString();
                    rgStatus.EditValue = Convert.ToInt32(drCP["StatusID"].ToString());
                    //txeName.ReadOnly = true;
                    glueCREATE.EditValue = drCP["Created By"].ToString();
                    txeDATE.Text = drCP["Created Date"].ToString();
                }
            }
        }

        private void gvFunction_RowCellClick(object sender, RowCellClickEventArgs e)
        {
            if (gvFunction.IsFilterRow(e.RowHandle)) return;
        }

        private void sbOpenFile_Click(object sender, EventArgs e)
        {
            xtraOpenFileDialog1.Filter = "Application Files|*.exe";
            xtraOpenFileDialog1.FileName = "";
            xtraOpenFileDialog1.Title = "Select Application File";

            if (xtraOpenFileDialog1.ShowDialog() == DialogResult.OK)
            {
                txeBrowse.Text = xtraOpenFileDialog1.FileName;
            }

            txeBrowse.Focus();
        }

        private void txeVersion_InvalidValue(object sender, DevExpress.XtraEditors.Controls.InvalidValueExceptionEventArgs e)
        {
            //e.ExceptionMode = DevExpress.XtraEditors.Controls.ExceptionMode.NoAction;
            //MessageBox.Show("Enter a date within the current month.", "Error");
        }
    }
}