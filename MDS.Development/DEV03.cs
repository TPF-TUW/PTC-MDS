using System;
using System.Text;
using DBConnect;
using System.Windows.Forms;
using System.Globalization;
using System.Data;
using DevExpress.LookAndFeel;
using DevExpress.Utils;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Controls;
using DevExpress.XtraLayout.Utils;
using TheepClass;

namespace MDS.Development
{
    public partial class DEV03 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        cDatabase db;
        CultureInfo clinfo = new CultureInfo("en-US");
        DateTimeFormatInfo dtfinfo;
        DataTable dtFG, dtBomDetail;
        private Functionality.Function FUNC = new Functionality.Function();

        int chkReadWrite = 0;

        public LogIn UserLogin { get; set; }
        public int Company { get; set; }
        public string ConnectionString { get; set; }

        public DEV03()
        {
            InitializeComponent();
        }

        private void NewData()
        {
            try
            {
                ClearSampleRequestDetail();
                ClearBOMDetail();
                ClearBOMDetail_Detail();
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


        }
        private void SaveData()
        {
            if (sleFGProductCode.EditValue == null || sleFGProductCode.EditValue==System.DBNull.Value)
            {
                XtraMessageBox.Show("Please input <FG:Product Code>.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else if (gleUnit.EditValue == null || gleUnit.EditValue==System.DBNull.Value)
            {
                XtraMessageBox.Show("Please input <Unit>.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else if (txtReviseNo.Text == "")
            {
                XtraMessageBox.Show("Please input <Revision No>.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else if (dtpLastDate.EditValue == null ||dtpLastDate.EditValue==System.DBNull.Value)
            {
                XtraMessageBox.Show("Please input <Last Date>.","Error",MessageBoxButtons.OK,MessageBoxIcon.Error);
                return;
            }

            db.ConnectionOpen();
            //------------------------------------------Save bom header------------------------------------------------
            try
            {
                db.BeginTrans();
                string strSQL = "SELECT COUNT(OIDBOM) FROM BOM WHERE OIDSMPL="+txtSmplNo_Header.Tag+
                    " AND OIDSIZE="+txtSize.Tag+" AND OIDCOLOR="+ txtColor.Tag;
                if (db.ExecuteFirstValue(strSQL) == "0")
                {
                    strSQL = "INSERT INTO BOM(OIDSMPL,OIDITEM,OIDSIZE,OIDCOLOR,OIDCUST,OIDSTYLE,OIDCATEGORY,OIDUNIT"+
                        ",BOMNO,REVISIONNO,ISSUEDATE,SEASON,PATTERNZONE,SMPLITEMNO,MODELNAME,COST,STATUS,CREATEDBY"+
                        ",CREATEDDATE,UPDATEDBY,UPDATEDDATE)VALUES(";
                    strSQL += txtSmplNo_Header.Tag;
                    strSQL += "," + sleFGProductCode.EditValue;
                    strSQL += "," + txtSize.Tag;
                    strSQL += "," + txtColor.Tag;
                    strSQL += "," + txtCustomer_Header.Tag;
                    strSQL += "," + txtStyleName.Tag;
                    strSQL += "," + txtCategory.Tag;
                    strSQL += "," + gleUnit.EditValue;
                    strSQL += ",'" + txtBomNo.Text + "'";
                    strSQL += "," + txtReviseNo.Text;
                    strSQL += ",'" + ((DateTime)dtpLastDate.EditValue).ToString("yyyy-MM-dd", dtfinfo)+"'";
                    strSQL += ",'" + txtSeason_Header.Text + "'";
                    strSQL += "," + txtPatternSizeZone.Text;
                    strSQL += ",'" + txtItemNo.Text + "'";
                    strSQL += ",'" + txtModelName.Text + "'";
                    strSQL += "," + txtUnitCost.Text;
                    strSQL += "," + optStatus.SelectedIndex;
                    strSQL += "," + UserLogin.OIDUser;
                    strSQL += ",'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss",dtfinfo)+"'";
                    strSQL += "," + UserLogin.OIDUser;
                    strSQL += ",'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", dtfinfo)+"'";
                    strSQL += ")";
                    db.Execute(strSQL);
                }
                else
                {
                    strSQL = "UPDATE BOM SET ";
                    strSQL += "OIDITEM=" + sleFGProductCode.EditValue;
                    strSQL += ",OIDUNIT=" + gleUnit.EditValue;
                    strSQL += ",BOMNO='" + txtBomNo.Text + "'";
                    strSQL += ",REVISIONNO=" + txtReviseNo.Text;
                    strSQL += ",ISSUEDATE='" + ((DateTime)dtpLastDate.EditValue).ToString("yyyy-MM-dd HH:mm:ss", dtfinfo)+"'";
                    strSQL += ",COST=" + txtUnitCost.Text;
                    strSQL += ",STATUS=" + optStatus.SelectedIndex;
                    strSQL += ",UPDATEDBY=" + UserLogin.OIDUser;
                    strSQL += ",UPDATEDDATE='" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", dtfinfo)+"'";
                    strSQL += " WHERE OIDBOM=" + txtBomNo.Tag;
                    db.Execute(strSQL);
                }
                db.CommitTrans();
                
            }
            catch (Exception ex)
            {
                db.RollbackTrans();
                XtraMessageBox.Show(ex.Message, "Error--Save BOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                db.ConnectionClose();
                return;
            }
            //----------------------------------------Save bom detail----------------------------------------------
            try
            {
                db.BeginTrans();
                string strSQL = "SELECT TOP 1 OIDBOM FROM BOM WHERE OIDSMPL=" + txtSmplNo_Header.Tag +
                        " AND OIDSIZE=" + txtSize.Tag + " AND OIDCOLOR=" + txtColor.Tag;
                var oidBOM = db.ExecuteFirstValue(strSQL);
                if (oidBOM != "")
                {
                    strSQL = "DELETE FROM BOMDETAIL WHERE OIDBOM=" + oidBOM;
                    db.Execute(strSQL);
                    for (int i = 0; i < gridView4.DataRowCount; i++)
                    {
                        strSQL = "INSERT INTO BOMDETAIL(OIDBOM,OIDITEM,OIDUNIT,OIDVEND,OIDDEPT,MATNO,MATDETAIL" +
                            ",CURRENCY,PRICE,CONSUMPTION,COST,PERCENTLOSS,SMPLCOLOR,SMPLLOTNO)VALUES(";
                        strSQL += oidBOM;
                        strSQL += "," + gridView4.GetRowCellValue(i, "OIDITEM");
                        strSQL += "," + gridView4.GetRowCellValue(i, "OIDUNIT");
                        strSQL += "," + gridView4.GetRowCellValue(i, "OIDVEND");
                        strSQL += "," + gridView4.GetRowCellValue(i, "OIDDEPT");
                        strSQL += ",'" + gridView4.GetRowCellValue(i, "VENDOR_CODE")+"'";
                        strSQL += ",'" + gridView4.GetRowCellValue(i, "COMPOSITION")+"'";
                        strSQL += ",'" + gridView4.GetRowCellValue(i, "CURRENCY")+"'";
                        strSQL += "," + gridView4.GetRowCellValue(i, "PRICE");
                        strSQL += "," + gridView4.GetRowCellValue(i, "CONSUMPTION");
                        strSQL += "," + gridView4.GetRowCellValue(i, "COST");
                        strSQL += "," + gridView4.GetRowCellValue(i, "LOSS");
                        strSQL += ",'" + gridView4.GetRowCellValue(i, "COLOR")+"'";
                        strSQL += ",'" + gridView4.GetRowCellValue(i, "SMPLOTNO")+"'";
                        strSQL += ")";
                        db.Execute(strSQL);
                    }
                }
                db.CommitTrans();
                XtraMessageBox.Show("Save complete.", "Save", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                db.RollbackTrans();
                XtraMessageBox.Show(ex.Message, "Error--Save BOMDetail", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            db.ConnectionClose();
        }



        private void ClearBomList()
        {

        }
        private void ClearSampleRequestDetail()
        {
            gridControl2.DataSource = null;
            txtCreateBy.Text = "";
            txtCreateDate.Text = "";
            txtUpdateBy.Text = "";
            txtUpdateDate.Text = "";
        }
        private void ClearBOMDetail()
        {
            gridControl3.DataSource = null;

            txtBomNo.Text = "";
            txtReviseNo.Text = "";
            dtpLastDate.EditValue = DateTime.Today;
            txtSmplNo_Header.Text = "";
            txtSmplNo_Header.Tag = null;
            txtPatternNo.Text = "";
            txtPatternSizeZone.Text = "";
            txtItemNo.Text = "";
            txtItemNo.Tag = null;
            txtModelName.Text = "";
            txtStyleName.Text = "";
            txtStyleName.Tag = null;
            txtCategory.Text = "";
            txtCategory.Tag = null;
            txtSeason_Header.Text = "";
            txtCustomer_Header.Text = "";
            txtCustomer_Header.Tag = null;
            sleFGProductCode.EditValue = null;
            txtColor.Text = "";
            txtColor.Tag = null;
            txtSize.Text = "";
            txtSize.Tag = null;
            gleUnit.EditValue = 15;
            txtUnitCost.Text = "0";
            optStatus.SelectedIndex = -1;
            txtCostsheetNo.Text = "";
            treeBom.DataSource = null;

            dtBomDetail = new DataTable();
            dtBomDetail.BeginInit();
            dtBomDetail.Columns.Add("X", typeof(bool));
            dtBomDetail.Columns.Add("ROWID", typeof(int));
            dtBomDetail.Columns.Add("TYPE", typeof(string));
            dtBomDetail.Columns.Add("OIDITEM", typeof(int));
            dtBomDetail.Columns.Add("ITEMNO", typeof(string));
            dtBomDetail.Columns.Add("COMPOSITION", typeof(string));
            dtBomDetail.Columns.Add("OIDCOLOR", typeof(int));
            dtBomDetail.Columns.Add("COLOR", typeof(string));
            dtBomDetail.Columns.Add("OIDSIZE", typeof(int));
            dtBomDetail.Columns.Add("SIZE", typeof(string));
            dtBomDetail.Columns.Add("OIDUNIT", typeof(int));
            dtBomDetail.Columns.Add("UNIT", typeof(string));
            dtBomDetail.Columns.Add("OIDCURR", typeof(int));
            dtBomDetail.Columns.Add("CURRENCY", typeof(string));
            dtBomDetail.Columns.Add("CONSUMPTION", typeof(decimal));
            dtBomDetail.Columns.Add("PRICE", typeof(decimal));
            dtBomDetail.Columns.Add("COST", typeof(decimal));
            dtBomDetail.Columns.Add("OIDVEND", typeof(int));
            dtBomDetail.Columns.Add("VENDOR_NAME", typeof(string));
            dtBomDetail.Columns.Add("VENDOR_CODE", typeof(string));
            dtBomDetail.Columns.Add("OIDDEPT", typeof(int));
            dtBomDetail.Columns.Add("DEPARTMENT", typeof(string));
            dtBomDetail.Columns.Add("SMPLOTNO", typeof(string));
            dtBomDetail.EndInit();
            gridControl4.DataSource = dtBomDetail;

            gridView4.OptionsClipboard.PasteMode = DevExpress.Export.PasteMode.Append;
            gridView4.OptionsView.EnableAppearanceEvenRow = true;
            gridView4.OptionsView.EnableAppearanceOddRow = true;
            gridView4.OptionsView.ColumnAutoWidth = false;
            gridView4.BestFitColumns();
        }
        private void ClearBOMDetail_Detail()
        {
            txtListNo.Text = "";
            txtMaterialType.Text = "";
            txtItemNo_Detail.Text = "";
            txtMatColor.Text = "";
            txtMatSize.Text = "";
            txtComposition.Text = "";
            txtCurrency.Text = "";
            txtPrice.Text = "0";
            txtComposition.Text = "0";
            txtCost.Text = "0";
            txtVendor.Text = "";
            txtVendMatCode.Text = "";
            txtSmplLotNo.Text = "";
            txtWorkStation.Text = "";
            txtMatLoss.Text = "0";
            txtMatUnit.Text = "";
        }
        private void GetUnit()
        {
            string strSQL = "SELECT OIDUNIT,UNITNAME FROM UNIT";
            DataTable dt = db.GetDataTable(strSQL);
            gleUnit.Properties.DataSource = dt;
            gleUnit.Properties.PopulateViewColumns();
            gleUnit.Properties.DisplayMember = "UNITNAME";
            gleUnit.Properties.ValueMember = "OIDUNIT";
            gleUnit.Properties.View.Columns["OIDUNIT"].Visible = false;
            gleUnit.Properties.BestFitMode = BestFitMode.BestFitResizePopup;
            
        }
        private void GetFGItem()
        {
            string strSQL = "SELECT OIDITEM,CODE FROM ITEMS WHERE MATERIALTYPE=0";
            dtFG = db.GetDataTable(strSQL);
            //sleFGProductCode.Properties.DataSource = null;
            sleFGProductCode.Properties.DataSource = dtFG;
            sleFGProductCode.Properties.PopulateViewColumns();
            sleFGProductCode.Properties.DisplayMember = "CODE";
            sleFGProductCode.Properties.ValueMember = "OIDITEM";
            sleFGProductCode.Properties.View.Columns["OIDITEM"].Visible = false;
            sleFGProductCode.Properties.BestFitMode = BestFitMode.BestFitResizePopup;
        }
        private void GetBOMList(int userID)
        {
            string strSQL = "EXEC spDEV03_GetBOM "+userID;
            DataTable dt = db.GetDataTable(strSQL);
            gridControl1.DataSource = dt;
            if (dt == null) return;
            gridView1.Columns["SMPLITEM"].Caption = "SMPL Item";
            gridView1.Columns["SMPLNO"].Caption = "SMPL No.";
            gridView1.Columns["SEASON"].Caption = "Season";
            gridView1.Columns["CUSTOMER"].Caption = "Customer";
            gridView1.Columns["ITEM"].Caption = "Item";
            gridView1.Columns["CATEGORYNAME"].Caption = "Category";
            gridView1.Columns["STYLENAME"].Caption = "Style";
            gridView1.Columns["SMPLPATTERNNO"].Caption = "Pattern No.";
            gridView1.Columns["USERNAME"].Caption = "Created By";
            gridView1.OptionsView.EnableAppearanceEvenRow = true;
            gridView1.OptionsView.EnableAppearanceOddRow = true;
            gridView1.OptionsView.ColumnAutoWidth = false;
            gridView1.BestFitColumns();
        }
        private void GetSampleRequest(int oidSMPL)
        {
            string strSQL = "EXEC spDEV03_GetSampleRequest " + oidSMPL;
            DataTable dt = db.GetDataTable(strSQL);
            gridControl2.DataSource = dt;
            if (dt == null) return;
            foreach (DataRow dr in dt.Rows)
            {
                txtCreateBy.Text = dr["CREATEDBY"].ToString();
                txtCreateDate.Text = dr["CREATEDDATE"].ToString();
                txtUpdateBy.Text = dr["UPDATEDBY"].ToString();
                txtUpdateDate.Text = dr["UPDATEDDATE"].ToString();
            }
            gridView2.Columns["OIDSMPLDT"].Caption = "ID";
            gridView2.Columns["SMPLITEM"].Caption = "Item No.";
            gridView2.Columns["COLORNAME"].Caption = "Color";
            gridView2.Columns["SIZENAME"].Caption = "Size";
            gridView2.Columns["OIDSMPLDT"].Visible = false;
            gridView2.Columns["OIDSMPL"].Visible = false;
            gridView2.Columns["OIDCOLOR"].Visible = false;
            gridView2.Columns["OIDSIZE"].Visible = false;
            gridView2.Columns["CREATEDBY"].Visible = false;
            gridView2.Columns["CREATEDDATE"].Visible = false;
            gridView2.Columns["UPDATEDBY"].Visible = false;
            gridView2.Columns["UPDATEDDATE"].Visible = false;
            gridView2.OptionsView.EnableAppearanceEvenRow = true;
            gridView2.OptionsView.EnableAppearanceOddRow = true;
            gridView2.OptionsView.ColumnAutoWidth = false;
            gridView2.BestFitColumns();

        }
        private void GetBOMDetail(int oidSMPL, int oidSize, int oidColor)
        {
            string strSQL = "EXEC spDEV03_GetBOMDetail " + oidSMPL + "," + oidSize + "," + oidColor;
            DataSet ds = db.GetDataSet(strSQL);
            if (ds == null) return;
            gridControl3.DataSource = ds.Tables[0];
            gridView3.PopulateColumns();
            
            gridView3.Columns["ROWID"].Visible = false;
            gridView3.Columns["OIDITEM"].Visible = false;
            gridView3.Columns["OIDCOLOR"].Visible = false;
            gridView3.Columns["OIDSIZE"].Visible = false;
            gridView3.Columns["OIDUNIT"].Visible = false;
            gridView3.Columns["OIDCURR"].Visible = false;
            gridView3.Columns["OIDVEND"].Visible = false;
            gridView3.Columns["OIDDEPT"].Visible = false;

            gridView3.Columns["TYPE"].Caption = "Type";
            gridView3.Columns["ITEMNO"].Caption = "Item No.";
            gridView3.Columns["COMPOSITION"].Caption = "Composition";
            gridView3.Columns["COLOR"].Caption = "Color";
            gridView3.Columns["SIZE"].Caption = "Size";
            gridView3.Columns["UNIT"].Caption = "Unit";
            gridView3.Columns["CURRENCY"].Caption = "Currency";
            gridView3.Columns["CONSUMPTION"].Caption = "Consumption";
            gridView3.Columns["PRICE"].Caption = "Price";
            gridView3.Columns["COST"].Caption = "Cost";
            gridView3.Columns["VENDOR_NAME"].Caption = "Vendor";
            gridView3.Columns["VENDOR_CODE"].Caption = "Vendor Mat.Code";
            gridView3.Columns["DEPARTMENT"].Caption = "Work Station";
            gridView3.Columns["SMPLOTNO"].Caption = "Sample Lot No.";
            gridView3.Columns["LOSS"].Caption = "% Material Loss";

            gridView3.OptionsView.EnableAppearanceEvenRow = true;
            gridView3.OptionsView.EnableAppearanceOddRow = true;
            gridView3.OptionsView.ColumnAutoWidth = false;
            gridView3.BestFitColumns();
            foreach (DataRow dr in ds.Tables[1].Rows)
            {
                txtBomNo.EditValue = dr["BOMNO"];
                txtBomNo.Tag = dr["OIDBOM"];
                txtReviseNo.EditValue = dr["REVISIONNO"];
                dtpLastDate.EditValue = dr["ISSUEDATE"] == System.DBNull.Value ? (DateTime?)null : (DateTime)dr["ISSUEDATE"];
                txtSmplNo_Header.EditValue = dr["SMPLNO"];
                txtSmplNo_Header.Tag = dr["OIDSMPL"];
                txtPatternNo.EditValue = dr["SMPLPATTERNNO"];
                txtPatternSizeZone.EditValue = dr["PATTERNSIZEZONE"];
                txtItemNo.EditValue = dr["SMPLITEM"];
                txtModelName.EditValue = dr["MODELNAME"];
                txtStyleName.EditValue = dr["STYLENAME"];
                txtStyleName.Tag = dr["OIDSTYLE"];
                txtCategory.EditValue = dr["CATEGORYNAME"];
                txtCategory.Tag = dr["OIDCATEGORY"];
                txtSeason_Header.EditValue = dr["SEASON"];
                txtCustomer_Header.EditValue = dr["SHORTNAME"];
                txtCustomer_Header.Tag = dr["OIDCUST"];
                sleFGProductCode.EditValue = dr["OIDITEM"];
                txtColor.EditValue = dr["COLORNAME"];
                txtColor.Tag = dr["OIDCOLOR"];
                txtSize.EditValue = dr["SIZENAME"];
                txtSize.Tag = dr["OIDSIZE"];
                gleUnit.EditValue = dr["OIDUNIT"];
                txtCost.EditValue = dr["COST"];
                optStatus.SelectedIndex = Convert.ToInt32(dr["STATUS"]);
                txtCostsheetNo.EditValue = dr["SMPLITEM"];
            }
            gridControl4.DataSource = ds.Tables[2];
            gridView4.PopulateColumns();

            gridView4.Columns["X"].Visible = false;
            gridView4.Columns["ROWID"].Visible = false;
            gridView4.Columns["OIDITEM"].Visible = false;
            gridView4.Columns["OIDCOLOR"].Visible = false;
            gridView4.Columns["OIDSIZE"].Visible = false;
            gridView4.Columns["OIDUNIT"].Visible = false;
            gridView4.Columns["OIDVEND"].Visible = false;
            gridView4.Columns["OIDDEPT"].Visible = false;

            gridView4.Columns["TYPE"].Caption = "Type";
            gridView4.Columns["ITEMNO"].Caption = "Item No.";
            gridView4.Columns["COMPOSITION"].Caption = "Composition";
            gridView4.Columns["COLOR"].Caption = "Color";
            gridView4.Columns["SIZE"].Caption = "Size";
            gridView4.Columns["UNIT"].Caption = "Unit";
            gridView4.Columns["CURRENCY"].Caption = "Currency";
            gridView4.Columns["CONSUMPTION"].Caption = "Consumption";
            gridView4.Columns["PRICE"].Caption = "Price";
            gridView4.Columns["COST"].Caption = "Cost";
            gridView4.Columns["VENDOR_NAME"].Caption = "Vendor";
            gridView4.Columns["VENDOR_CODE"].Caption = "Vendor Mat.Code";
            gridView4.Columns["DEPARTMENT"].Caption = "Work Station";
            gridView4.Columns["SMPLOTNO"].Caption = "Sample Lot No.";
            gridView4.Columns["LOSS"].Caption = "% Material Loss";

            gridView4.OptionsView.EnableAppearanceEvenRow = true;
            gridView4.OptionsView.EnableAppearanceOddRow = true;
            gridView4.OptionsView.ColumnAutoWidth = false;
            gridView4.BestFitColumns();

            //ทำเครื่องหมายเช็คใน gridview 2 ที่มีรายการตรงกับใน gridview4
            for (int i = 0; i < gridView4.DataRowCount; i++)
            {
                string itemNo = gridView4.GetRowCellValue(i, "ITEMNO").ToString();
                for (int j = 0; j < gridView3.DataRowCount; j++)
                {
                    if (Equals(itemNo, gridView3.GetRowCellValue(j, "ITEMNO")))
                    {
                        gridView3.SetRowCellValue(j, "X", true);
                        break;
                    }
                }
            }
        }

        private void MyStyleChanged(object sender, EventArgs e)
        {
            UserLookAndFeel userLookAndFeel = (UserLookAndFeel)sender;
            cUtility.SaveRegistry(@"Software\MDS", "SkinName", userLookAndFeel.SkinName);
            cUtility.SaveRegistry(@"Software\MDS", "SkinPalette", userLookAndFeel.ActiveSvgPaletteName);
        }
        private void DEV03_Load(object sender, EventArgs e)
        {
            UserLookAndFeel.Default.StyleChanged += MyStyleChanged;
            IniFile ini = new IniFile(@"\\192.168.101.3\Software_tuw\PTC-MDS\FileConfig\Configue.ini");
            db = new cDatabase("Server=" + ini.Read("Server", "ConnectionString") + ";uid=" + ini.Read("Uid", "ConnectionString") + ";pwd=" + ini.Read("Pwd", "ConnectionString") + ";database=" + ini.Read("Database", "ConnectionString"));
            dtfinfo = clinfo.DateTimeFormat;
            try
            {
                NewData();
                GetBOMList(0);
                GetUnit();
                GetFGItem();

                tabbed_Master.SelectedTabPageIndex = 0;
                tabbedBom.SelectedTabPageIndex = 0;
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


            ////***** SET CONNECT DB ********
            //if (this.ConnectionString != null)
            //{
            //    if (this.ConnectionString != "")
            //    {
            //        CONNECT_STRING = this.ConnectionString;
            //    }
            //}

            //this.DBC = new DatabaseConnect(CONNECT_STRING);

            //if (this.DBC.chkCONNECTION_STING() == false)
            //{
            //    this.DBC.setCONNECTION_STRING_INIFILE();
            //    if (this.DBC.chkCONNECTION_STING() == false)
            //    {
            //        return;
            //    }
            //}
            //new ObjDE.setDatabase(this.DBC);
            //******************************

            lblUser.Text = "Login : " + UserLogin.FullName;
            //StringBuilder sbSQL = new StringBuilder();
            //sbSQL.Append("SELECT TOP (1) ReadWriteStatus FROM FunctionAccess WHERE (OIDUser = '" + UserLogin.OIDUser + "') AND(FunctionNo = 'DEV03') ");
            //chkReadWrite = this.DBC.DBQuery(sbSQL).getInt();

            //if (chkReadWrite == 0)
            //{
            //    ribbonPageGroup1.Visible = false;
            //    //rpgManage.Visible = false;

            //    layoutControlItem29.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            //    //simpleButton2.Enabled = false;
            //    //simpleButton3.Enabled = false;
            //    //simpleButton4.Enabled = false;
            //    //sbColor.Enabled = false;
            //    //sbSize.Enabled = false;
            //    //sbFBColor.Enabled = false;
            //    //sbTempCode.Enabled = false;
            //    //sbMTColor.Enabled = false;
            //    //sbTempCodeMat.Enabled = false;
            //    //btnOpenImg_Main.Enabled = false;
            //    //sbDelete_S.Enabled = false;
            //    //sbClear.Enabled = false;
            //    //simpleButton5.Enabled = false;
            //    //sbDelete_F.Enabled = false;
            //    //sbMatClear.Enabled = false;
            //    //btnUploadMat.Enabled = false;
            //    //simpleButton1.Enabled = false;

            //    //sbUseFor.Enabled = false;
            //    //sbUnit.Enabled = false;

            //    //sbPart.Enabled = false;
            //    //sbFBSupplier.Enabled = false;

            //    //sbMTSupplier.Enabled = false;
            //}

            //LoadListBOM();
        }
        private void gridView1_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator) e.Info.DisplayText = (e.RowHandle + 1).ToString();
            gridView1.IndicatorWidth = 45;
        }
        private void gridView2_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator) e.Info.DisplayText = (e.RowHandle + 1).ToString();
            gridView2.IndicatorWidth = 45;
        }
        private void gridView3_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator) e.Info.DisplayText = (e.RowHandle + 1).ToString();
            gridView3.IndicatorWidth = 45;
        }
        private void gridView4_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator) e.Info.DisplayText = (e.RowHandle + 1).ToString();
            gridView4.IndicatorWidth = 45;
        }
        private void gridView1_ShowingEditor(object sender, System.ComponentModel.CancelEventArgs e)
        {
            e.Cancel = true;
        }
        private void gridView2_ShowingEditor(object sender, System.ComponentModel.CancelEventArgs e)
        {
            e.Cancel = true;
        }
        private void gridView3_ShowingEditor(object sender, System.ComponentModel.CancelEventArgs e)
        {
            GridView gv = (GridView)sender;
            if (gv.FocusedColumn.FieldName == "X")
                e.Cancel = false;
            else
                e.Cancel = true;
        }
        private void gridView4_ShowingEditor(object sender, System.ComponentModel.CancelEventArgs e)
        {
            e.Cancel = true;
        }
        private void gridView1_DoubleClick(object sender, EventArgs e)
        {
            DXMouseEventArgs ea = e as DXMouseEventArgs;
            GridView view = sender as GridView;
            GridHitInfo info = view.CalcHitInfo(ea.Location);
            if (info.InRow || info.InRowCell)
            {
                //string colCaption = info.Column == null ? "N/A" : info.Column.GetCaption();
                ClearSampleRequestDetail();
                ClearBOMDetail();
                ClearBOMDetail_Detail();
                GetSampleRequest(Convert.ToInt32(view.GetRowCellValue(info.RowHandle, "OIDSMPL")));
                tabbed_Master.SelectedTabPageIndex = 1;
                //MessageBox.Show(string.Format("DoubleClick on row: {0}, column: {1}.", info.RowHandle, colCaption));
            }
        }
        private void gridView2_RowCellClick(object sender, RowCellClickEventArgs e)
        {
            try
            {
                ClearBOMDetail();
                ClearBOMDetail_Detail();
                GetBOMDetail(Convert.ToInt32(gridView2.GetRowCellValue(e.RowHandle, "OIDSMPL")), Convert.ToInt32(gridView2.GetRowCellValue(e.RowHandle, "OIDSIZE")), Convert.ToInt32(gridView2.GetRowCellValue(e.RowHandle, "OIDCOLOR")));
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void gridView4_RowCellClick(object sender, RowCellClickEventArgs e)
        {
            try
            {
                txtListNo.EditValue = gridView4.GetFocusedRowCellValue("OIDBOMDT");
                txtMaterialType.EditValue = gridView4.GetFocusedRowCellValue("TYPE");
                txtItemNo_Detail.EditValue = gridView4.GetFocusedRowCellValue("ITEMNO");
                txtMatColor.EditValue = gridView4.GetFocusedRowCellValue("COLOR");
                txtMatSize.EditValue = gridView4.GetFocusedRowCellValue("SIZE");
                txtComposition.EditValue = gridView4.GetFocusedRowCellValue("COMPOSITION");
                txtCurrency.EditValue = gridView4.GetFocusedRowCellValue("CURRENCY");
                txtPrice.EditValue = gridView4.GetFocusedRowCellValue("PRICE");
                txtConsumption.EditValue = gridView4.GetFocusedRowCellValue("CONSUMPTION");
                txtCost.EditValue = gridView4.GetFocusedRowCellValue("COST");
                txtVendor.EditValue = gridView4.GetFocusedRowCellValue("VENDOR_NAME");
                txtVendMatCode.EditValue = gridView4.GetFocusedRowCellValue("VENDOR_CODE");
                txtSmplLotNo.EditValue = gridView4.GetFocusedRowCellValue("SMPLOTNO");
                txtWorkStation.EditValue = gridView4.GetFocusedRowCellValue("DEPARTMENT");
                txtMatLoss.EditValue = gridView4.GetFocusedRowCellValue("LOSS");
                txtMatUnit.EditValue = gridView4.GetFocusedRowCellValue("UNIT");
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void btnRefreshFG_Click(object sender, EventArgs e)
        {
            GetFGItem();
        }
        private void gridView3_CellValueChanging(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            if (e.Column.FieldName == "X")
            {
                var selected = (bool)e.Value;
                var itemNO = (string)gridView3.GetRowCellValue(e.RowHandle, "ITEMNO");
                if (selected)
                {
                    for (int i = 0; i < gridView4.DataRowCount; i++)
                    {
                        if (Equals(itemNO, gridView4.GetRowCellValue(i, "ITEMNO"))) return;
                    }
                    gridView4.AddNewRow();
                    gridView4.FocusedRowHandle = DevExpress.XtraGrid.GridControl.NewItemRowHandle;
                    for (int j = 0; j < gridView3.Columns.Count; j++)
                    {
                        gridView4.SetFocusedRowCellValue(gridView3.Columns[j].FieldName, gridView3.GetFocusedRowCellValue(gridView3.Columns[j].FieldName));
                    }
                    gridView4.UpdateCurrentRow();
                    //gridView3.CopyToClipboard();
                    //gridView4.PasteFromClipboard();
                }
                //else
                //{
                //    for (int i = 0; i < gridView4.DataRowCount; i++)
                //    {
                //        if (Equals(itemNO, gridView4.GetRowCellValue(i, "ITEMNO"))) gridView4.DeleteRow(i); 
                //    }

                //}
                
                
            }
        }
        private void txtComposition_EditValueChanged(object sender, EventArgs e)
        {
            gridView4.SetFocusedRowCellValue("COMPOSITION", txtComposition.EditValue);
        }
        private void txtPrice_EditValueChanged(object sender, EventArgs e)
        {
            gridView4.SetFocusedRowCellValue("PRICE", txtPrice.EditValue);
        }
        private void txtConsumption_EditValueChanged(object sender, EventArgs e)
        {
            gridView4.SetFocusedRowCellValue("CONSUMPTION",txtConsumption.EditValue);
        }
        private void txtCost_EditValueChanged(object sender, EventArgs e)
        {
            gridView4.SetFocusedRowCellValue("COST", txtCost.EditValue);
        }
        private void txtVendMatCode_EditValueChanged(object sender, EventArgs e)
        {
            gridView4.SetFocusedRowCellValue("VENDOR_CODE", txtVendMatCode.EditValue);
        }
        private void txtSmplLotNo_EditValueChanged(object sender, EventArgs e)
        {
            gridView4.SetFocusedRowCellValue("SMPLOTNO", txtSmplLotNo.EditValue);
        }
        private void txtMatLoss_EditValueChanged(object sender, EventArgs e)
        {
            gridView4.SetFocusedRowCellValue("LOSS", txtMatLoss.EditValue);
        }
        private void optUser_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (optUser.SelectedIndex == 0)
                GetBOMList(UserLogin.OIDUser);
            else
                GetBOMList(0);
        }
        private void gridControl4_ProcessGridKey(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                if (gridView4.IsEditing == false) { gridView4.DeleteSelectedRows(); }
            }
        }
        private void gridView4_RowDeleting(object sender, DevExpress.Data.RowDeletingEventArgs e)
        {
            string itemNO = ((DataRowView)e.Row)["ITEMNO"].ToString();
            for (int i = 0; i < gridView3.DataRowCount; i++)
            {
                if (Equals(itemNO, gridView3.GetRowCellValue(i, "ITEMNO")))
                {
                    gridView3.SetRowCellValue(i, "X", false);
                    break;
                }
            }
        }

        private void bbiNew_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            NewData();
        }
        private void bbiSave_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            SaveData();
        }
        private void bbiRefresh_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            //GetBOMList();
        }
        private void bbiPrintPreview_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            //gcPTerm.ShowPrintPreview();
        }
        private void bbiPrint_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            //gcPTerm.Print();
        }
        private void bbiExcel_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            //string pathFile = new ObjSet.Folder(@"C:\MDS\Export\").GetPath() + "PaymentTermList_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
            //gvPTerm.ExportToXlsx(pathFile);
            //System.Diagnostics.Process.Start(pathFile);
        }

        
        private bool chkDuplicate()
        {
            bool chkDup = true;
            //if (txeName.Text != "")
            //{
            //    txeName.Text = txeName.Text.Trim();
            //    if (txeName.Text.Trim() != "" && lblStatus.Text == "* Add Payment Term")
            //    {
            //        StringBuilder sbSQL = new StringBuilder();
            //        sbSQL.Append("SELECT TOP(1) Name FROM PaymentTerm WHERE (Name = N'" + txeName.Text.Trim() + "') ");
            //        if (new DBQuery(sbSQL).getString() != "")
            //        {
            //            FUNC.msgWarning("Duplicate payment term. !! Please Change.");
            //            txeName.Text = "";
            //            chkDup = false;
            //        }
            //    }
            //    else if (txeName.Text.Trim() != "" && lblStatus.Text == "* Edit Payment Term")
            //    {
            //        StringBuilder sbSQL = new StringBuilder();
            //        sbSQL.Append("SELECT TOP(1) OIDPayment ");
            //        sbSQL.Append("FROM PaymentTerm ");
            //        sbSQL.Append("WHERE (Name = N'" + txeName.Text.Trim() + "') ");
            //        string strCHK = new DBQuery(sbSQL).getString();
            //        if (strCHK != "" && strCHK != txeID.Text.Trim())
            //        {
            //            FUNC.msgWarning("Duplicate payment term. !! Please Change.");
            //            txeName.Text = "";
            //            chkDup = false;
            //        }
            //    }
            //}
            return chkDup;
        }

        



        

       

        

        

        

        
       
    }
}