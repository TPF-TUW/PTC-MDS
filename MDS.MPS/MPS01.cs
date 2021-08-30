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
using DevExpress.Spreadsheet;
using System.IO;

namespace MDS.MPS
{
    public partial class MPS01 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private Functionality.Function FUNC = new Functionality.Function();

        StringBuilder sbSTATUS = new StringBuilder();
        DataTable dtINPUT = new DataTable();
        public LogIn UserLogin { get; set; }
        public int Company { get; set; }


        public string ConnectionString { get; set; }
        string CONNECT_STRING = "";
        DatabaseConnect DBC;


        public MPS01()
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

            sbSTATUS.Clear();
            sbSTATUS.Append("SELECT '1' AS ID, 'Ph1 - New Order' AS Status ");
            sbSTATUS.Append("UNION ALL ");
            sbSTATUS.Append("SELECT '2' AS ID, 'Ph2 - Repeat Order' AS Status ");
            sbSTATUS.Append("UNION ALL ");
            sbSTATUS.Append("SELECT '3' AS ID, 'Ph3 - Urgent, (Adjust) Revise(+/-)' AS Status ");
            sbSTATUS.Append("UNION ALL ");
            sbSTATUS.Append("SELECT '4' AS ID, 'Cancel' AS Status ");
            sbSTATUS.Append("UNION ALL ");
            sbSTATUS.Append("SELECT '5' AS ID, 'End Order' AS Status ");

            glueUnit.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard;
            glueUnit.Properties.AcceptEditorTextAsNewValue = DevExpress.Utils.DefaultBoolean.True;

            glueLogisticsType.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard;
            glueLogisticsType.Properties.AcceptEditorTextAsNewValue = DevExpress.Utils.DefaultBoolean.True;


            dtINPUT.Columns.Add("ProductionPlanID", typeof(String));
            dtINPUT.Columns.Add("OIDCUST", typeof(String));
            dtINPUT.Columns.Add("FileOrderDate", typeof(DateTime));
            dtINPUT.Columns.Add("Year", typeof(Int32));
            dtINPUT.Columns.Add("Season", typeof(String));
            dtINPUT.Columns.Add("BusinessUnit", typeof(String));
            dtINPUT.Columns.Add("OIDCSITEM", typeof(String));
            dtINPUT.Columns.Add("ItemName", typeof(String));
            dtINPUT.Columns.Add("StyleNo", typeof(String));
            dtINPUT.Columns.Add("ModelNo", typeof(String));
            dtINPUT.Columns.Add("OIDVEND", typeof(String));
            dtINPUT.Columns.Add("SewingDifficulty", typeof(String));
            dtINPUT.Columns.Add("ProductionPlanType", typeof(String));
            dtINPUT.Columns.Add("DataUpdate", typeof(DateTime));
            dtINPUT.Columns.Add("BookingFabric", typeof(Int32));
            dtINPUT.Columns.Add("BookingAccessory", typeof(Int32));
            dtINPUT.Columns.Add("LastUpdate", typeof(DateTime));
            dtINPUT.Columns.Add("RequestedWHDate", typeof(DateTime));
            dtINPUT.Columns.Add("ContractedDate", typeof(DateTime));
            dtINPUT.Columns.Add("TransportMethod", typeof(String));
            dtINPUT.Columns.Add("LogisticsType", typeof(String));
            dtINPUT.Columns.Add("OrderQty", typeof(Int32));
            dtINPUT.Columns.Add("FabricOrderNO", typeof(String));
            dtINPUT.Columns.Add("FabricUpdateDate", typeof(DateTime));
            dtINPUT.Columns.Add("FabricActualOrderQty", typeof(Int32));
            dtINPUT.Columns.Add("ColorOrderNO", typeof(String));
            dtINPUT.Columns.Add("ColorUpdateDate", typeof(DateTime));
            dtINPUT.Columns.Add("ColorActualOrderQty", typeof(Int32));
            dtINPUT.Columns.Add("TrimOrderNO", typeof(String));
            dtINPUT.Columns.Add("TrimUpdateDate", typeof(DateTime));
            dtINPUT.Columns.Add("TrimActualOrderQty", typeof(Int32));
            dtINPUT.Columns.Add("POOrderNO", typeof(String));
            dtINPUT.Columns.Add("POUpdateDate", typeof(DateTime));
            dtINPUT.Columns.Add("POActualOrderQty", typeof(Int32));
            dtINPUT.Columns.Add("OrderQTYOld", typeof(Int32));
            gcINPUT.DataSource = dtINPUT;

            LoadData();
            NewData();


            tabbedControlGroup1.SelectedTabPage = layoutControlGroup1;
            //*******************************************************
            bbiRefresh.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
            bbiNew.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
            bbiSave.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
            bbiEdit.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
            bbiClone.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;

            ribbonPageGroup2.Visible = true;
            ribbonPageGroup5.Visible = true;
            //*******************************************************
        }

        private void LoadSummary()
        {
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT COF.OIDFC AS ID, COF.ProductionPlanID, COF.Season, COF.OIDCUST AS [Customer ID], CUS.Name AS Customer, COF.BusinessUnit, COF.OIDCSITEM, ITC.ItemCode, ITC.ItemName, COF.ModelNo AS MainSampleCode, ");
            sbSQL.Append("       COF.OIDVEND AS SupplierID, VD.Name AS Supplier, COF.SewingDifficulty, COF.ProductionPlanType, COF.Status AS StatusID, STS.Status, ");
            sbSQL.Append("       CASE WHEN COF.LastUpdate IS NULL THEN '' ELSE CONVERT(VARCHAR(10), COF.LastUpdate, 103) END AS LastUpdateDate, ");
            sbSQL.Append("       CASE WHEN COF.RequestedWHDate IS NULL THEN '' ELSE CONVERT(VARCHAR(10), COF.RequestedWHDate, 103) END AS RequestedWHDate, ");
            sbSQL.Append("       CASE WHEN COF.ContractedDate IS NULL THEN '' ELSE CONVERT(VARCHAR(10), COF.ContractedDate, 103) END AS ContractedDate, ");
            sbSQL.Append("       COF.TransportMethod AS TransportMethodName, COF.LogisticsType AS LogisticsTypeName, COF.OrderQty AS[Order Qty(Pcs)], COF.FabricOrderNO AS[FabricOrderNo.], COF.FabricActualOrderQty AS[FB Actual Order Qty(Pcs)], ");
            sbSQL.Append("       CASE WHEN COF.FabricUpdateDate IS NULL THEN '' ELSE CONVERT(VARCHAR(10), COF.FabricUpdateDate, 103) END AS [FB Update Date], ");
            sbSQL.Append("       COF.TrimOrderNO AS[TrimOrderNo.], COF.TrimActualOrderQty AS[TrimOrderActualQty(Pcs)], ");
            sbSQL.Append("       CASE WHEN COF.TrimUpdateDate IS NULL THEN '' ELSE CONVERT(VARCHAR(10), COF.TrimUpdateDate, 103) END AS TrimUpdateDate, ");
            sbSQL.Append("       COF.POOrderNO AS[PO Order No.], COF.POActualOrderQty AS[PO Actual Order Qty(Pcs)], ");
            sbSQL.Append("       CASE WHEN COF.POUpdateDate IS NULL THEN '' ELSE CONVERT(VARCHAR(10), COF.POUpdateDate, 103) END AS POUpdateDate, ");
            sbSQL.Append("       COF.ColorOrderNO AS[Color Order No.], COF.ColorActualOrderQty AS[Color Actual Order Qty(Pcs)], ");
            sbSQL.Append("       CASE WHEN COF.ColorUpdateDate IS NULL THEN '' ELSE CONVERT(VARCHAR(10), COF.ColorUpdateDate, 103) END AS ColorUpdateDate, ");
            sbSQL.Append("       COF.OrderQTYOld AS[Order Qty(Old)], COF.BookingFabric AS BFID, CASE WHEN COF.BookingFabric=0 THEN 'No' ELSE 'Yes' END AS BookingFabric, COF.BookingAccessory AS BAID, CASE WHEN COF.BookingAccessory=0 THEN 'No' ELSE 'Yes' END AS BookingAccessory, ");
            sbSQL.Append("       CASE WHEN COF.FileOrderDate IS NULL THEN '' ELSE CONVERT(VARCHAR(10), COF.FileOrderDate, 103) END AS FileOrderDate, ");
            sbSQL.Append("       CASE WHEN COF.DataUpdate IS NULL THEN '' ELSE CONVERT(VARCHAR(10), COF.DataUpdate, 103) END AS DataUpdate, ");
            sbSQL.Append("       COF.CreateBy, COF.CreateDate, COF.UpdateBy, COF.Updatedate ");
            sbSQL.Append("FROM   COForecast AS COF LEFT OUTER JOIN ");
            sbSQL.Append("       Customer AS CUS ON COF.OIDCUST = CUS.OIDCUST LEFT OUTER JOIN ");
            sbSQL.Append("       Vendor AS VD ON COF.OIDVEND = VD.OIDVEND LEFT OUTER JOIN ");
            sbSQL.Append("       ItemCustomer AS ITC ON COF.OIDCSITEM = ITC.OIDCSITEM LEFT OUTER JOIN ");
            sbSQL.Append("       (");
            sbSQL.Append(sbSTATUS);
            sbSQL.Append("       ) AS STS ON COF.Status = STS.ID ");
            sbSQL.Append("ORDER BY COF.ProductionPlanID ");
            new ObjDE.setGridControl(gcFO, gvFO, sbSQL).getData(false, false, false, true);

            gvFO.Columns["ID"].Visible = false;
            gvFO.Columns["Customer ID"].Visible = false;
            gvFO.Columns["OIDCSITEM"].Visible = false;
            gvFO.Columns["SupplierID"].Visible = false;
            gvFO.Columns["StatusID"].Visible = false;
            gvFO.Columns["BFID"].Visible = false;
            gvFO.Columns["BAID"].Visible = false;

            gvFO.Columns["SewingDifficulty"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            gvFO.Columns["Order Qty(Pcs)"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            gvFO.Columns["FB Actual Order Qty(Pcs)"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            gvFO.Columns["TrimOrderActualQty(Pcs)"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            gvFO.Columns["PO Actual Order Qty(Pcs)"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            gvFO.Columns["Color Actual Order Qty(Pcs)"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            gvFO.Columns["Order Qty(Old)"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;

        }

        private void LoadData()
        {
            new ObjDE.setGridLookUpEdit(glueStatus, sbSTATUS, "Status", "ID").getData();

            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT Code, Name AS Supplier, OIDVEND AS ID ");
            sbSQL.Append("FROM Vendor ");
            sbSQL.Append("WHERE (VendorType = 6) ");
            sbSQL.Append("ORDER BY Code ");
            new ObjDE.setSearchLookUpEdit(slueSupplier, sbSQL, "Supplier", "ID").getData();

            sbSQL.Clear();
            sbSQL.Append("SELECT Code, Name AS Customer, OIDCUST AS ID ");
            sbSQL.Append("FROM Customer ");
            sbSQL.Append("ORDER BY Code ");
            new ObjDE.setSearchLookUpEdit(slueCustomer, sbSQL, "Customer", "ID").getData();

            sbSQL.Clear();
            sbSQL.Append("SELECT ITC.ItemCode, ITC.ItemName, CUS.Name AS Customer, ITC.StyleNo AS [StyleNo.], ITC.Season, ITC.OIDCSITEM AS ID ");
            sbSQL.Append("FROM   ItemCustomer AS ITC LEFT OUTER JOIN ");
            sbSQL.Append("       Customer AS CUS ON ITC.OIDCUST = CUS.OIDCUST ");
            sbSQL.Append("ORDER BY ITC.ItemCode ");
            new ObjDE.setSearchLookUpEdit(slueItemCode, sbSQL, "ItemCode", "ID").getData();

            sbSQL.Clear();
            sbSQL.Append("SELECT DISTINCT BusinessUnit ");
            sbSQL.Append("FROM COForecast ");
            sbSQL.Append("ORDER BY BusinessUnit");
            new ObjDE.setGridLookUpEdit(glueUnit, sbSQL, "BusinessUnit", "BusinessUnit").getData();

            sbSQL.Clear();
            sbSQL.Append("SELECT SeasonNo AS [Season No.], SeasonName AS [Season Name] ");
            sbSQL.Append("FROM Season ");
            sbSQL.Append("ORDER BY OIDSEASON");
            new ObjDE.setGridLookUpEdit(glueSeason, sbSQL, "Season No.", "Season No.").getData();

            sbSQL.Clear();
            sbSQL.Append("SELECT TransportMethod ");
            sbSQL.Append("FROM (SELECT DISTINCT TransportMethod ");
            sbSQL.Append("      FROM (SELECT DISTINCT TransportMethod ");
            sbSQL.Append("            FROM COForecast ");
            sbSQL.Append("            UNION ALL ");
            sbSQL.Append("            SELECT '' AS TransportMethod ");
            sbSQL.Append("            UNION ALL ");
            sbSQL.Append("            SELECT 'Air' AS TransportMethod ");
            sbSQL.Append("            UNION ALL ");
            sbSQL.Append("            SELECT 'Ship' AS TransportMethod ");
            sbSQL.Append("            UNION ALL ");
            sbSQL.Append("            SELECT 'Truck' AS TransportMethod) AS DTM) AS TM ");
            sbSQL.Append("ORDER BY CASE TransportMethod ");
            sbSQL.Append("    WHEN '' THEN '0' ");
            sbSQL.Append("    WHEN 'Air' THEN '1' ");
            sbSQL.Append("    WHEN 'Ship' THEN '2' ");
            sbSQL.Append("    WHEN 'Truck' THEN '3' ");
            sbSQL.Append("    ELSE TransportMethod ");
            sbSQL.Append("END ");
            new ObjDE.setGridLookUpEdit(glueTransport, sbSQL, "TransportMethod", "TransportMethod").getData();

            sbSQL.Clear();
            sbSQL.Append("SELECT LogisticsType ");
            sbSQL.Append("FROM (SELECT DISTINCT LogisticsType ");
            sbSQL.Append("      FROM (SELECT DISTINCT LogisticsType ");
            sbSQL.Append("            FROM COForecast ");
            sbSQL.Append("            UNION ALL ");
            sbSQL.Append("            SELECT '' AS LogisticsType ");
            sbSQL.Append("            UNION ALL ");
            sbSQL.Append("            SELECT 'ADC' AS LogisticsType ");
            sbSQL.Append("            UNION ALL ");
            sbSQL.Append("            SELECT 'MDC' AS LogisticsType) AS DTM) AS TM ");
            sbSQL.Append("ORDER BY CASE LogisticsType ");
            sbSQL.Append("    WHEN '' THEN '0' ");
            sbSQL.Append("    WHEN 'ADC' THEN '1' ");
            sbSQL.Append("    WHEN 'MDC' THEN '2' ");
            sbSQL.Append("    ELSE LogisticsType ");
            sbSQL.Append("END ");
            new ObjDE.setGridLookUpEdit(glueLogisticsType, sbSQL, "LogisticsType", "LogisticsType").getData();

            //*** SET GRIDCONTROL COLUMN *****
            repositoryItemSearchLookUpEdit1.DataSource = slueCustomer.Properties.DataSource;
            repositoryItemSearchLookUpEdit1.DisplayMember = "Customer";
            repositoryItemSearchLookUpEdit1.ValueMember = "ID";
            repositoryItemSearchLookUpEdit1.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;

            repositoryItemGridLookUpEdit1.DataSource = glueSeason.Properties.DataSource;
            repositoryItemGridLookUpEdit1.DisplayMember = "Season No.";
            repositoryItemGridLookUpEdit1.ValueMember = "Season No.";
            repositoryItemGridLookUpEdit1.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;

            repositoryItemGridLookUpEdit2.DataSource = glueUnit.Properties.DataSource;
            repositoryItemGridLookUpEdit2.DisplayMember = "BusinessUnit";
            repositoryItemGridLookUpEdit2.ValueMember = "BusinessUnit";
            repositoryItemGridLookUpEdit2.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
            repositoryItemGridLookUpEdit2.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard;
            repositoryItemGridLookUpEdit2.AcceptEditorTextAsNewValue = DevExpress.Utils.DefaultBoolean.True;

            repositoryItemSearchLookUpEdit2.DataSource = slueItemCode.Properties.DataSource;
            repositoryItemSearchLookUpEdit2.DisplayMember = "ItemCode";
            repositoryItemSearchLookUpEdit2.ValueMember = "ItemCode";
            repositoryItemSearchLookUpEdit2.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;

            repositoryItemSearchLookUpEdit3.DataSource = slueSupplier.Properties.DataSource;
            repositoryItemSearchLookUpEdit3.DisplayMember = "Supplier";
            repositoryItemSearchLookUpEdit3.ValueMember = "ID";
            repositoryItemSearchLookUpEdit3.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;

            repositoryItemGridLookUpEdit4.DataSource = glueTransport.Properties.DataSource;
            repositoryItemGridLookUpEdit4.DisplayMember = "TransportMethod";
            repositoryItemGridLookUpEdit4.ValueMember = "TransportMethod";
            repositoryItemGridLookUpEdit4.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;

            repositoryItemGridLookUpEdit3.DataSource = glueLogisticsType.Properties.DataSource;
            repositoryItemGridLookUpEdit3.DisplayMember = "LogisticsType";
            repositoryItemGridLookUpEdit3.ValueMember = "LogisticsType";
            repositoryItemGridLookUpEdit3.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
            repositoryItemGridLookUpEdit3.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard;
            repositoryItemGridLookUpEdit3.AcceptEditorTextAsNewValue = DevExpress.Utils.DefaultBoolean.True;


            LoadSummary();
        }

        private void NewData()
        {
            lblStatus.Text = "* Add Forecast";
            lblStatus.ForeColor = Color.Green;

            txeID.Text = this.DBC.DBQuery("SELECT CASE WHEN ISNULL(MAX(OIDFC), '') = '' THEN 1 ELSE MAX(OIDFC) + 1 END AS NewNo FROM COForecast").getString();
            txePlanID.Text = "";
            slueCustomer.EditValue = "";
            dteOrderDate.EditValue = DateTime.Now;

            if (Convert.ToInt32(DateTime.Now.ToString("yyyy")) > 2500)
                speSeason.Value = Convert.ToInt32(DateTime.Now.ToString("yyyy")) - 543;
            else
                speSeason.Value = Convert.ToInt32(DateTime.Now.ToString("yyyy"));

            glueSeason.EditValue = "";
            glueUnit.EditValue = "";
            slueItemCode.EditValue = "";
            txeItemName.Text = "";
            txeSampleCode.Text = "";
            txeStyle.Text = "";
            slueSupplier.EditValue = "";
            txeSewingDifficulty.Text = "";
            txePlanType.Text = "";
            glueStatus.EditValue = "";
            dteDataUpdate.EditValue = DateTime.Now;

            chkBookFabric.Checked = false;
            chkBookAcc.Checked = false;
            dteLastUpdate.EditValue = DateTime.Now;
            dteWHDate.EditValue = DateTime.Now;
            dteContractDate.EditValue = DateTime.Now;
            glueTransport.EditValue = "";
            glueLogisticsType.EditValue = "";
            txeOrderQty.Text = "";

            txeFabricOrderNo.Text = "";
            dteFabricUpdate.EditValue = DateTime.Now;
            txeFabricQty.Text = "";
            txeColorOrderNo.Text = "";
            dteColorUpdate.EditValue = DateTime.Now;
            txeColorQty.Text = "";

            txeTrimOrderNo.Text = "";
            dteTrimUpdate.EditValue = DateTime.Now;
            txeTrimQty.Text = "";

            txePOOrderNo.Text = "";
            dtePOUpdate.EditValue = DateTime.Now;
            txePOQty.Text = "";

            txeOrderQtyOld.Text = "";

            gcINPUT.Enabled = true;
            sbClearTable.Enabled = true;
            for (int i = gvINPUT.RowCount - 1; i >= 0; i--)
                gvINPUT.DeleteRow(i);
            gvINPUT.OptionsView.ColumnAutoWidth = false;
            gvINPUT.BestFitColumns();
            gvINPUT.Columns["ProductionPlanID"].Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;

            txeFilePath.Text = "";
            spsForecast.CloseCellEditor(DevExpress.XtraSpreadsheet.CellEditorEnterValueMode.Default);
            spsForecast.CreateNewDocument();
            cbeSheet.Properties.Items.Clear();
            cbeSheet.Text = "";

            txePlanID.Focus();
        }

        private void bbiNew_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            LoadData();
            NewData();
        }

        private void bbiSave_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (tabbedControlGroup1.SelectedTabPage == layoutControlGroup2) //Input Form & Table
            {
                gvINPUT.CloseEditor();
                gvINPUT.UpdateCurrentRow();

                if (gvINPUT.HasColumnErrors == true)
                {
                    FUNC.msgError("Can not save. Because found error in table. Please check.");
                }
                else if (CountError() > 0)
                {
                    FUNC.msgWarning("Please enter all information in the table.");
                }
                else
                {
                    if (txePlanID.Text.Trim() == "" && gvINPUT.RowCount == 1)
                    {
                        txePlanID.Focus();
                        FUNC.msgWarning("Please input data.");
                    }
                    else
                    {
                        bool chkPass = true;

                        string strCREATE = "0";
                        //string strCREATE = txeCREATE.Text.Trim() != "" ? txeCREATE.Text.Trim() : "0";

                        StringBuilder sbSQL = new StringBuilder();
                        //Form Input
                        if (txePlanID.Text.Trim() != "")
                        {
                            if (slueCustomer.Text.Trim() == "")
                            {
                                chkPass = false;
                                FUNC.msgWarning("Please select customer.");
                                slueCustomer.Focus();
                            }
                            else if (slueItemCode.Text.Trim() == "")
                            {
                                chkPass = false;
                                FUNC.msgWarning("Please select item code.");
                                slueItemCode.Focus();
                            }
                            else if (slueSupplier.Text.Trim() == "")
                            {
                                chkPass = false;
                                FUNC.msgWarning("Please select raw material supplier.");
                                slueSupplier.Focus();
                            }
                            else
                            {
                                string OrderDate = dteOrderDate.Text.Trim() != "" ? "'" + Convert.ToDateTime(dteOrderDate.EditValue.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL";
                                string DataUpdate = dteDataUpdate.Text.Trim() != "" ? "'" + Convert.ToDateTime(dteDataUpdate.EditValue.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL";
                                string LastUpdate = dteLastUpdate.Text.Trim() != "" ? "'" + Convert.ToDateTime(dteLastUpdate.EditValue.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL";
                                string WHDate = dteWHDate.Text.Trim() != "" ? "'" + Convert.ToDateTime(dteWHDate.EditValue.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL";
                                string ContractDate = dteContractDate.Text.Trim() != "" ? "'" + Convert.ToDateTime(dteContractDate.EditValue.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL";
                                string FabricUpdate = dteFabricUpdate.Text.Trim() != "" ? "'" + Convert.ToDateTime(dteFabricUpdate.EditValue.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL";
                                string ColorUpdate = dteColorUpdate.Text.Trim() != "" ? "'" + Convert.ToDateTime(dteColorUpdate.EditValue.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL";
                                string TrimUpdate = dteTrimUpdate.Text.Trim() != "" ? "'" + Convert.ToDateTime(dteTrimUpdate.EditValue.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL";
                                string POUpdate = dtePOUpdate.Text.Trim() != "" ? "'" + Convert.ToDateTime(dtePOUpdate.EditValue.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL";

                                string SewingDifficulty = txeSewingDifficulty.Text.Trim() != "" ? txeSewingDifficulty.Text.Trim() : "0";
                                string OrderQty = txeOrderQty.Text.Trim() != "" ? txeOrderQty.Text.Trim() : "0";
                                string FabricQty = txeFabricQty.Text.Trim() != "" ? txeFabricQty.Text.Trim() : "0";
                                string ColorQty = txeColorQty.Text.Trim() != "" ? txeColorQty.Text.Trim() : "0";
                                string TrimQty = txeTrimQty.Text.Trim() != "" ? txeTrimQty.Text.Trim() : "0";
                                string POQty = txePOQty.Text.Trim() != "" ? txePOQty.Text.Trim() : "0";
                                string OrderQtyOld = txeOrderQtyOld.Text.Trim() != "" ? txeOrderQtyOld.Text.Trim() : "0";

                                if (lblStatus.Text == "* Add Forecast")
                                {
                                    sbSQL.Append("  INSERT INTO COForecast(ProductionPlanID, OIDCUST, FileOrderDate, Season, BusinessUnit, OIDCSITEM, ModelNo, OIDVEND, SewingDifficulty, ProductionPlanType, Status, DataUpdate, BookingFabric, BookingAccessory, LastUpdate, ");
                                    sbSQL.Append("      RequestedWHDate, ContractedDate, TransportMethod, LogisticsType, OrderQty, FabricOrderNO, FabricUpdateDate, FabricActualOrderQty, ColorOrderNO, ColorUpdateDate, ColorActualOrderQty, TrimOrderNO, TrimUpdateDate, ");
                                    sbSQL.Append("      TrimActualOrderQty, POOrderNO, POUpdateDate, POActualOrderQty, OrderQTYOld, CreateBy, CreateDate, UpdateBy, Updatedate) ");
                                    sbSQL.Append("  VALUES(");
                                    sbSQL.Append("      N'" + txePlanID.Text.ToUpper().Trim().Replace("'", "''") + "', ");
                                    sbSQL.Append("      '" + slueCustomer.EditValue.ToString().Trim() + "', ");
                                    sbSQL.Append("      " + OrderDate + ", ");
                                    sbSQL.Append("      N'" + speSeason.Value.ToString() + glueSeason.EditValue.ToString() + "', ");
                                    sbSQL.Append("      N'" + glueUnit.Text.Trim() + "', ");
                                    sbSQL.Append("      '" + slueItemCode.EditValue.ToString().Trim() + "', ");
                                    sbSQL.Append("      N'" + txeSampleCode.Text.ToUpper().Trim().Replace("'", "''") + "', ");
                                    sbSQL.Append("      '" + slueSupplier.EditValue.ToString().Trim() + "', ");
                                    sbSQL.Append("      '" + SewingDifficulty + "', ");
                                    sbSQL.Append("      N'" + txePlanType.Text.Trim() + "', ");
                                    sbSQL.Append("      '" + glueStatus.EditValue.ToString().Trim() + "', ");
                                    sbSQL.Append("      " + DataUpdate + ", ");
                                    sbSQL.Append("      '" + chkBookFabric.EditValue.ToString().Trim() + "', ");
                                    sbSQL.Append("      '" + chkBookAcc.EditValue.ToString().Trim() + "', ");
                                    sbSQL.Append("      " + LastUpdate + ", ");
                                    sbSQL.Append("      " + WHDate + ", ");
                                    sbSQL.Append("      " + ContractDate + ", ");
                                    sbSQL.Append("      N'" + glueTransport.EditValue.ToString().Trim() + "', ");
                                    sbSQL.Append("      N'" + glueLogisticsType.Text.Trim() + "', ");
                                    sbSQL.Append("      '" + OrderQty + "', ");
                                    sbSQL.Append("      N'" + txeFabricOrderNo.Text.ToUpper().Trim().Replace("'", "''") + "', ");
                                    sbSQL.Append("      " + FabricUpdate + ", ");
                                    sbSQL.Append("      '" + FabricQty + "', ");
                                    sbSQL.Append("      N'" + txeColorOrderNo.Text.ToUpper().Trim().Replace("'", "''") + "', ");
                                    sbSQL.Append("      " + ColorUpdate + ", ");
                                    sbSQL.Append("      '" + ColorQty + "', ");
                                    sbSQL.Append("      N'" + txeTrimOrderNo.Text.ToUpper().Trim().Replace("'", "''") + "', ");
                                    sbSQL.Append("      " + TrimUpdate + ", ");
                                    sbSQL.Append("      '" + TrimQty + "', ");
                                    sbSQL.Append("      N'" + txePOOrderNo.Text.ToUpper().Trim().Replace("'", "''") + "', ");
                                    sbSQL.Append("      " + POUpdate + ", ");
                                    sbSQL.Append("      '" + POQty + "', ");
                                    sbSQL.Append("      '" + OrderQtyOld + "', ");
                                    sbSQL.Append("      '" + strCREATE + "', GETDATE(), '" + strCREATE + "', GETDATE() ");
                                    sbSQL.Append("  ) ");
                                }
                                else if (lblStatus.Text == "* Edit Forecast")
                                {
                                    sbSQL.Append("  UPDATE COForecast SET ");
                                    sbSQL.Append("      ProductionPlanID=N'" + txePlanID.Text.ToUpper().Trim().Replace("'", "''") + "', ");
                                    sbSQL.Append("      OIDCUST='" + slueCustomer.EditValue.ToString().Trim() + "', ");
                                    sbSQL.Append("      FileOrderDate=" + OrderDate + ", ");
                                    sbSQL.Append("      Season=N'" + speSeason.Value.ToString() + glueSeason.EditValue.ToString() + "', ");
                                    sbSQL.Append("      BusinessUnit=N'" + glueUnit.Text.Trim() + "', ");
                                    sbSQL.Append("      OIDCSITEM='" + slueItemCode.EditValue.ToString().Trim() + "', ");
                                    sbSQL.Append("      ModelNo=N'" + txeSampleCode.Text.ToUpper().Trim().Replace("'", "''") + "', ");
                                    sbSQL.Append("      OIDVEND='" + slueSupplier.EditValue.ToString().Trim() + "', ");
                                    sbSQL.Append("      SewingDifficulty='" + SewingDifficulty + "', ");
                                    sbSQL.Append("      ProductionPlanType=N'" + txePlanType.Text.Trim() + "', ");
                                    sbSQL.Append("      Status='" + glueStatus.EditValue.ToString().Trim() + "', ");
                                    sbSQL.Append("      DataUpdate=" + DataUpdate + ", ");
                                    sbSQL.Append("      BookingFabric='" + chkBookFabric.EditValue.ToString().Trim() + "', ");
                                    sbSQL.Append("      BookingAccessory='" + chkBookAcc.EditValue.ToString().Trim() + "', ");
                                    sbSQL.Append("      LastUpdate=" + LastUpdate + ", ");
                                    sbSQL.Append("      RequestedWHDate=" + WHDate + ", ");
                                    sbSQL.Append("      ContractedDate=" + ContractDate + ", ");
                                    sbSQL.Append("      TransportMethod=N'" + glueTransport.EditValue.ToString().Trim() + "', ");
                                    sbSQL.Append("      LogisticsType=N'" + glueLogisticsType.Text.Trim() + "', ");
                                    sbSQL.Append("      OrderQty='" + OrderQty + "', ");
                                    sbSQL.Append("      FabricOrderNO=N'" + txeFabricOrderNo.Text.ToUpper().Trim().Replace("'", "''") + "', ");
                                    sbSQL.Append("      FabricUpdateDate=" + FabricUpdate + ", ");
                                    sbSQL.Append("      FabricActualOrderQty='" + FabricQty + "', ");
                                    sbSQL.Append("      ColorOrderNO=N'" + txeColorOrderNo.Text.ToUpper().Trim().Replace("'", "''") + "', ");
                                    sbSQL.Append("      ColorUpdateDate=" + ColorUpdate + ", ");
                                    sbSQL.Append("      ColorActualOrderQty='" + ColorQty + "', ");
                                    sbSQL.Append("      TrimOrderNO=N'" + txeTrimOrderNo.Text.ToUpper().Trim().Replace("'", "''") + "', ");
                                    sbSQL.Append("      TrimUpdateDate=" + TrimUpdate + ", ");
                                    sbSQL.Append("      TrimActualOrderQty='" + TrimQty + "', ");
                                    sbSQL.Append("      POOrderNO=N'" + txePOOrderNo.Text.ToUpper().Trim().Replace("'", "''") + "', ");
                                    sbSQL.Append("      POUpdateDate=" + POUpdate + ", ");
                                    sbSQL.Append("      POActualOrderQty='" + POQty + "', ");
                                    sbSQL.Append("      OrderQTYOld='" + OrderQtyOld + "', ");
                                    sbSQL.Append("      UpdateBy='" + strCREATE + "', Updatedate=GETDATE() ");
                                    sbSQL.Append("  WHERE(OIDFC = '" + txeID.Text.Trim() + "')  ");
                                }
                            }
                        }

                        if (lblStatus.Text == "* Add Forecast") //เพิ่มข้อมูลในตารางได้เฉพาะกรณีเพิ่มข้อมูลใหม่เท่านั้น
                        {
                            if (chkPass == true)
                            {
                                //Table Input
                                if (gvINPUT.RowCount > 1)
                                {
                                    DataTable dt = (DataTable)gcINPUT.DataSource;
                                    foreach (DataRow row in dt.Rows)
                                    {
                                        string OrderDate = row["FileOrderDate"].ToString().Trim() != "" ? "'" + Convert.ToDateTime(row["FileOrderDate"].ToString().Trim()).ToString("yyyy-MM-dd") + "'" : "NULL";
                                        string DataUpdate = row["DataUpdate"].ToString().Trim() != "" ? "'" + Convert.ToDateTime(row["DataUpdate"].ToString().Trim()).ToString("yyyy-MM-dd") + "'" : "NULL";
                                        string LastUpdate = row["LastUpdate"].ToString().Trim() != "" ? "'" + Convert.ToDateTime(row["LastUpdate"].ToString().Trim()).ToString("yyyy-MM-dd") + "'" : "NULL";
                                        string WHDate = row["RequestedWHDate"].ToString().Trim() != "" ? "'" + Convert.ToDateTime(row["RequestedWHDate"].ToString().Trim()).ToString("yyyy-MM-dd") + "'" : "NULL";
                                        string ContractDate = row["ContractedDate"].ToString().Trim() != "" ? "'" + Convert.ToDateTime(row["ContractedDate"].ToString().Trim()).ToString("yyyy-MM-dd") + "'" : "NULL";
                                        string FabricUpdate = row["FabricUpdateDate"].ToString().Trim() != "" ? "'" + Convert.ToDateTime(row["FabricUpdateDate"].ToString().Trim()).ToString("yyyy-MM-dd") + "'" : "NULL";
                                        string ColorUpdate = row["ColorUpdateDate"].ToString().Trim() != "" ? "'" + Convert.ToDateTime(row["ColorUpdateDate"].ToString().Trim()).ToString("yyyy-MM-dd") + "'" : "NULL";
                                        string TrimUpdate = row["TrimUpdateDate"].ToString().Trim() != "" ? "'" + Convert.ToDateTime(row["TrimUpdateDate"].ToString().Trim()).ToString("yyyy-MM-dd") + "'" : "NULL";
                                        string POUpdate = row["POUpdateDate"].ToString().Trim() != "" ? "'" + Convert.ToDateTime(row["POUpdateDate"].ToString().Trim()).ToString("yyyy-MM-dd") + "'" : "NULL";

                                        string SewingDifficulty = row["SewingDifficulty"].ToString().Trim() != "" ? row["SewingDifficulty"].ToString().Trim() : "0";
                                        string OrderQty = row["OrderQty"].ToString().Trim() != "" ? row["OrderQty"].ToString().Trim() : "0";
                                        string FabricQty = row["FabricActualOrderQty"].ToString().Trim() != "" ? row["FabricActualOrderQty"].ToString().Trim() : "0";
                                        string ColorQty = row["ColorActualOrderQty"].ToString().Trim() != "" ? row["ColorActualOrderQty"].ToString().Trim() : "0";
                                        string TrimQty = row["TrimActualOrderQty"].ToString().Trim() != "" ? row["TrimActualOrderQty"].ToString().Trim() : "0";
                                        string POQty = row["POActualOrderQty"].ToString().Trim() != "" ? row["POActualOrderQty"].ToString().Trim() : "0";
                                        string OrderQtyOld = row["OrderQTYOld"].ToString().Trim() != "" ? row["OrderQTYOld"].ToString().Trim() : "0";

                                        sbSQL.Append("  INSERT INTO COForecast(ProductionPlanID, OIDCUST, FileOrderDate, Season, BusinessUnit, OIDCSITEM, ModelNo, OIDVEND, SewingDifficulty, ProductionPlanType, Status, DataUpdate, BookingFabric, BookingAccessory, LastUpdate, ");
                                        sbSQL.Append("      RequestedWHDate, ContractedDate, TransportMethod, LogisticsType, OrderQty, FabricOrderNO, FabricUpdateDate, FabricActualOrderQty, ColorOrderNO, ColorUpdateDate, ColorActualOrderQty, TrimOrderNO, TrimUpdateDate, ");
                                        sbSQL.Append("      TrimActualOrderQty, POOrderNO, POUpdateDate, POActualOrderQty, OrderQTYOld, CreateBy, CreateDate, UpdateBy, Updatedate) ");
                                        sbSQL.Append("  VALUES(");
                                        sbSQL.Append("      N'" + row["ProductionPlanID"].ToString().ToUpper().Trim().Replace("'", "''") + "', ");
                                        sbSQL.Append("      '" + row["OIDCUST"].ToString().Trim() + "', ");
                                        sbSQL.Append("      " + OrderDate + ", ");
                                        sbSQL.Append("      N'" + row["Year"].ToString().Trim() + row["Season"].ToString().Trim() + "', ");
                                        sbSQL.Append("      N'" + row["BusinessUnit"].ToString().Trim() + "', ");
                                        sbSQL.Append("      '" + row["OIDCSITEM"].ToString().Trim() + "', ");
                                        sbSQL.Append("      N'" + row["ModelNo"].ToString().ToUpper().Trim().Replace("'", "''") + "', ");
                                        sbSQL.Append("      '" + row["OIDVEND"].ToString().Trim() + "', ");
                                        sbSQL.Append("      '" + SewingDifficulty + "', ");
                                        sbSQL.Append("      N'" + row["ProductionPlanType"].ToString().Trim() + "', ");
                                        sbSQL.Append("      '" + row["Status"].ToString().Trim() + "', ");
                                        sbSQL.Append("      " + DataUpdate + ", ");
                                        sbSQL.Append("      '" + row["BookingFabric"].ToString().Trim() + "', ");
                                        sbSQL.Append("      '" + row["BookingAccessory"].ToString().Trim() + "', ");
                                        sbSQL.Append("      " + LastUpdate + ", ");
                                        sbSQL.Append("      " + WHDate + ", ");
                                        sbSQL.Append("      " + ContractDate + ", ");
                                        sbSQL.Append("      N'" + row["TransportMethod"].ToString().Trim() + "', ");
                                        sbSQL.Append("      N'" + row["LogisticsType"].ToString().Trim() + "', ");
                                        sbSQL.Append("      '" + OrderQty + "', ");
                                        sbSQL.Append("      N'" + row["FabricOrderNO"].ToString().ToUpper().Trim().Replace("'", "''") + "', ");
                                        sbSQL.Append("      " + FabricUpdate + ", ");
                                        sbSQL.Append("      '" + FabricQty + "', ");
                                        sbSQL.Append("      N'" + row["ColorOrderNO"].ToString().ToUpper().Trim().Replace("'", "''") + "', ");
                                        sbSQL.Append("      " + ColorUpdate + ", ");
                                        sbSQL.Append("      '" + ColorQty + "', ");
                                        sbSQL.Append("      N'" + row["TrimOrderNO"].ToString().ToUpper().Trim().Replace("'", "''") + "', ");
                                        sbSQL.Append("      " + TrimUpdate + ", ");
                                        sbSQL.Append("      '" + TrimQty + "', ");
                                        sbSQL.Append("      N'" + row["POOrderNO"].ToString().ToUpper().Trim().Replace("'", "''") + "', ");
                                        sbSQL.Append("      " + POUpdate + ", ");
                                        sbSQL.Append("      '" + POQty + "', ");
                                        sbSQL.Append("      '" + OrderQtyOld + "', ");
                                        sbSQL.Append("      '" + strCREATE + "', GETDATE(), '" + strCREATE + "', GETDATE() ");
                                        sbSQL.Append("  ) ");
                                    }
                                }
                            }
                        }

                        if (sbSQL.Length > 0)
                        {
                            if (FUNC.msgQuiz("Confirm save data ?") == true)
                            {
                                //MessageBox.Show(sbSQL.ToString());
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
            else if (tabbedControlGroup1.SelectedTabPage == layoutControlGroup7) //Import File
            {
                if (txeFilePath.Text.Trim() == "")
                {
                    FUNC.msgWarning("Please select excel file.");
                    txeFilePath.Focus();
                }
                else if (cbeSheet.Text.Trim() == "")
                {
                    FUNC.msgWarning("Please select excel sheet.");
                    cbeSheet.Focus();
                }
                else
                {
                    if (FUNC.msgQuiz("Confirm save excel file import data ?") == true)
                    {
                        StringBuilder sbSQL = new StringBuilder();
                        string strCREATE = "0";
                        //if (txeCREATE.Text.Trim() != "")
                        //{
                        //    strCREATE = txeCREATE.Text.Trim();
                        //}
                        bool chkSAVE = false;

                        IWorkbook workbook = spsForecast.Document;
                        Worksheet WSHEET = workbook.Worksheets[0];

                        lciPregressSave.Visibility = LayoutVisibility.Always;
                        pbcSave.Properties.Step = 1;
                        pbcSave.Properties.PercentView = true;
                        pbcSave.Properties.Maximum = WSHEET.GetDataRange().RowCount;
                        pbcSave.Properties.Minimum = 0;
                        pbcSave.EditValue = 0;

                        string Customer = "";
                        string OIDCUST = "";

                        for (int i = 1; i < WSHEET.GetDataRange().RowCount; i++)
                        {
                            string PlanID = WSHEET.Rows[i][0].DisplayText.ToString().Trim();
            
                            if (PlanID != "")
                            {
                                string strCustomer = WSHEET.Rows[i][4].DisplayText.ToString().Trim().Replace("'", "''");
                                if (Customer != strCustomer.Replace(" ", "").Replace(".", "").Replace(",", ""))
                                {
                                    Customer = strCustomer.Replace(" ", "").Replace(".", "").Replace(",", "");
                                    string CustomerCode = Customer.Length > 20 ? Customer.Substring(0, 20) : Customer;
                                    string CustomerShort = Customer.Length > 10 ? Customer.Substring(0, 10) : Customer;
                                    StringBuilder sbCUST = new StringBuilder();
                                    sbCUST.Append("IF NOT EXISTS(SELECT OIDCUST FROM Customer WHERE Name LIKE N'%" + Customer + "%') ");
                                    sbCUST.Append(" BEGIN ");
                                    sbCUST.Append("   INSERT INTO Customer(Code, Name, ShortName) VALUES(N'" + CustomerCode + "', N'" + strCustomer + "', N'" + CustomerShort + "') ");
                                    sbCUST.Append(" END ");
                                    sbCUST.Append("SELECT TOP(1) OIDCUST FROM Customer WHERE Name LIKE N'%" + Customer + "%' ");
                                    OIDCUST = this.DBC.DBQuery(sbCUST).getString();
                                }

                                string ItemCode = WSHEET.Rows[i][6].DisplayText.ToString().Trim().Replace("'", "''");
                                ItemCode = ItemCode.Length > 20 ? ItemCode.Substring(0, 20) : ItemCode;
                                string ItemName = WSHEET.Rows[i][7].DisplayText.ToString().Trim().Replace("'", "''");
                                string StyleNo = WSHEET.Rows[i][8].DisplayText.ToString().Trim().Replace("'", "''");
                                StyleNo = StyleNo.Length > 10 ? StyleNo.Substring(0, 10) : StyleNo;

                                string strStyle = StyleNo.Replace(Regex.Match(StyleNo, @"\d+([,\.]\d+)?").Value, ""); //5 ตัวสุดท้าย
                                StringBuilder sbSTYLE = new StringBuilder();
                                sbSTYLE.Append("IF NOT EXISTS(SELECT OIDSTYLE FROM ProductStyle WHERE StyleName = N'" + strStyle + "') ");
                                sbSTYLE.Append("  BEGIN ");
                                sbSTYLE.Append("       INSERT INTO ProductStyle(StyleName) VALUES(N'" + strStyle + "') ");
                                sbSTYLE.Append("  END ");
                                sbSTYLE.Append("SELECT OIDSTYLE FROM ProductStyle WHERE(StyleName = N'" + strStyle + "') ");
                                string OIDSTYLE = this.DBC.DBQuery(sbSTYLE).getString();

                                string Season = WSHEET.Rows[i][1].DisplayText.ToString().Trim() + WSHEET.Rows[i][2].DisplayText.ToString().Trim();
                                string Supplier = WSHEET.Rows[i][19].DisplayText.ToString().Trim().Replace("'", "''"); //Raw Material Supplier Code

                                string Unit = WSHEET.Rows[i][5].DisplayText.ToString().Trim().Replace("'", "''");
                                string ModelNo = WSHEET.Rows[i][8].DisplayText.ToString().Trim().Replace("'", "''");
                                string SewingDifficulty = WSHEET.Rows[i][20].DisplayText.ToString().Trim();
                                SewingDifficulty = SewingDifficulty == "" ? "0" : SewingDifficulty;

                                string PlanType = WSHEET.Rows[i][25].DisplayText.ToString().Trim().Replace("'", "''");
                                string Status = WSHEET.Rows[i][76].DisplayText.ToString().Trim();
                                Status = Status == "" ? "0" : Status;

                                string LastUpdate = WSHEET.Rows[i][27].Value.ToString().Trim();
                                LastUpdate = LastUpdate != "" ? "'" + Convert.ToDateTime(LastUpdate).ToString("yyyy-MM-dd") + "'" : "NULL";
                                string WHDate = WSHEET.Rows[i][28].Value.ToString().Trim();
                                WHDate = WHDate != "" ? "'" + Convert.ToDateTime(WHDate).ToString("yyyy-MM-dd") + "'" : "NULL";
                                string ContractedDate = WSHEET.Rows[i][29].Value.ToString().Trim();
                                ContractedDate = ContractedDate != "" ? "'" + Convert.ToDateTime(ContractedDate).ToString("yyyy-MM-dd") + "'" : "NULL";
                                string TransportMethod = WSHEET.Rows[i][30].DisplayText.ToString().Trim().Replace("'", "''");
                                string LogisticsType = WSHEET.Rows[i][31].DisplayText.ToString().Trim().Replace("'", "''");
                                string OrderQty = WSHEET.Rows[i][32].DisplayText.ToString().Trim();
                                OrderQty = OrderQty == "" ? "0" : OrderQty;

                                string FabricOrderNo = WSHEET.Rows[i][44].DisplayText.ToString().Trim().Replace("'", "''");
                                string FabricUpdateDate = WSHEET.Rows[i][46].Value.ToString().Trim();
                                FabricUpdateDate = FabricUpdateDate != "" ? "'" + Convert.ToDateTime(FabricUpdateDate).ToString("yyyy-MM-dd") + "'" : "NULL";
                                string FabricOrderQty = WSHEET.Rows[i][45].DisplayText.ToString().Trim();
                                FabricOrderQty = FabricOrderQty == "" ? "0" : FabricOrderQty;

                                string TrimOrderNo = WSHEET.Rows[i][58].DisplayText.ToString().Trim().Replace("'", "''");
                                string TrimUpdateDate = WSHEET.Rows[i][60].Value.ToString().Trim();
                                TrimUpdateDate = TrimUpdateDate != "" ? "'" + Convert.ToDateTime(TrimUpdateDate).ToString("yyyy-MM-dd") + "'" : "NULL";
                                string TrimOrderQty = WSHEET.Rows[i][59].DisplayText.ToString().Trim();
                                TrimOrderQty = TrimOrderQty == "" ? "0" : TrimOrderQty;

                                string POOrderNo = WSHEET.Rows[i][65].DisplayText.ToString().Trim().Replace("'", "''");
                                string POUpdateDate = WSHEET.Rows[i][67].Value.ToString().Trim();
                                POUpdateDate = POUpdateDate != "" ? "'" + Convert.ToDateTime(POUpdateDate).ToString("yyyy-MM-dd") + "'" : "NULL";
                                string POOrderQty = WSHEET.Rows[i][66].DisplayText.ToString().Trim();
                                POOrderQty = POOrderQty == "" ? "0" : POOrderQty;

                                string ColorOrderNo = WSHEET.Rows[i][51].DisplayText.ToString().Trim().Replace("'", "''");
                                string ColorUpdateDate = WSHEET.Rows[i][53].Value.ToString().Trim();
                                ColorUpdateDate = ColorUpdateDate != "" ? "'" + Convert.ToDateTime(ColorUpdateDate).ToString("yyyy-MM-dd") + "'" : "NULL";
                                string ColorOrderQty = WSHEET.Rows[i][52].DisplayText.ToString().Trim();
                                ColorOrderQty = ColorOrderQty == "" ? "0" : ColorOrderQty;

                                string OrderQtyOld = WSHEET.Rows[i][68].DisplayText.ToString().Trim();
                                OrderQtyOld = OrderQtyOld == "" ? "0" : OrderQtyOld;

                                string CustCode = Customer.ToUpper().Trim().Replace(" ", "").Replace("'", "");
                                if (CustCode.Length > 20)
                                    CustCode = CustCode.Substring(0, 20);

                                string CustShort = Customer.ToUpper().Trim().Replace(" ", "").Replace("'", "");
                                if (CustShort.Length > 10)
                                    CustShort = CustShort.Substring(0, 10);

                                string SupplierCode = Supplier.ToUpper().Trim().Replace(" ", "").Replace("'", "");
                                if (SupplierCode.Length > 20)
                                    SupplierCode = SupplierCode.Substring(0, 20);

                                sbSQL.Clear();
                                //sbSQL.Append("IF NOT EXISTS(SELECT OIDCUST FROM Customer WHERE Name LIKE N'" + Customer + "%') ");
                                //sbSQL.Append(" BEGIN ");
                                //sbSQL.Append("   INSERT INTO Customer(Code, Name, ShortName) VALUES(N'" + CustCode + "', N'" + Customer + "', N'" + CustShort + "') ");
                                //sbSQL.Append(" END ");

                                //sbSQL.Append("IF NOT EXISTS(SELECT OIDCSITEM FROM ItemCustomer WHERE ItemCode = N'" + ItemCode + "' AND OIDCUST = '" + OIDCUST + "') ");
                                //sbSQL.Append(" BEGIN ");
                                //sbSQL.Append("   INSERT INTO ItemCustomer(OIDCUST, ItemCode, ItemName, StyleNo, OIDSTYLE, Season) ");
                                //sbSQL.Append("   SELECT (SELECT TOP(1) OIDCUST FROM Customer WHERE Name LIKE N'" + Customer + "%') AS OIDCUST, N'" + ItemCode + "' AS ItemCode, N'" + ItemName + "' AS ItemName, N'" + StyleNo + "' AS StyleNo, '" + OIDSTYLE + "' AS OIDSTYLE, N'" + Season + "' AS Season ");
                                //sbSQL.Append(" END ");

                                sbSQL.Append("IF NOT EXISTS(SELECT OIDCSITEM FROM ItemCustomer WHERE (OIDCUST='" + OIDCUST + "') AND (ItemCode = N'" + ItemCode + "')) ");
                                sbSQL.Append(" BEGIN ");
                                sbSQL.Append("   INSERT INTO ItemCustomer(OIDCUST, ItemCode, ItemName, OIDSTYLE, Season, StyleNo) VALUES('" + OIDCUST + "', N'" + ItemCode + "', N'" + ItemName + "', '" + OIDSTYLE + "', N'" + Season + "', N'" + StyleNo + "') ");
                                sbSQL.Append(" END ");
                                sbSQL.Append("ELSE ");
                                sbSQL.Append(" BEGIN ");
                                sbSQL.Append("   UPDATE ItemCustomer SET  ");
                                sbSQL.Append("     ItemName=N'" + ItemName + "', OIDSTYLE='" + OIDSTYLE + "', Season=N'" + Season + "', StyleNo=N'" + StyleNo + "'  ");
                                sbSQL.Append("   WHERE (OIDCUST='" + OIDCUST + "') AND (ItemCode = N'" + ItemCode + "')  ");
                                sbSQL.Append(" END ");
                                

                                sbSQL.Append("IF NOT EXISTS(SELECT OIDVEND FROM Vendor WHERE Name LIKE N'" + Supplier + "%' AND VendorType = 6) ");
                                sbSQL.Append(" BEGIN ");
                                sbSQL.Append("   INSERT INTO Vendor(Code, Name, VendorType) VALUES(N'" + SupplierCode + "', N'" + Supplier + "', '6') ");
                                sbSQL.Append(" END ");

                                sbSQL.Append("IF NOT EXISTS(SELECT OIDFC FROM COForecast WHERE ProductionPlanID = N'" + PlanID + "') ");
                                sbSQL.Append(" BEGIN ");
                                sbSQL.Append("   INSERT INTO COForecast(ProductionPlanID, OIDCUST, Season, BusinessUnit, OIDCSITEM, ModelNo, OIDVEND, SewingDifficulty, ProductionPlanType, Status, LastUpdate, RequestedWHDate, ContractedDate, TransportMethod, ");
                                sbSQL.Append("                         LogisticsType, OrderQty, FabricOrderNO, FabricUpdateDate, FabricActualOrderQty, TrimOrderNO, TrimUpdateDate, TrimActualOrderQty, POOrderNO, POUpdateDate, POActualOrderQty, ColorOrderNO, ColorUpdateDate, ");
                                sbSQL.Append("                         ColorActualOrderQty, OrderQTYOld, BookingFabric, BookingAccessory, FileOrderDate, DataUpdate, CreateBy, CreateDate, UpdateBy, Updatedate) ");
                                sbSQL.Append("   SELECT N'" + PlanID + "' AS ProductionPlanID, ");
                                sbSQL.Append("      '" + OIDCUST + "' AS OIDCUST, ");
                                sbSQL.Append("      N'" + Season + "' AS Season, ");
                                sbSQL.Append("      N'" + Unit + "' AS BusinessUnit, ");
                                sbSQL.Append("      (SELECT TOP(1) OIDCSITEM FROM ItemCustomer WHERE ItemCode = N'" + ItemCode + "' AND OIDCUST='" + OIDCUST + "') AS OIDCSITEM, ");
                                sbSQL.Append("      N'" + ModelNo + "' AS ModelNo, ");
                                sbSQL.Append("      (SELECT TOP(1) OIDVEND FROM Vendor WHERE Name LIKE N'" + Supplier + "%' AND VendorType = 6) AS OIDVEND, ");
                                sbSQL.Append("      '" + SewingDifficulty + "' AS SewingDifficulty, ");
                                sbSQL.Append("      N'" + PlanType + "' AS ProductionPlanType, ");
                                sbSQL.Append("      '" + Status + "' AS Status, ");
                                sbSQL.Append("      " + LastUpdate + " AS LastUpdate, ");
                                sbSQL.Append("      " + WHDate + " AS RequestedWHDate, ");
                                sbSQL.Append("      " + ContractedDate + " AS ContractedDate, ");
                                sbSQL.Append("      N'" + TransportMethod + "' AS TransportMethod, ");
                                sbSQL.Append("      N'" + LogisticsType + "' AS LogisticsType, ");
                                sbSQL.Append("      '" + OrderQty + "' AS OrderQty, ");
                                sbSQL.Append("      N'" + FabricOrderNo + "' AS FabricOrderNO, ");
                                sbSQL.Append("      " + FabricUpdateDate + " AS FabricUpdateDate, ");
                                sbSQL.Append("      '" + FabricOrderQty + "' AS FabricActualOrderQty, ");
                                sbSQL.Append("      N'" + TrimOrderNo + "' AS TrimOrderNO, ");
                                sbSQL.Append("      " + TrimUpdateDate + " AS TrimUpdateDate, ");
                                sbSQL.Append("      '" + TrimOrderQty + "' AS TrimActualOrderQty, ");
                                sbSQL.Append("      N'" + POOrderNo + "' AS POOrderNO, ");
                                sbSQL.Append("      " + POUpdateDate + " AS POUpdateDate, ");
                                sbSQL.Append("      '" + POOrderQty + "' AS POActualOrderQty, ");
                                sbSQL.Append("      N'" + ColorOrderNo + "' AS ColorOrderNO, ");
                                sbSQL.Append("      " + ColorUpdateDate + " AS ColorUpdateDate, ");
                                sbSQL.Append("      '" + ColorOrderQty + "' AS ColorActualOrderQty, ");
                                sbSQL.Append("      '" + OrderQtyOld + "' AS OrderQTYOld, ");
                                sbSQL.Append("      '0' AS BookingFabric, ");
                                sbSQL.Append("      '0' AS BookingAccessory, ");
                                sbSQL.Append("      GETDATE() AS FileOrderDate, ");
                                sbSQL.Append("      GETDATE() AS DataUpdate, ");
                                sbSQL.Append("      '" + strCREATE + "' AS CreateBy, ");
                                sbSQL.Append("      GETDATE() AS CreateDate, ");
                                sbSQL.Append("      '" + strCREATE + "' AS UpdateBy, ");
                                sbSQL.Append("      GETDATE() AS Updatedate ");
                                sbSQL.Append(" END ");
                                sbSQL.Append("ELSE ");
                                sbSQL.Append(" BEGIN ");
                                sbSQL.Append("   UPDATE COForecast SET ");
                                sbSQL.Append("      OIDCUST = '" + OIDCUST + "', ");
                                sbSQL.Append("      Season = N'" + Season + "', ");
                                sbSQL.Append("      BusinessUnit = N'" + Unit + "', ");
                                sbSQL.Append("      OIDCSITEM = (SELECT TOP(1) OIDCSITEM FROM ItemCustomer WHERE ItemCode = N'" + ItemCode + "' AND OIDCUST='" + OIDCUST + "'), ");
                                sbSQL.Append("      ModelNo = N'" + ModelNo + "', ");
                                sbSQL.Append("      OIDVEND = (SELECT TOP(1) OIDVEND FROM Vendor WHERE Name LIKE N'" + Supplier + "%' AND VendorType = 6), ");
                                sbSQL.Append("      SewingDifficulty = '" + SewingDifficulty + "', ");
                                sbSQL.Append("      ProductionPlanType = N'" + PlanType + "', ");
                                sbSQL.Append("      Status = '" + Status + "', ");
                                sbSQL.Append("      LastUpdate = " + LastUpdate + ", ");
                                sbSQL.Append("      RequestedWHDate = " + WHDate + ", ");
                                sbSQL.Append("      ContractedDate = " + ContractedDate + ", ");
                                sbSQL.Append("      TransportMethod = N'" + TransportMethod + "', ");
                                sbSQL.Append("      LogisticsType = N'" + LogisticsType + "', ");
                                sbSQL.Append("      OrderQty = '" + OrderQty + "', ");
                                sbSQL.Append("      FabricOrderNO = N'" + FabricOrderNo + "', ");
                                sbSQL.Append("      FabricUpdateDate = " + FabricUpdateDate + ", ");
                                sbSQL.Append("      FabricActualOrderQty = '" + FabricOrderQty + "', ");
                                sbSQL.Append("      TrimOrderNO = N'" + TrimOrderNo + "', ");
                                sbSQL.Append("      TrimUpdateDate = " + TrimUpdateDate + ", ");
                                sbSQL.Append("      TrimActualOrderQty = '" + TrimOrderQty + "', ");
                                sbSQL.Append("      POOrderNO = N'" + POOrderNo + "', ");
                                sbSQL.Append("      POUpdateDate = " + POUpdateDate + ", ");
                                sbSQL.Append("      POActualOrderQty = '" + POOrderQty + "', ");
                                sbSQL.Append("      ColorOrderNO = N'" + ColorOrderNo + "', ");
                                sbSQL.Append("      ColorUpdateDate = " + ColorUpdateDate + ", ");
                                sbSQL.Append("      ColorActualOrderQty = '" + ColorOrderQty + "', ");
                                sbSQL.Append("      OrderQTYOld = '" + OrderQtyOld + "', ");
                                sbSQL.Append("      UpdateBy = '" + strCREATE + "', ");
                                sbSQL.Append("      Updatedate = GETDATE() ");
                                sbSQL.Append("    WHERE(ProductionPlanID = N'" + PlanID + "') ");
                                sbSQL.Append(" END   ");

                                //memoEdit1.EditValue = sbSQL.ToString();
                                //break;
                                //MessageBox.Show(sbSQL.ToString());
                                try
                                {
                                    chkSAVE = this.DBC.DBQuery(sbSQL).runSQL();
                                    if (chkSAVE == false)
                                    {
                                        break;
                                    }
                                    else
                                    {
                                        pbcSave.PerformStep();
                                        pbcSave.Update();
                                    }
                                }
                                catch (Exception)
                                { }

                            }

                        }

                        if (chkSAVE == true)
                        {
                            lciPregressSave.Visibility = LayoutVisibility.Never;
                            FUNC.msgInfo("Save complete.");
                            bbiNew.PerformClick();
                        }

                        //if (sbSQL.Length > 0)
                        //{
                        //    //MessageBox.Show(sbSQL.ToString());
                        //    try
                        //    {
                        //        chkSAVE = this.DBC.DBQuery(sbSQL).runSQL();
                        //        if (chkSAVE == true)
                        //        {
                        //            FUNC.msgInfo("Save complete.");
                        //            bbiNew.PerformClick();
                        //        }
                        //    }
                        //    catch (Exception)
                        //    { }
                        //}

                    }
                }
            }
            
        }

        private void bbiExcel_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            string pathFile = new ObjSet.Folder(@"C:\MDS\Export\").GetPath() + "COForecastList_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
            gvFO.ExportToXlsx(pathFile);
            System.Diagnostics.Process.Start(pathFile);
        }


        private void bbiPrintPreview_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            gcFO.ShowPrintPreview();
        }

        private void bbiPrint_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            gcFO.Print();
        }

        private void txeCode_KeyDown(object sender, KeyEventArgs e)
        {
            //if (e.KeyCode == Keys.Enter)
            //{
            //    txeName.Focus();
            //}
        }

        private void txeCode_Leave(object sender, EventArgs e)
        {
            //if (txeCode.Text.Trim() != "")
            //{
            //    txeCode.Text = txeCode.Text.ToUpper().Trim();
            //    bool chkDup = chkDuplicateNo();
            //    if (chkDup == true)
            //    {
            //        txeName.Focus();
            //    }
            //    else
            //    {
            //        txeCode.Text = "";
            //        txeCode.Focus();
            //        //FUNC.msgWarning("Duplicate code. !! Please Change.");

            //    }
            //}
        }

        private void txeName_KeyDown(object sender, KeyEventArgs e)
        {
            //if (e.KeyCode == Keys.Enter)
            //{
            //    txeCity.Focus();
            //}
        }

        private void txeName_Leave(object sender, EventArgs e)
        {
            //if (txeName.Text.Trim() != "")
            //{
            //    txeName.Text = txeName.Text.ToUpper().Trim();
            //    bool chkDup = chkDuplicateName();
            //    if (chkDup == true)
            //    {
            //        txeCity.Focus();
            //    }
            //    else
            //    {
            //        txeName.Text = "";
            //        txeName.Focus();
            //        //FUNC.msgWarning("Duplicate name. !! Please Change.");

            //    }
            //}
        }

        private void txeCity_KeyDown(object sender, KeyEventArgs e)
        {
            //if (e.KeyCode == Keys.Enter)
            //{
            //    slueCountry.Focus();
            //}
        }

        private void txeCity_Leave(object sender, EventArgs e)
        {

        }

        //private bool chkDuplicateNo()
        //{
        //    bool chkDup = true;
        //    if (txeCode.Text != "")
        //    {
        //        txeCode.Text = txeCode.Text.ToUpper().Trim();
        //        if (txeCode.Text.Trim() != "" && lblStatus.Text == "* Add Port")
        //        {
        //            StringBuilder sbSQL = new StringBuilder();
        //            sbSQL.Append("SELECT TOP(1) PortCode FROM PortAndCity WHERE (PortCode = N'" + txeCode.Text.Trim().Trim().Replace("'", "''") + "') ");
        //            if (this.DBC.DBQuery(sbSQL).getString() != "")
        //            {
        //                txeCode.Text = "";
        //                txeCode.Focus();
        //                chkDup = false;
        //                FUNC.msgWarning("Duplicate code. !! Please Change.");
        //            }
        //        }
        //        else if (txeCode.Text.Trim() != "" && lblStatus.Text == "* Edit Port")
        //        {
        //            StringBuilder sbSQL = new StringBuilder();
        //            sbSQL.Append("SELECT TOP(1) OIDPORT ");
        //            sbSQL.Append("FROM PortAndCity ");
        //            sbSQL.Append("WHERE (PortCode = N'" + txeCode.Text.Trim().Trim().Replace("'", "''") + "') ");
        //            string strCHK = this.DBC.DBQuery(sbSQL).getString();
        //            if (strCHK != "" && strCHK != txeID.Text.Trim())
        //            {
        //                txeCode.Text = "";
        //                txeCode.Focus();
        //                chkDup = false;
        //                FUNC.msgWarning("Duplicate code. !! Please Change.");
        //            }
        //        }
        //    }
        //    return chkDup;
        //}

        //private bool chkDuplicateName()
        //{
        //    bool chkDup = true;
        //    if (txeName.Text != "")
        //    {
        //        txeName.Text = txeName.Text.ToUpper().Trim();
        //        if (txeName.Text.Trim() != "" && lblStatus.Text == "* Add Port")
        //        {
        //            StringBuilder sbSQL = new StringBuilder();
        //            sbSQL.Append("SELECT TOP(1) PortName FROM PortAndCity WHERE (PortName = N'" + txeName.Text.Trim().Replace("'", "''") + "') ");
        //            if (this.DBC.DBQuery(sbSQL).getString() != "")
        //            {
        //                txeName.Text = "";
        //                txeName.Focus();
        //                chkDup = false;
        //                FUNC.msgWarning("Duplicate name. !! Please Change.");
        //            }
        //        }
        //        else if (txeName.Text.Trim() != "" && lblStatus.Text == "* Edit Port")
        //        {
        //            StringBuilder sbSQL = new StringBuilder();
        //            sbSQL.Append("SELECT TOP(1) OIDPORT ");
        //            sbSQL.Append("FROM PortAndCity ");
        //            sbSQL.Append("WHERE (PortName = N'" + txeName.Text.Trim().Replace("'", "''") + "') ");
        //            string strCHK = this.DBC.DBQuery(sbSQL).getString();
        //            if (strCHK != "" && strCHK != txeID.Text.Trim())
        //            {
        //                txeName.Text = "";
        //                txeName.Focus();
        //                chkDup = false;
        //                FUNC.msgWarning("Duplicate name. !! Please Change.");
        //            }
        //        }
        //    }
        //    return chkDup;
        //}

        //*********** REGION ************
        public class LocalesRetrievalException : Exception
        {
            public LocalesRetrievalException(string message)
                : base(message)
            {
            }
        }

        #region Windows API

        private delegate bool EnumLocalesProcExDelegate(
           [MarshalAs(UnmanagedType.LPWStr)] String lpLocaleString,
           LocaleType dwFlags, int lParam);

        [DllImport(@"kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern bool EnumSystemLocalesEx(EnumLocalesProcExDelegate pEnumProcEx,
           LocaleType dwFlags, int lParam, IntPtr lpReserved);

        private enum LocaleType : uint
        {
            LocaleAll = 0x00000000,             // Enumerate all named based locales
            LocaleWindows = 0x00000001,         // Shipped locales and/or replacements for them
            LocaleSupplemental = 0x00000002,    // Supplemental locales only
            LocaleAlternateSorts = 0x00000004,  // Alternate sort locales
            LocaleNeutralData = 0x00000010,     // Locales that are "neutral" (language only, region data is default)
            LocaleSpecificData = 0x00000020,    // Locales that contain language and region data
        }

        #endregion

        public enum CultureTypes : uint
        {
            SpecificCultures = LocaleType.LocaleSpecificData,
            NeutralCultures = LocaleType.LocaleNeutralData,
            AllCultures = LocaleType.LocaleWindows
        }

        public static List<CultureInfo> GetCultures(CultureTypes cultureTypes)
        {
            List<CultureInfo> cultures = new List<CultureInfo>();
            EnumLocalesProcExDelegate enumCallback = (locale, flags, lParam) =>
            {
                try
                {
                    cultures.Add(new CultureInfo(locale));
                }
                catch (CultureNotFoundException)
                {
                    // This culture is not supported by .NET (not happened so far)
                    // Must be ignored.
                }
                return true;
            };

            if (EnumSystemLocalesEx(enumCallback, (LocaleType)cultureTypes, 0, (IntPtr)0) == false)
            {
                int errorCode = Marshal.GetLastWin32Error();
                throw new LocalesRetrievalException("Win32 error " + errorCode + " while trying to get the Windows locales");
            }

            // Add the two neutral cultures that Windows misses 
            // (CultureInfo.GetCultures adds them also):
            if (cultureTypes == CultureTypes.NeutralCultures || cultureTypes == CultureTypes.AllCultures)
            {
                cultures.Add(new CultureInfo("en-US"));
                //cultures.Add(new CultureInfo("zh-CHS"));
                //cultures.Add(new CultureInfo("zh-CHT"));
            }

            return cultures;
        }

        public static List<string> GetCountries()
        {
            List<CultureInfo> cultures = GetCultures(CultureTypes.SpecificCultures);
            List<string> countries = new List<string>();
           
            foreach (CultureInfo culture in cultures)
            {
                RegionInfo region = new RegionInfo(culture.Name);

                if (!(countries.Contains(region.EnglishName)))
                {
                    countries.Add(region.EnglishName);
                }
            }
            countries.Sort();
            return countries;
        }


        //*********** END-REGION ********

        private void slueCountry_Popup(object sender, EventArgs e)
        {
            //(sender as SearchLookUpEdit).Properties.View.Columns["Country"].SortOrder = DevExpress.Data.ColumnSortOrder.Ascending;
        }

        private void gvPort_DoubleClick(object sender, EventArgs e)
        {
        //    GridView view = (GridView)sender;
        //    Point pt = view.GridControl.PointToClient(Control.MousePosition);
        //    DevExpress.XtraGrid.Views.Grid.ViewInfo.GridHitInfo info = view.CalcHitInfo(pt);
        //    if (info.InRow || info.InRowCell)
        //    {
        //        DataTable dtCP = (DataTable)gcPort.DataSource;
        //        if (dtCP.Rows.Count > 0)
        //        {
        //            lblStatus.Text = "* Edit Port";
        //            lblStatus.ForeColor = Color.Red;

        //            DataRow drCP = dtCP.Rows[info.RowHandle];
        //            txeID.Text = drCP["OIDPORT"].ToString();
        //            txeCode.Text = drCP["PortCode"].ToString();
        //            txeName.Text = drCP["PortName"].ToString();
        //            txeCity.Text = drCP["City"].ToString();
        //            slueCountry.EditValue = drCP["Country"].ToString();

        //            txeCREATE.Text = drCP["CreatedBy"].ToString();
        //            txeDATE.Text = drCP["CreatedDate"].ToString();
        //        }
        //    }
        }

        private void tabbedControlGroup1_SelectedPageChanged(object sender, DevExpress.XtraLayout.LayoutTabPageChangedEventArgs e)
        {
            if (tabbedControlGroup1.SelectedTabPage == layoutControlGroup1)
            {
                bbiRefresh.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
                bbiNew.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
                bbiSave.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
                bbiEdit.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
                bbiClone.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;

                ribbonPageGroup2.Visible = true;
                ribbonPageGroup5.Visible = true;
            }
            else if (tabbedControlGroup1.SelectedTabPage == layoutControlGroup2)
            {
                bbiRefresh.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
                bbiNew.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
                bbiSave.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
                bbiEdit.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
                bbiClone.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;

                ribbonPageGroup2.Visible = false;
                ribbonPageGroup5.Visible = false;
            }
            else if (tabbedControlGroup1.SelectedTabPage == layoutControlGroup7)
            {
                bbiRefresh.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
                bbiNew.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
                bbiSave.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
                bbiEdit.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
                bbiClone.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;

                ribbonPageGroup2.Visible = false;
                ribbonPageGroup5.Visible = false;
            }
        }

        private void bbiRefresh_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            LoadSummary();
        }

        private void slueItemCode_EditValueChanged(object sender, EventArgs e)
        {
            txeItemName.Text = "";
            txeStyle.Text = "";
            if (slueItemCode.Text.Trim() != "")
            {
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT ItemName, StyleNo ");
                sbSQL.Append("FROM ItemCustomer ");
                sbSQL.Append("WHERE (OIDCSITEM = N'" + slueItemCode.EditValue.ToString() + "') ");
                string[] arrITEM = this.DBC.DBQuery(sbSQL).getMultipleValue();
                if (arrITEM.Length > 0)
                {
                    txeItemName.Text = arrITEM[0];
                    txeStyle.Text = arrITEM[1];
                }
            }

            txeSampleCode.Focus();
        }

        private void gvINPUT_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            if (e.Column.FieldName == "OIDCSITEM")
            {
                if (gvINPUT.GetRowCellValue(e.RowHandle, "OIDCSITEM") != null)
                {
                    string strITEM = gvINPUT.GetRowCellValue(e.RowHandle, "OIDCSITEM").ToString();
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT ITC.ItemCode, ITC.ItemName, CUS.Name AS Customer, ITC.StyleNo AS [StyleNo.], ITC.Season, ITC.OIDCSITEM AS ID ");
                    sbSQL.Append("FROM   ItemCustomer AS ITC LEFT OUTER JOIN ");
                    sbSQL.Append("       Customer AS CUS ON ITC.OIDCUST = CUS.OIDCUST ");
                    sbSQL.Append("WHERE (ITC.ItemCode = '" + strITEM + "') ");
                    string[] arrITEM = this.DBC.DBQuery(sbSQL).getMultipleValue();
                    if (arrITEM.Length > 0)
                    {
                        gvINPUT.SetRowCellValue(e.RowHandle, "ItemName", arrITEM[1]);
                        gvINPUT.SetRowCellValue(e.RowHandle, "StyleNo", arrITEM[3]);
                    }
                    else
                    {
                        gvINPUT.SetRowCellValue(e.RowHandle, "ItemName", "");
                        gvINPUT.SetRowCellValue(e.RowHandle, "StyleNo", "");
                    }
                }
            }
        }

        private void gvINPUT_ShownEditor(object sender, EventArgs e)
        {
            //GridView view = sender as GridView;
            //if (view.IsNewItemRow(view.FocusedRowHandle))
            //    view.ActiveEditor.IsModified = true;
        }

        private void repositoryItemSearchLookUpEdit2_EditValueChanged(object sender, EventArgs e)
        {

        }


        private void gvINPUT_InitNewRow(object sender, InitNewRowEventArgs e)
        {
            gvINPUT.SetRowCellValue(e.RowHandle, "BookingFabric", 0);
            gvINPUT.SetRowCellValue(e.RowHandle, "BookingAccessory", 0);
        }

        private void InputByTable()
        {
            layoutControlItem3.Visibility = LayoutVisibility.Never;
            layoutControlItem38.Visibility = LayoutVisibility.Never;
            layoutControlItem40.Visibility = LayoutVisibility.Never;
            emptySpaceItem3.Visibility = LayoutVisibility.Never;
            layoutControlItem4.Visibility = LayoutVisibility.Never;
            layoutControlItem5.Visibility = LayoutVisibility.Never;
            emptySpaceItem9.Visibility = LayoutVisibility.Never;
            layoutControlItem39.Visibility = LayoutVisibility.Never;
            layoutControlItem6.Visibility = LayoutVisibility.Never;
            emptySpaceItem8.Visibility = LayoutVisibility.Never;
            layoutControlItem7.Visibility = LayoutVisibility.Never;
            emptySpaceItem6.Visibility = LayoutVisibility.Never;
            layoutControlItem8.Visibility = LayoutVisibility.Never;
            emptySpaceItem7.Visibility = LayoutVisibility.Never;
            layoutControlItem9.Visibility = LayoutVisibility.Never;
            layoutControlItem10.Visibility = LayoutVisibility.Never;
            emptySpaceItem14.Visibility = LayoutVisibility.Never;
            layoutControlItem11.Visibility = LayoutVisibility.Never;
            emptySpaceItem17.Visibility = LayoutVisibility.Never;
            layoutControlItem12.Visibility = LayoutVisibility.Never;
            layoutControlItem13.Visibility = LayoutVisibility.Never;
            emptySpaceItem19.Visibility = LayoutVisibility.Never;
            layoutControlItem14.Visibility = LayoutVisibility.Never;
            emptySpaceItem18.Visibility = LayoutVisibility.Never;
            layoutControlItem15.Visibility = LayoutVisibility.Never;
            emptySpaceItem13.Visibility = LayoutVisibility.Never;
            layoutControlItem17.Visibility = LayoutVisibility.Never;
            emptySpaceItem10.Visibility = LayoutVisibility.Never;
            layoutControlItem16.Visibility = LayoutVisibility.Never;
            layoutControlItem18.Visibility = LayoutVisibility.Never;
            layoutControlItem19.Visibility = LayoutVisibility.Never;
            layoutControlItem20.Visibility = LayoutVisibility.Never;
            layoutControlItem21.Visibility = LayoutVisibility.Never;
            layoutControlItem22.Visibility = LayoutVisibility.Never;
            layoutControlItem23.Visibility = LayoutVisibility.Never;
            layoutControlItem24.Visibility = LayoutVisibility.Never;
            layoutControlGroup4.Visibility = LayoutVisibility.Never;
            layoutControlGroup5.Visibility = LayoutVisibility.Never;
            layoutControlGroup6.Visibility = LayoutVisibility.Never;
            emptySpaceItem16.Visibility = LayoutVisibility.Never;

            gcINPUT.Visible = true;
        }

        private void sbClearTable_Click(object sender, EventArgs e)
        {
            gvINPUT.CloseEditor();
            gvINPUT.UpdateCurrentRow();
            for (int i = gvINPUT.RowCount - 1; i >= 0; i--)
                gvINPUT.DeleteRow(i);
            gvINPUT.OptionsView.ColumnAutoWidth = false;
            gvINPUT.BestFitColumns();
            gvINPUT.Columns["ProductionPlanID"].Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
        }

        private bool chkDuplicateNo()
        {
            bool chkDup = true;
            if (txePlanID.Text != "")
            {
                txePlanID.Text = txePlanID.Text.ToUpper().Trim();
                if (txePlanID.Text.Trim() != "" && lblStatus.Text == "* Add Forecast")
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT TOP(1) ProductionPlanID FROM COForecast WHERE (ProductionPlanID = N'" + txePlanID.Text.Trim().Trim().Replace("'", "''") + "') ");
                    if (this.DBC.DBQuery(sbSQL).getString() != "")
                    {
                        txePlanID.Text = "";
                        txePlanID.Focus();
                        chkDup = false;
                        FUNC.msgWarning("Duplicate production plan ID. !! Please Change.");
                    }
                }
                else if (txePlanID.Text.Trim() != "" && lblStatus.Text == "* Edit Forecast")
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT TOP(1) OIDFC ");
                    sbSQL.Append("FROM COForecast ");
                    sbSQL.Append("WHERE (ProductionPlanID = N'" + txePlanID.Text.Trim().Trim().Replace("'", "''") + "') ");
                    string strCHK = this.DBC.DBQuery(sbSQL).getString();
                    if (strCHK != "" && strCHK != txeID.Text.Trim())
                    {
                        txePlanID.Text = "";
                        txePlanID.Focus();
                        chkDup = false;
                        FUNC.msgWarning("Duplicate production plan ID. !! Please Change.");
                    }
                }
            }
            return chkDup;
        }

        private void txePlanID_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                slueCustomer.Focus();
        }

        private void txePlanID_Leave(object sender, EventArgs e)
        {
            if (txePlanID.Text.Trim() != "")
            {
                txePlanID.Text = txePlanID.Text.ToUpper().Trim();
                DataTable dtCHK = (DataTable)gcINPUT.DataSource;
                if (dtCHK.Rows.Count > 0)
                {
                    int iRow = 0;
                    foreach (DataRow row in dtCHK.Rows)
                    {
                        string PlanID = row["ProductionPlanID"].ToString().ToUpper().Trim();
                        if (txePlanID.Text == PlanID)
                        {
                            FUNC.msgWarning("This production plan ID is the same as the production plan ID above !! Please change.");
                            txePlanID.Text = "";
                            txePlanID.Focus();
                            break;
                        }
                        iRow++;
                    }
                }
                else
                {
                    bool chkDup = chkDuplicateNo();
                    if (chkDup == true)
                    {
                        slueCustomer.Focus();
                    }
                    else
                    {
                        txePlanID.Text = "";
                        txePlanID.Focus();
                    }
                }
            }
        }

        private void slueCustomer_EditValueChanged(object sender, EventArgs e)
        {
            dteOrderDate.Focus();
        }

        private void dteOrderDate_EditValueChanged(object sender, EventArgs e)
        {
            //speSeason.Focus();
        }

        private void speSeason_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                glueSeason.Focus();
        }

        private void glueSeason_EditValueChanged(object sender, EventArgs e)
        {
            glueUnit.Focus();
        }

        private void glueUnit_EditValueChanged(object sender, EventArgs e)
        {
            //slueItemCode.Focus();
        }

        private void glueUnit_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode==Keys.Enter)
                slueItemCode.Focus();
        }

        private void txeSampleCode_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                slueSupplier.Focus();
        }

        private void slueSupplier_EditValueChanged(object sender, EventArgs e)
        {
            txeSewingDifficulty.Focus();
        }

        private void txeSewingDifficulty_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txePlanType.Focus();
        }

        private void txePlanType_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                glueStatus.Focus();
        }

        private void glueStatus_EditValueChanged(object sender, EventArgs e)
        {
            dteDataUpdate.Focus();
        }

        private void dteDataUpdate_EditValueChanged(object sender, EventArgs e)
        {
            //dteLastUpdate.Focus();
        }

        private void dteLastUpdate_EditValueChanged(object sender, EventArgs e)
        {
            //dteWHDate.Focus();
        }

        private void dteWHDate_EditValueChanged(object sender, EventArgs e)
        {
            //dteContractDate.Focus();
        }

        private void dteContractDate_EditValueChanged(object sender, EventArgs e)
        {
            //glueTransport.Focus();
        }

        private void glueTransport_EditValueChanged(object sender, EventArgs e)
        {
            glueLogisticsType.Focus();
        }

        private void glueLogisticsType_EditValueChanged(object sender, EventArgs e)
        {
            //txeOrderQty.Focus();
        }

        private void glueLogisticsType_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txeOrderQty.Focus();
        }

        private void txeOrderQty_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txeFabricOrderNo.Focus();
        }

        private void txeFabricOrderNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txeFabricOrderNo.Focus();
        }

        private void dteFabricUpdate_EditValueChanged(object sender, EventArgs e)
        {
            //txeFabricQty.Focus();
        }

        private void txeFabricQty_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txeColorOrderNo.Focus();
        }

        private void txeColorOrderNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                dteColorUpdate.Focus();
        }

        private void dteColorUpdate_EditValueChanged(object sender, EventArgs e)
        {
            //txeColorQty.Focus();
        }

        private void txeColorQty_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txeTrimOrderNo.Focus();
        }

        private void txeTrimOrderNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                dteTrimUpdate.Focus();
        }

        private void dteTrimUpdate_EditValueChanged(object sender, EventArgs e)
        {
            //txeTrimQty.Focus();
        }

        private void txeTrimQty_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txePOOrderNo.Focus();
        }

        private void txePOOrderNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                dtePOUpdate.Focus();
        }

        private void dtePOUpdate_EditValueChanged(object sender, EventArgs e)
        {
            //txePOQty.Focus();
        }

        private void txePOQty_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txeOrderQtyOld.Focus();
        }

        private void txeOrderQtyOld_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txeID.Focus();
        }

        private void gvINPUT_ValidateRow(object sender, DevExpress.XtraGrid.Views.Base.ValidateRowEventArgs e)
        {

            GridView view = sender as GridView;
            DevExpress.XtraGrid.Columns.GridColumn PlanIDCol = view.Columns["ProductionPlanID"];
            string strPlanID = (String)view.GetRowCellValue(e.RowHandle, PlanIDCol);
            if (strPlanID.ToUpper().Trim() == txePlanID.Text.ToUpper().Trim())
            {
                e.Valid = false;
                view.SetColumnError(PlanIDCol, "This production plan ID is the same as the production plan ID above !! Please change.");
            }
            else
            {
                bool chkPlan = chkDupPlanID(strPlanID, e.RowHandle);
                //Validity criterion
                if (chkPlan == false)
                {
                    e.Valid = false;
                    //Set errors with specific descriptions for the columns
                    view.SetColumnError(PlanIDCol, "Duplicate production plan ID. !! Please change.");
                }

                //DevExpress.XtraGrid.Columns.GridColumn OIDCUST = view.Columns["OIDCUST"];
                //string strOIDCUST = (String)view.GetRowCellValue(e.RowHandle, OIDCUST);
                //if(strOIDCUST == "")
                //    view.SetColumnError(OIDCUST, "Customer is not null.");
            }
        }

        private bool chkDupPlanID(string PlanID, int rowIndex)
        {
            gvINPUT.CloseEditor();
            gvINPUT.UpdateCurrentRow();

            PlanID = PlanID.ToUpper().Trim();
            bool chkDup = true;

            if (PlanID != "")
            {
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT COUNT(OIDFC) AS COUNT_ID ");
                sbSQL.Append("FROM COForecast ");
                sbSQL.Append("WHERE (ProductionPlanID = N'" + PlanID + "') ");
                if (this.DBC.DBQuery(sbSQL).getInt() > 0)
                    chkDup = false;
                else
                {
                    int countPlan = 0;
                    DataTable dtFind = (DataTable)gcINPUT.DataSource;
                    int xRow = 0;
                    foreach (DataRow row in dtFind.Rows)
                    {
                        string chkPlanID = row["ProductionPlanID"].ToString().ToUpper().Trim();
                        if (chkPlanID == PlanID && xRow != rowIndex)
                            countPlan++;
                        xRow++;
                    }
                    // MessageBox.Show(countPlan.ToString());
                    if (countPlan > 0)
                        chkDup = false;
                }
            }
            return chkDup;
        }

        private void gvINPUT_InvalidValueException(object sender, DevExpress.XtraEditors.Controls.InvalidValueExceptionEventArgs e)
        {

        }

        private void gvINPUT_InvalidRowException(object sender, DevExpress.XtraGrid.Views.Base.InvalidRowExceptionEventArgs e)
        {
            e.ExceptionMode = DevExpress.XtraEditors.Controls.ExceptionMode.NoAction;
        }

        private void gvINPUT_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {
            if (e.Column.FieldName == "ProductionPlanID" 
                || e.Column.FieldName == "OIDCUST" 
                || e.Column.FieldName == "OIDCSITEM"
                || e.Column.FieldName == "OIDVEND")
            {
                DevExpress.XtraEditors.ViewInfo.BaseEditViewInfo info = ((DevExpress.XtraGrid.Views.Grid.ViewInfo.GridCellInfo)e.Cell).ViewInfo;
                string error = GetError(e.CellValue, e.RowHandle, e.Column);
                if (e.CellValue == null || String.IsNullOrEmpty(e.CellValue.ToString()) == true)
                {
                    SetError(info, error);
                }
            }
        }

        private void SetError(DevExpress.XtraEditors.ViewInfo.BaseEditViewInfo cellInfo, string errorIconText)
        {
            if (errorIconText == string.Empty) return;
            cellInfo.ErrorIconText = errorIconText;
            cellInfo.ShowErrorIcon = true;
            cellInfo.FillBackground = true;
            cellInfo.ErrorIcon = DevExpress.XtraEditors.DXErrorProvider.DXErrorProvider.GetErrorIconInternal(DevExpress.XtraEditors.DXErrorProvider.ErrorType.Critical);
        }

        private string GetError(object value, int rowHandle, DevExpress.XtraGrid.Columns.GridColumn column)
        {
            //some code here
            return "Value doesn't exist";
        }

        private int CountError()
        {
            int CError = 0;
            DataTable dtError = (DataTable)gcINPUT.DataSource;
            if (dtError.Rows.Count > 0)
            {
                int iRow = 0;
                //MessageBox.Show("CountRow:" + dtError.Rows.Count.ToString());
                foreach (DataRow row in dtError.Rows)
                {
                    if (iRow < dtError.Rows.Count)
                    {
                        foreach (DataColumn c in row.Table.Columns) 
                        {
                            if (c.ColumnName == "ProductionPlanID" || c.ColumnName == "OIDCUST" || c.ColumnName == "OIDCSITEM" || c.ColumnName == "OIDVEND")
                                //MessageBox.Show(row[c.ColumnName].ToString());
                                if (row[c.ColumnName].ToString() == "")
                                    CError++;
                        }  
                    }
                    iRow++;
                }
            }
            return CError;
        }


        private void LoadForecast(string typeLoad, string ID)
        {
            ID = ID.Trim();
            if (ID != "")
            {
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT TOP (1) OIDFC, ProductionPlanID, OIDCUST, FileOrderDate, Season, BusinessUnit, OIDCSITEM, ModelNo, OIDVEND, SewingDifficulty, ProductionPlanType, Status, DataUpdate, BookingFabric, BookingAccessory, LastUpdate, ");
                sbSQL.Append("      RequestedWHDate, ContractedDate, TransportMethod, LogisticsType, OrderQty, FabricOrderNO, FabricUpdateDate, FabricActualOrderQty, ColorOrderNO, ColorUpdateDate, ColorActualOrderQty, TrimOrderNO, TrimUpdateDate, ");
                sbSQL.Append("      TrimActualOrderQty, POOrderNO, POUpdateDate, POActualOrderQty, OrderQTYOld, CreateBy, CreateDate, UpdateBy, Updatedate ");
                sbSQL.Append("FROM  COForecast ");
                sbSQL.Append("WHERE (OIDFC = '" + ID + "') ");
                string[] arrFC = this.DBC.DBQuery(sbSQL).getMultipleValue();
                if (arrFC.Length > 0)
                {
                    if (typeLoad == "New")
                    {
                        lblStatus.Text = "* Add Forecast";
                        lblStatus.ForeColor = Color.Green;

                        txeID.Text = this.DBC.DBQuery("SELECT CASE WHEN ISNULL(MAX(OIDFC), '') = '' THEN 1 ELSE MAX(OIDFC) + 1 END AS NewNo FROM COForecast").getString();
                        txePlanID.Text = "";

                        sbClearTable.Enabled = true;
                        gcINPUT.Enabled = true;
                    }
                    else if (typeLoad == "Edit")
                    {
                        lblStatus.Text = "* Edit Forecast";
                        lblStatus.ForeColor = Color.Red;

                        txeID.Text = arrFC[0];
                        txePlanID.Text = arrFC[1];

                        sbClearTable.PerformClick();
                        sbClearTable.Enabled = false;
                        gcINPUT.Enabled = false;
                        
                    }

                    slueCustomer.EditValue = arrFC[2];
                    dteOrderDate.EditValue = Convert.ToDateTime(arrFC[3]);

                    string Season = arrFC[4];
                    speSeason.Value = Convert.ToInt32(Regex.Match(Season, @"\d+([,\.]\d+)?").Value);
                    glueSeason.EditValue = arrFC[4].Replace(Regex.Match(Season, @"\d+([,\.]\d+)?").Value, "");

                    glueUnit.Text = arrFC[5];
                    slueItemCode.EditValue = Convert.ToInt32(arrFC[6]);
                    txeSampleCode.Text = arrFC[7];
                    slueSupplier.EditValue = arrFC[8];
                    txeSewingDifficulty.Text = arrFC[9];
                    txePlanType.Text = arrFC[10];
                    glueStatus.EditValue = arrFC[11];
                    if (arrFC[12] != "")
                        dteDataUpdate.EditValue = Convert.ToDateTime(arrFC[12]);
                    else
                        dteDataUpdate.EditValue = null;
                    if (arrFC[13] != "")
                        chkBookFabric.EditValue = Convert.ToInt32(arrFC[13]);
                    else
                        chkBookFabric.EditValue = null;
                    if (arrFC[14] != "")
                        chkBookAcc.EditValue = Convert.ToInt32(arrFC[14]);
                    else
                        chkBookAcc.EditValue = null;
                    if (arrFC[15] != "")
                        dteLastUpdate.EditValue = Convert.ToDateTime(arrFC[15]);
                    else
                        dteLastUpdate.EditValue = null;
                    if (arrFC[16] != "")
                        dteWHDate.EditValue = Convert.ToDateTime(arrFC[16]);
                    else
                        dteWHDate.EditValue = null;
                    if (arrFC[17] != "")
                        dteContractDate.EditValue = Convert.ToDateTime(arrFC[17]);
                    else
                        dteContractDate.EditValue = null;
                    glueTransport.EditValue = arrFC[18];
                    glueLogisticsType.Text = arrFC[19];
                    txeOrderQty.Text = arrFC[20];
                    txeFabricOrderNo.Text = arrFC[21];
                    if (arrFC[22] != "")
                        dteFabricUpdate.EditValue = Convert.ToDateTime(arrFC[22]);
                    else
                        dteFabricUpdate.EditValue = null;
                    txeFabricQty.Text = arrFC[23];
                    txeColorOrderNo.Text = arrFC[24];
                    if (arrFC[25] != "")
                        dteColorUpdate.EditValue = Convert.ToDateTime(arrFC[25]);
                    else
                        dteColorUpdate.EditValue = null;
                    txeColorQty.Text = arrFC[26];
                    txeTrimOrderNo.Text = arrFC[27];
                    if (arrFC[28] != "")
                        dteTrimUpdate.EditValue = Convert.ToDateTime(arrFC[28]);
                    else
                        dteTrimUpdate.EditValue = null;
                    txeTrimQty.Text = arrFC[29];
                    txePOOrderNo.Text = arrFC[30];
                    if (arrFC[31] != "")
                        dtePOUpdate.EditValue = Convert.ToDateTime(arrFC[31]);
                    else
                        dtePOUpdate.EditValue = null;
                    txePOQty.Text = arrFC[32];
                    txeOrderQtyOld.Text = arrFC[33];

                    txePlanID.Focus();
                    tabbedControlGroup1.SelectedTabPage = layoutControlGroup2;
                }
            }
        }

        private void bbiEdit_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (gvFO.RowCount > 0)
            {
                if (gvFO.Columns["ID"] == null)
                    FUNC.msgWarning("Please select data.");
                else
                {
                    string ID = gvFO.GetFocusedRowCellValue(gvFO.Columns["ID"]).ToString();
                    if (ID == "")
                        FUNC.msgWarning("Data is incorrect !!");
                    else
                        LoadForecast("Edit", ID);
                }
            }
            else
            {
                FUNC.msgWarning("Data not found.");
            }
        }

        private void bbiClone_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (gvFO.RowCount > 0)
            {
                if (gvFO.Columns["ID"] == null)
                    FUNC.msgWarning("Please select data.");
                else
                {
                    string ID = gvFO.GetFocusedRowCellValue(gvFO.Columns["ID"]).ToString();
                    if (ID == "")
                        FUNC.msgWarning("Data is incorrect !!");
                    else
                        LoadForecast("New", ID);
                }
            }
            else
            {
                FUNC.msgWarning("Data not found.");
            }
        }

        private void gvFO_DoubleClick(object sender, EventArgs e)
        {
            GridView view = (GridView)sender;
            Point pt = view.GridControl.PointToClient(Control.MousePosition);
            DevExpress.XtraGrid.Views.Grid.ViewInfo.GridHitInfo info = view.CalcHitInfo(pt);
            if (info.InRow || info.InRowCell)
            {
                DataTable dtCP = (DataTable)gcFO.DataSource;
                if (dtCP.Rows.Count > 0)
                {
                    DataRow drCP = dtCP.Rows[info.RowHandle];
                    string ID =  drCP["ID"].ToString();
                    LoadForecast("Edit", ID);
                }
            }
        }

        private void sbBrowse_Click(object sender, EventArgs e)
        {
            cbeSheet.Properties.Items.Clear();
            cbeSheet.Text = "";

            xtraOpenFileDialog1.Filter = "Excel files |*.xlsx;*.xls;*.csv";
            xtraOpenFileDialog1.FileName = "";
            xtraOpenFileDialog1.Title = "Select Excel File";

            if (xtraOpenFileDialog1.ShowDialog() == DialogResult.OK)
            {
                txeFilePath.Text = xtraOpenFileDialog1.FileName;
                DevExpress.XtraSpreadsheet.SpreadsheetControl xss = new DevExpress.XtraSpreadsheet.SpreadsheetControl();
                IWorkbook workbook = xss.Document;
                using (FileStream stream = new FileStream(txeFilePath.Text, FileMode.Open))
                {
                    string ext = Path.GetExtension(txeFilePath.Text);
                    if (ext == ".xlsx")
                        workbook.LoadDocument(stream, DocumentFormat.Xlsx);
                    else if (ext == ".xls")
                        workbook.LoadDocument(stream, DocumentFormat.Xls);
                    else if (ext == ".csv")
                        workbook.LoadDocument(stream, DocumentFormat.Csv);
                }
                WorksheetCollection worksheets = workbook.Worksheets;
                for (int i = 0; i < worksheets.Count; i++)
                    cbeSheet.Properties.Items.Add(worksheets[i].Name);
            }
        }

        private void cbeSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txeFilePath.Text.Trim() != "" && cbeSheet.Text.Trim() != "")
            {
                IWorkbook workbook = spsForecast.Document;

                try
                {
                    using (FileStream stream = new FileStream(txeFilePath.Text, FileMode.Open))
                    {
                        // workbook.CalculateFull();
                        string ext = Path.GetExtension(txeFilePath.Text);
                        if (ext == ".xlsx")
                            workbook.LoadDocument(stream, DocumentFormat.Xlsx);
                        else if (ext == ".xls")
                            workbook.LoadDocument(stream, DocumentFormat.Xls);
                        else if (ext == ".csv")
                            workbook.LoadDocument(stream, DocumentFormat.Csv);
                        //workbook.Worksheets.ActiveWorksheet = workbook.Worksheets[0];

                        //***Delete sheet
                        if (workbook.Worksheets.Count > 0)
                        {
                            for (int i = workbook.Worksheets.Count - 1; i >= 0; i--)
                            {
                                if(workbook.Worksheets[i].Name != cbeSheet.Text)
                                    workbook.Worksheets.RemoveAt(i);
                            }
                        }
                    }
                }
                catch (Exception)
                {
                    FUNC.msgWarning("Please close excel file before import.");
                    txeFilePath.Text = "";
                }
            }
        }

        private void sbAddItem_Click(object sender, EventArgs e)
        {
            var frm = new MPS01_01();
            frm.ShowDialog(this);
        }

        private void gvFO_RowCellClick(object sender, RowCellClickEventArgs e)
        {
            if (gvFO.IsFilterRow(e.RowHandle)) return;
        }

        private void gvINPUT_RowCellClick(object sender, RowCellClickEventArgs e)
        {
            if (gvINPUT.IsFilterRow(e.RowHandle)) return;
        }

        private void ribbonControl_Click(object sender, EventArgs e)
        {

        }

        private void gvFO_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            //if (e.Info.IsRowIndicator) e.Info.DisplayText = (e.RowHandle + 1).ToString();
        }
    }
}