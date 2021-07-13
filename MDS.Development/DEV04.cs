using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using DevExpress.LookAndFeel;
using DevExpress.XtraGrid.Views.Grid;
using TheepClass;
using DBConnect;
using System.Collections.Generic;
using DevExpress.XtraGrid.Views.BandedGrid;

namespace MDS.Development
{
    public partial class DEV04 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private Functionality.Function FUNC = new Functionality.Function();
        List<PatternZone> PZone;

        DataTable dtFB = new DataTable();
        public LogIn UserLogin { get; set; }
        public int Company { get; set; }

        public DEV04()
        {
            InitializeComponent();
            UserLookAndFeel.Default.StyleChanged += MyStyleChanged;

            PZone = new List<PatternZone>();
            PZone.Add(new PatternZone { ID = 0, Zone = "Japan" });
            PZone.Add(new PatternZone { ID = 1, Zone = "EU" });
            PZone.Add(new PatternZone { ID = 2, Zone = "US" });
        }

        private void MyStyleChanged(object sender, EventArgs e)
        {
            UserLookAndFeel userLookAndFeel = (UserLookAndFeel)sender;
            cUtility.SaveRegistry(@"Software\MDS", "SkinName", userLookAndFeel.SkinName);
            cUtility.SaveRegistry(@"Software\MDS", "SkinPalette", userLookAndFeel.ActiveSvgPaletteName);
        }

        private void CreateBands(out DevExpress.XtraGrid.Views.BandedGrid.GridBand gridBand1, out DevExpress.XtraGrid.Views.BandedGrid.GridBand gridBand2)
        {
            gridBand1 = new DevExpress.XtraGrid.Views.BandedGrid.GridBand();
            gridBand2 = new DevExpress.XtraGrid.Views.BandedGrid.GridBand();
            gridBand1.Name = "BandsHeader";
            gridBand1.Caption = "FABRIC COST DATA";
            gridBand1.AppearanceHeader.Font = new Font(gridBand1.AppearanceHeader.Font.Name, 40);
            gridBand1.RowCount = 1;

            gridBand2.Name = "BandsEstCost";
            gridBand2.Caption = "ESTIMATE COST";
            gridBand1.RowCount = 1;

            bgvFabric.Bands.Add(gridBand1);
            bgvFabric.Bands.Add(gridBand2);

            bgvFabric.Bands[0].AppearanceHeader.Font = new Font(gridBand1.AppearanceHeader.Font.Name, 10, FontStyle.Bold);
            bgvFabric.Bands[0].AppearanceHeader.BackColor = Color.FromArgb(64, 0, 0);
            bgvFabric.Bands[0].AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;


            bgvFabric.Bands[1].AppearanceHeader.Font = new Font(gridBand1.AppearanceHeader.Font.Name, 9);
            bgvFabric.Bands[1].AppearanceHeader.BackColor = Color.DarkRed;
            bgvFabric.Bands[1].AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
        }

        private void CreateColumns(DevExpress.XtraGrid.Views.BandedGrid.GridBand gridBand1, DevExpress.XtraGrid.Views.BandedGrid.GridBand gridBand2)
        {
            //***** 1 **********
            //DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn bandedGridColumn1 = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn bandedGridColumn2 = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn bandedGridColumn3 = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn bandedGridColumn4 = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn bandedGridColumn5 = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn bandedGridColumn6 = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn bandedGridColumn7 = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn bandedGridColumn8 = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn bandedGridColumn9 = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn bandedGridColumn10 = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn bandedGridColumn11 = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn bandedGridColumn12 = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn bandedGridColumn13 = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
            DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn bandedGridColumn14 = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();

            //bandedGridColumn1.Caption = "Type";
            //bandedGridColumn1.FieldName = "bcType";
            //bandedGridColumn1.Visible = true;

            //bandedGridColumn1.RowCount = 2;

            bandedGridColumn2.Caption = "Size";
            bandedGridColumn2.FieldName = "bcSize";
            bandedGridColumn2.Visible = true;

            bandedGridColumn3.Caption = "Fabric No.";
            bandedGridColumn3.FieldName = "bcFabric";
            bandedGridColumn3.Visible = true;

            bandedGridColumn4.Caption = "Vendor";
            bandedGridColumn4.FieldName = "bcVendor";
            bandedGridColumn4.Visible = true;

            bandedGridColumn5.Caption = "Garment" + Environment.NewLine + "Part";
            bandedGridColumn5.FieldName = "bcGarmentPart";
            bandedGridColumn5.Visible = true;

            bandedGridColumn6.Caption = "Color No.";
            bandedGridColumn6.FieldName = "bcColorNo";
            bandedGridColumn6.Visible = true;

            bandedGridColumn7.Caption = "Color Name";
            bandedGridColumn7.FieldName = "bcColorName";
            bandedGridColumn7.Visible = true;

            bandedGridColumn8.Caption = "Width" + Environment.NewLine + "(All)";
            bandedGridColumn8.FieldName = "bcWidthAll";
            bandedGridColumn8.Visible = true;

            bandedGridColumn9.Caption = "Width" + Environment.NewLine + "(Use)";
            bandedGridColumn9.FieldName = "bcWidthUse";
            bandedGridColumn9.Visible = true;

            bandedGridColumn10.Caption = "g/m";
            bandedGridColumn10.FieldName = "bcGM";
            bandedGridColumn10.Visible = true;

            bandedGridColumn11.Caption = "m/1P";
            bandedGridColumn11.FieldName = "bcM1P";
            bandedGridColumn11.Visible = true;

            bandedGridColumn12.Caption = "kg/1P";
            bandedGridColumn12.FieldName = "bcKg1P";
            bandedGridColumn12.Visible = true;

            bandedGridColumn13.Caption = "Price" + Environment.NewLine + "(Baht)";
            bandedGridColumn13.FieldName = "bcPrice";
            bandedGridColumn13.Visible = true;

            bandedGridColumn14.Caption = "%Loss";
            bandedGridColumn14.FieldName = "bcLoss";
            bandedGridColumn14.Visible = true;


            //gridBand1.Columns.Add(bandedGridColumn1);
            gridBand1.Columns.Add(bandedGridColumn2);
            gridBand1.Columns.Add(bandedGridColumn3);
            gridBand1.Columns.Add(bandedGridColumn4);
            gridBand1.Columns.Add(bandedGridColumn5);
            gridBand1.Columns.Add(bandedGridColumn6);
            gridBand1.Columns.Add(bandedGridColumn7);
            gridBand1.Columns.Add(bandedGridColumn8);
            gridBand1.Columns.Add(bandedGridColumn9);
            gridBand1.Columns.Add(bandedGridColumn10);
            gridBand1.Columns.Add(bandedGridColumn11);
            gridBand1.Columns.Add(bandedGridColumn12);
            gridBand1.Columns.Add(bandedGridColumn13);
            gridBand1.Columns.Add(bandedGridColumn14);


            //***** 2 **********
            List<string> listSize = new List<string>();
            listSize.Add("S");
            listSize.Add("M");
            listSize.Add("S");
            listSize.Add("M");


            List<string> listColor = new List<string>();
            listColor.Add("03 Gray");
            listColor.Add("03 Gray");
            listColor.Add("16 Red");
            listColor.Add("16 Red");

            int x = 0;
            foreach (string items in listColor)
            {
                DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn bGridCol = new DevExpress.XtraGrid.Views.BandedGrid.BandedGridColumn();
                bGridCol.Caption = listSize[x] + Environment.NewLine  + items;
                bGridCol.FieldName = "BCS" + (x + 1).ToString();
                bGridCol.Visible = true;
                
                gridBand2.Columns.Add(bGridCol);
                x++;
            }

            bgvFabric.ColumnPanelRowHeight = 38;
            bgvFabric.Appearance.HeaderPanel.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
            bgvFabric.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            bgvFabric.Appearance.BandPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            bgvFabric.Appearance.FooterPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;

            for (int ii = 0; ii < bgvFabric.Columns.Count; ii++)
                bgvFabric.Columns[ii].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;

            if (bgvFabric.Columns.Count > 13)
            {
                for (int ii = 13; ii < bgvFabric.Columns.Count; ii++)
                {
                    bgvFabric.Columns[ii].AppearanceHeader.BackColor = Color.FromArgb(255, 222, 222);
                }
            }

            bgvFabric.OptionsView.ColumnAutoWidth = false;
            bgvFabric.BestFitColumns();
            //gridBand2.Columns.Add(listColumnColor);
        }

        private void XtraForm1_Load(object sender, EventArgs e)
        {
            LoadData();
            NewData();

            DevExpress.XtraGrid.Views.BandedGrid.GridBand gridBand1;
            DevExpress.XtraGrid.Views.BandedGrid.GridBand gridBand2;

            CreateBands(out gridBand1, out gridBand2);
            CreateColumns(gridBand1, gridBand2);

            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT 'S' AS bcSize, 'FDSSTRKJ91' AS bcFabric, 'Nan Yang' AS bcVendor, 'Body' As bcGarmentPart, '03 Gray' As bcColorNo, '03 GRAY(BC07)-F142C' AS bcColorName, '167.0' AS bcWidthAll, '162.0' As bcWidthUse, '275.550' AS bcGM, '0.1758' AS bcM1P, '0.0484' AS bcKg1P, '565.00' AS bcPrice, '3.0%' AS bcLoss, '2.1664' AS BCS1, '' AS BCS2, '' AS BCS3, '' AS BCS4 ");
            sbSQL.Append("UNION ALL ");
            sbSQL.Append("SELECT 'M' AS bcSize, 'FDSSTRKJ91' AS bcFabric, 'Nan Yang' AS bcVendor, 'Body' As bcGarmentPart, '03 Gray' As bcColorNo, '03 GRAY(BC07)-F142C' AS bcColorName, '167.0' AS bcWidthAll, '162.0' As bcWidthUse, '275.550' AS bcGM, '0.1758' AS bcM1P, '0.0484' AS bcKg1P, '565.00' AS bcPrice, '3.0%' AS bcLoss, '' AS BCS1, '31.9491' AS BCS2, '' AS BCS3, '' AS BCS4 ");
            sbSQL.Append("UNION ALL ");
            sbSQL.Append("SELECT 'S' AS bcSize, 'FDSSTRKJ91' AS bcFabric, 'Nan Yang' AS bcVendor, 'Body' As bcGarmentPart, '16 Red' As bcColorNo, '16K RED-F142C' AS bcColorName, '167.0' AS bcWidthAll, '162.0' As bcWidthUse, '275.550' AS bcGM, '0.1758' AS bcM1P, '0.0484' AS bcKg1P, '565.00' AS bcPrice, '3.0%' AS bcLoss, '' AS BCS1, '' AS BCS2, '22.8821' AS BCS3, '' AS BCS4 ");
            sbSQL.Append("UNION ALL ");
            sbSQL.Append("SELECT 'M' AS bcSize, 'FDSSTRKJ91' AS bcFabric, 'Nan Yang' AS bcVendor, 'Body' As bcGarmentPart, '16 Red' As bcColorNo, '16K RED-F142C' AS bcColorName, '167.0' AS bcWidthAll, '162.0' As bcWidthUse, '275.550' AS bcGM, '0.1758' AS bcM1P, '0.0484' AS bcKg1P, '565.00' AS bcPrice, '3.0%' AS bcLoss, '' AS BCS1, '' AS BCS2, '' AS BCS3, '25.9551' AS BCS4 ");
            new ObjDevEx.setGridControl(bgcFabric, bgvFabric, sbSQL).getData();
            bgvFabric.OptionsView.ColumnAutoWidth = false;
            bgvFabric.BestFitColumns();
            //List<string> listSize = new List<string>();
            //listSize.Add("S");
            //listSize.Add("M");
            //listSize.Add("S");
            //listSize.Add("M");

            //List<string> listColor = new List<string>();
            //listColor.Add("03 Gray");
            //listColor.Add("03 Gray");
            //listColor.Add("16 Red");
            //listColor.Add("16 Red");

            //gbEstimateCost.Children.Clear();

            //int i = 0;
            //foreach (string item in listSize) // Loop through List with foreach
            //{
            //    GridBand bandSize = new GridBand();
            //    bandSize.Name = "BCSSize" + (i + 1).ToString();
            //    bandSize.Caption = item;
            //    bandSize.RowCount = 1;
            //    bandSize.Children.AddBand(listColor[i]);
            //    gbEstimateCost.Children.Add(bandSize);
            //    i++;
            //}

            //////DevExpress.XtraGrid.Columns.GridColumn gbcType = new DevExpress.XtraGrid.Columns.GridColumn();
            //BandedGridColumn gbcType = new BandedGridColumn();
            //gbcType.FieldName = "CsType";
            //gbcType.Caption = "Type";
            //gbcType.AppearanceCell.BackColor = Color.Red;
            ////gvFabric.Columns.Add(gbcType);
            //CSType.Columns.Contains(gbcType);
            ////CSType.Columns.Add(gbcType);

            //dtFB = new DataTable("FABRIC");
            //dtFB.Columns.Add("CsType", typeof(String));
            //dtFB.Columns.Add("CsSize", typeof(String));
            //dtFB.Columns.Add("CSFBNo", typeof(String));
            //dtFB.Columns.Add("CSVendor", typeof(String));
            //dtFB.Columns.Add("CSGarment", typeof(String));
            //dtFB.Columns.Add("CSColorNo", typeof(String));
            //dtFB.Columns.Add("CSColorName", typeof(String));
            //dtFB.Columns.Add("CSWidthAll", typeof(String));
            //dtFB.Columns.Add("CSWidthUse", typeof(String));
            //dtFB.Columns.Add("CSGM", typeof(String));
            //dtFB.Columns.Add("CSM1P", typeof(String));
            //dtFB.Columns.Add("CSKg1P", typeof(String));
            //dtFB.Columns.Add("CSPrice", typeof(String));
            //dtFB.Columns.Add("CSLoss", typeof(String));
            //dtFB.Columns.Add("CSSize1", typeof(String));
            //dtFB.Columns.Add("CSSize2", typeof(String));
            //dtFB.Columns.Add("CSSize3", typeof(String));
            //dtFB.Columns.Add("CSSize4", typeof(String));
            //dtFB.Columns.Add("CSSize5", typeof(String));
            //dtFB.Columns.Add("CSSize6", typeof(String));
            //dtFB.Columns.Add("CSSize7", typeof(String));
            //dtFB.Columns.Add("CSSize8", typeof(String));
            //dtFB.Columns.Add("CSSize9", typeof(String));
            //dtFB.Columns.Add("CSSize10", typeof(String));

            //dtFB.Rows.Add(new object[] { "F", "S", "FDSSTRKJ91", "Nan Yang", "Body", "03 Gray", "03 GRAY(BC07)-F142C", "167.0", "162.0", "275.550", "0.1758", "0.0484", "565.00", "3.0%", "2.1664", "", "", "" });
            //dtFB.Rows.Add(new object[] { "A", "S", "FDSSTRKJ91", "Nan Yang", "Body", "03 Gray", "03 GRAY(BC07)-F142C", "167.0", "162.0", "275.550", "0.1758", "0.0484", "565.00", "3.0%", "2.1664", "", "", "" });

            //gcFabric.DataSource = dtFB;


            ////StringBuilder sbSQL = new StringBuilder();
            ////sbSQL.Append("SELECT 'F' AS CsType, 'S' AS CSSize, 'FDSSTRKJ91', 'Nan Yang', 'Body', '03 Gray', '03 GRAY(BC07)-F142C', '167.0', '162.0', '275.550', '0.1758', '0.0484', '565.00', '3.0%', '2.1664', '', '', '' ");
            ////sbSQL.Append("UNION ALL ");
            ////sbSQL.Append("SELECT 'A' AS CsType, 'S' AS CSSize, 'FDSSTRKJ91', 'Nan Yang', 'Body', '03 Gray', '03 GRAY(BC07)-F142C', '167.0', '162.0', '275.550', '0.1758', '0.0484', '565.00', '3.0%', '2.1664', '', '', '' ");
            ////new ObjDevEx.setGridControl(gcFabric, gvFabric, sbSQL).getData();

            ////DataTable dtFABRIC = dtFB;

            ////dtFABRIC.Rows.Add("F", "S", "FDSSTRKJ91", "Nan Yang", "Body", "03 Gray", "03 GRAY(BC07)-F142C", "167.0", "162.0", "275.550", "0.1758", "0.0484", "565.00", "3.0%", "2.1664", "", "", "");
            ////gcFabric.DataSource = dtFABRIC;

        }

        private void LoadBOM()
        {
            gcBOM.DataSource = null;
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT B.BOMNo AS [BOM No.], B.ModelName AS Model, B.OIDSIZE, PS.SizeNo AS [Size], PS.SizeName, B.OIDColor, PC.ColorNo AS [Color], PC.ColorName, B.OIDBOM AS ID ");
            sbSQL.Append("FROM   BOM AS B INNER JOIN ");
            sbSQL.Append("       ProductColor AS PC ON B.OIDColor = PC.OIDCOLOR INNER JOIN ");
            sbSQL.Append("       ProductSize AS PS ON B.OIDSIZE = PS.OIDSIZE ");
            if (glueSeason.Text.Trim() != "" || slueStyle.Text.Trim() != "" || glueZone.Text.Trim() != "")
            {
                sbSQL.Append("WHERE (B.OIDBOM > 0)  ");
                if (glueSeason.Text.Trim() != "")
                    sbSQL.Append("AND (B.Season = N'" + glueSeason.EditValue.ToString() + "')  ");
                if (slueStyle.Text.Trim() != "")
                    sbSQL.Append("AND (B.OIDSTYLE = '" + slueStyle.EditValue.ToString() + "')  ");
                if (glueZone.Text.Trim() != "")
                    sbSQL.Append("AND (B.PatternZone = '" + glueZone.EditValue.ToString() + "') ");
            }
            else
                sbSQL.Append("WHERE (B.OIDBOM = 0)  ");
            sbSQL.Append("ORDER BY B.OIDBOM ");

            new ObjDevEx.setGridControl(gcBOM, gvBOM, sbSQL).getData(false, false , false, true);
            gvBOM.Columns["OIDSIZE"].Visible = false;
            gvBOM.Columns["OIDColor"].Visible = false;
            gvBOM.Columns["ID"].Visible = false;
        }

        private void bbiNew_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            LoadData();
            NewData();
        }

        private void LoadData()
        {
            glueZone.Properties.DataSource = PZone;
            glueZone.Properties.DisplayMember = "Zone";
            glueZone.Properties.ValueMember = "ID";
            glueZone.Properties.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;

            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT DISTINCT Season FROM BOM ORDER BY Season");
            new ObjDevEx.setGridLookUpEdit(glueSeason, sbSQL, "Season", "Season").getData();

            sbSQL.Clear();
            sbSQL.Append("SELECT PS.StyleName, GC.CategoryName, PS.OIDSTYLE AS ID ");
            sbSQL.Append("FROM   ProductStyle AS PS INNER JOIN ");
            sbSQL.Append("       GarmentCategory AS GC ON PS.OIDGCATEGORY = GC.OIDGCATEGORY ");
            sbSQL.Append("ORDER BY PS.StyleName");
            new ObjDevEx.setSearchLookUpEdit(slueStyle, sbSQL, "StyleName", "ID").getData();

            LoadBOM();
        }

        private void NewData()
        {
            txeCostSheet.Text = "";
            dteDate.EditValue = DateTime.Now;
            glueSeason.EditValue = "";
            slueStyle.EditValue = "";
            glueZone.EditValue = 0;
            gcBOM.DataSource = null;
        }

        private void bbiSave_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {

        }

        private void bbiExcel_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
        }

        private void bbiPrintPreview_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {

        }

        private void bbiPrint_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {

        }

        private void glueSeason_EditValueChanged(object sender, EventArgs e)
        {
            LoadBOM();
            slueStyle.Focus();
        }

        private void slueStyle_EditValueChanged(object sender, EventArgs e)
        {
            LoadBOM();
            glueZone.Focus();
        }

        private void glueZone_EditValueChanged(object sender, EventArgs e)
        {
            LoadBOM();
        }

        private void gvBOM_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator) e.Info.DisplayText = (e.RowHandle + 1).ToString();
        }

        private void gvBOM_RowCellClick(object sender, RowCellClickEventArgs e)
        {
            //if (e.RowHandle > -1)
            //{
            //    txeCustomer.Text = "";
            //    txeItemNo.Text = "";
            //    txePatternNo.Text = "";
            //    txeModelNo.Text = "";
            //    txeModelName.Text = "";

            //    StringBuilder sbSQL = new StringBuilder();
            //    sbSQL.Append("SELECT TOP (1) B.OIDCUST, CUS.Name AS Customer, B.OIDITEM, IT.Code AS ItemNo, SR.SMPLPatternNo, B.BOMNo, SR.ModelName ");
            //    sbSQL.Append("FROM   BOM AS B LEFT OUTER JOIN ");
            //    sbSQL.Append("       SMPLRequest AS SR ON B.OIDSMPL = SR.OIDSMPL LEFT OUTER JOIN ");
            //    sbSQL.Append("       Customer AS CUS ON B.OIDCUST = CUS.OIDCUST LEFT OUTER JOIN ");
            //    sbSQL.Append("       Items AS IT ON B.OIDITEM = IT.OIDITEM ");
            //    sbSQL.Append("WHERE (B.OIDBOM = '" + gvBOM.GetRowCellValue(e.RowHandle, "ID").ToString() + "') ");
            //    string[] arrBOM = DB.DBQuery(sbSQL).getMultipleValue();
            //    if (arrBOM.Length > 0)
            //    {
            //        txeCustomer.Text = arrBOM[1];
            //        txeItemNo.Text = arrBOM[3];
            //        txePatternNo.Text = arrBOM[4];
            //        txeModelNo.Text = arrBOM[5];
            //        txeModelName.Text = arrBOM[6];
            //    }
            //}
        }

        private void ribbonControl_Click(object sender, EventArgs e)
        {

        }
    }

    public class PatternZone
    {
        public int ID { get; set; }
        public string Zone { get; set; }
    }
}