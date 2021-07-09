using System;
using System.Text;
using DBConnection;
using System.Windows.Forms;
using DevExpress.LookAndFeel;
using DevExpress.Utils.Extensions;
using System.Drawing;
using DevExpress.XtraGrid.Views.Grid;
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
    public partial class MPS02 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private Functionality.Function FUNC = new Functionality.Function();

        StringBuilder sbSTATUS = new StringBuilder();

        DataTable dtPOEntry = new DataTable();

        List<DocumentStatus> documentStatuses;
        List<DocumentStatus> documentDOStatuses;
        List<TransportationMethod> transportationMethod;
        public LogIn UserLogin { get; set; }

        public MPS02()
        {
            InitializeComponent();
            UserLookAndFeel.Default.StyleChanged += MyStyleChanged;
            documentStatuses = new List<DocumentStatus>();
            documentStatuses.Add(new DocumentStatus { ID = 0, Status = "New" });
            documentStatuses.Add(new DocumentStatus { ID = 1, Status = "Revise" });
            documentStatuses.Add(new DocumentStatus { ID = 2, Status = "Change" });
            documentStatuses.Add(new DocumentStatus { ID = 3, Status = "Cancel" });
            documentStatuses.Add(new DocumentStatus { ID = 9, Status = "Finished" });

            documentDOStatuses = new List<DocumentStatus>();
            documentDOStatuses.Add(new DocumentStatus { ID = 0, Status = "New" });
            documentDOStatuses.Add(new DocumentStatus { ID = 2, Status = "Change" });
            documentDOStatuses.Add(new DocumentStatus { ID = 3, Status = "Cancel" });
            documentDOStatuses.Add(new DocumentStatus { ID = 9, Status = "Finished" });

            transportationMethod = new List<TransportationMethod>();
            transportationMethod.Add(new TransportationMethod { ID = 0, TransportMethod = "Ship" });
            transportationMethod.Add(new TransportationMethod { ID = 1, TransportMethod = "Air" });
        }

        private void MyStyleChanged(object sender, EventArgs e)
        {
            UserLookAndFeel userLookAndFeel = (UserLookAndFeel)sender;
            cUtility.SaveRegistry(@"Software\MDS", "SkinName", userLookAndFeel.SkinName);
            cUtility.SaveRegistry(@"Software\MDS", "SkinPalette", userLookAndFeel.ActiveSvgPaletteName);
        }

        private void XtraForm1_Load(object sender, EventArgs e)
        {
            gluePoDocumentStatus.Properties.DataSource = documentStatuses;
            gluePoDocumentStatus.Properties.DisplayMember = "Status";
            gluePoDocumentStatus.Properties.ValueMember = "ID";
            gluePoDocumentStatus.Properties.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;

            glueDoDocStatus.Properties.DataSource = documentDOStatuses;
            glueDoDocStatus.Properties.DisplayMember = "Status";
            glueDoDocStatus.Properties.ValueMember = "ID";
            glueDoDocStatus.Properties.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;

            glueDoTransportationMethod.Properties.DataSource = transportationMethod;
            glueDoTransportationMethod.Properties.DisplayMember = "TransportMethod";
            glueDoTransportationMethod.Properties.ValueMember = "ID";
            glueDoTransportationMethod.Properties.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;


            tabbedControlGroup3.SelectedTabPage = layoutControlGroup4;
            tabbedControlGroup2.SelectedTabPage = layoutControlGroup2;

            sbSTATUS.Clear();
            sbSTATUS.Append("SELECT '0' AS ID, 'New' AS Status ");
            sbSTATUS.Append("UNION ALL ");
            sbSTATUS.Append("SELECT '1' AS ID, 'Revise' AS Status ");
            sbSTATUS.Append("UNION ALL ");
            sbSTATUS.Append("SELECT '2' AS ID, 'Change' AS Status ");
            sbSTATUS.Append("UNION ALL ");
            sbSTATUS.Append("SELECT '3' AS ID, 'Cancel' AS Status ");
            sbSTATUS.Append("UNION ALL ");
            sbSTATUS.Append("SELECT '9' AS ID, 'Finished' AS Status ");


            dtPOEntry.Columns.Add("OIDCOLOR", typeof(Int32));
            dtPOEntry.Columns.Add("ColorName", typeof(String));
            dtPOEntry.Columns.Add("OIDSIZE", typeof(Int32));
            dtPOEntry.Columns.Add("SizeName", typeof(String));
            dtPOEntry.Columns.Add("PatternDimensionCode", typeof(String));
            dtPOEntry.Columns.Add("SKUCode", typeof(String));
            dtPOEntry.Columns.Add("SampleCode", typeof(String));
            dtPOEntry.Columns.Add("OrderQtyPCS", typeof(Int32));
            dtPOEntry.Columns.Add("OID", typeof(String));
            gcEntryPO.DataSource = dtPOEntry;


            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT No AS ID, Name AS Incoterms FROM ENUMTYPE WHERE (Module = 'CODO') ORDER BY OIDENUM ");
            new ObjDevEx.setGridLookUpEdit(glueDoIncoterms, sbSQL, "Incoterms", "ID").getData();

            gluePoBusinessUnit.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard;
            gluePoBusinessUnit.Properties.AcceptEditorTextAsNewValue = DevExpress.Utils.DefaultBoolean.True;

            repositoryItemGridLookUpEdit3.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard;
            repositoryItemGridLookUpEdit3.AcceptEditorTextAsNewValue = DevExpress.Utils.DefaultBoolean.True;

            LoadData();
            NewData();

            //tabbedControlGroup1.SelectedTabPage = layoutControlGroup1;
            ////*******************************************************
            //bbiRefresh.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
            //bbiNew.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
            //bbiSave.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
            //bbiEdit.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
            //bbiClone.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;

            //ribbonPageGroup2.Visible = true;
            //ribbonPageGroup5.Visible = true;
            ////*******************************************************
        }

        private void LoadSummary()
        {
            gcSumPO.DataSource = null;
            StringBuilder sbSQL = new StringBuilder();
            //**** NEW (PO) ******
            sbSQL.Append("SELECT PO.OrderNo AS [PO Order No.], PO.RevisionNo AS [Revision No.], CONVERT(NVARCHAR(10), PO.RevisedDate, 103) AS RevisedDate, ");
            sbSQL.Append("       PO.DocumentStatus AS DocStatus, ST.Status AS DocumentStatus, PO.Lot AS [Lot No.], PO.OrderPlanNumber AS [Order Plan No.], PO.OIDCUST, CUS.Name AS Customer, ");
            sbSQL.Append("       PO.Season, PO.BusinessUnit, PO.PatternDimensionCode, PO.ItemCode, IC.ItemName, PO.SampleCode, SUM(PO.OrderQtyPCS) AS [Order Qty. (Pcs)], PO.OriginalSalesPrice, PO.Approver, ");
            sbSQL.Append("       CONVERT(NVARCHAR(10), PO.ApprovalDate, 103) AS ApprovalDate, PO.OIDBillto, VD.Name AS BillTo, VD.Address1 + ' ' + VD.Address2 + ' ' + VD.Address3 AS Address, VD.TelephoneNo AS Telephone, ");
            sbSQL.Append("       PO.PaymentTerms, PO.OIDCURR, CUR.Currency AS PaymentCurrency, PO.Remark, PO.AllocationOrderNumber ");
            sbSQL.Append("FROM   COPO AS PO INNER JOIN ");
            sbSQL.Append("       ( ");
            sbSQL.Append(sbSTATUS);
            sbSQL.Append("       ) AS ST ON PO.DocumentStatus = ST.ID LEFT OUTER JOIN ");
            sbSQL.Append("       Customer AS CUS ON PO.OIDCUST = CUS.OIDCUST LEFT OUTER JOIN ");
            sbSQL.Append("       ItemCustomer AS IC ON PO.ItemCode = IC.OIDCSITEM LEFT OUTER JOIN ");
            sbSQL.Append("       Vendor AS VD ON VD.VendorType = 6 AND PO.OIDBillto = VD.OIDVEND LEFT OUTER JOIN ");
            sbSQL.Append("       Currency AS CUR ON PO.OIDCURR = CUR.OIDCURR ");
            sbSQL.Append("GROUP BY PO.OrderNo, PO.RevisionNo, PO.RevisedDate, PO.DocumentStatus, ST.Status, PO.Lot, PO.OrderPlanNumber, PO.OIDCUST, CUS.Name, ");
            sbSQL.Append("       PO.Season, PO.BusinessUnit, PO.PatternDimensionCode, PO.ItemCode, IC.ItemName, PO.SampleCode, ");
            sbSQL.Append("       PO.OriginalSalesPrice, PO.Approver, PO.ApprovalDate, PO.OIDBillto, VD.Name, VD.Address1, VD.Address2, VD.Address3, VD.TelephoneNo, ");
            sbSQL.Append("       PO.PaymentTerms, PO.OIDCURR, CUR.Currency, PO.Remark, PO.AllocationOrderNumber ");
            sbSQL.Append("ORDER BY PO.OrderNo, PO.RevisionNo ");
            DataTable dtPO = new DBQuery(sbSQL).getDataTable();

            sbSQL.Clear();
            sbSQL.Append("SELECT PO.OrderNo AS [PO Order No.], PO.RevisionNo AS [Revision No.], PO.OIDCOLOR, PC.ColorNo AS ColorCode, PC.ColorName AS Color, PO.OIDSIZE, PS.SizeNo AS SizeCode, PS.SizeName AS Size, ");
            sbSQL.Append("       PO.PatternDimensionCode, PO.SKUCode, PO.SampleCode, PO.OrderQtyPCS AS [Order Qty. (Pcs)] ");
            sbSQL.Append("FROM   COPO AS PO LEFT OUTER JOIN ");
            sbSQL.Append("       ProductColor AS PC ON PO.OIDCOLOR = PC.OIDCOLOR LEFT OUTER JOIN ");
            sbSQL.Append("       ProductSize AS PS ON PO.OIDSIZE = PS.OIDSIZE ");
            sbSQL.Append("ORDER BY PO.OrderNo, PO.RevisionNo, PO.OID ");
            DataTable dtPO2 = new DBQuery(sbSQL).getDataTable();

            DataSet dsPOS = new DataSet();
            dsPOS.Tables.Add(dtPO);
            dsPOS.Tables.Add(dtPO2);

            dsPOS.Relations.Add("PO Detail", 
                new DataColumn[] { dsPOS.Tables[0].Columns["PO Order No."], dsPOS.Tables[0].Columns["Revision No."] }, 
                new DataColumn[] { dsPOS.Tables[1].Columns["PO Order No."], dsPOS.Tables[1].Columns["Revision No."] }
                );
            
            gcSumPO.DataSource = dsPOS.Tables[0];
            gvSumPO.OptionsView.ColumnAutoWidth = false;
            gvSumPO.BestFitColumns();
            gvSumPO.Columns["PO Order No."].Width = 160;

            gvSumPO.Columns["PO Order No."].Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
            gvSumPO.Columns["Revision No."].Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;

            gvSumPO.Columns["DocStatus"].Visible = false;
            gvSumPO.Columns["OIDCUST"].Visible = false;
            gvSumPO.Columns["ItemCode"].Visible = false;
            gvSumPO.Columns["OIDCURR"].Visible = false;
            gvSumPO.Columns["Order Qty. (Pcs)"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            //********************


            //**** NEW (DO) ******
            gcSumDO.DataSource = null;
            sbSQL.Clear();
            sbSQL.Append("SELECT DO.DONo AS [DO No.], DO.OrderNo AS [PO Order No.], DO.RevisionNo AS [Revision No.], DO.DocStatus, ST.Status AS DocumentStatus, DO.ItemCode, IC.ItemName, DO.PatternDimensionCode,  ");
            sbSQL.Append("       DO.TransportationMethod AS TransMethod, CASE WHEN DO.TransportationMethod = 0 THEN 'Ship' ELSE CASE WHEN DO.TransportationMethod = 1 THEN 'Air' ELSE '' END END AS TransportationMethod, ");
            sbSQL.Append("       DO.PortCode AS [Ship to Port Code], PC.PortName AS [Ship to Port], PC.City, PC.Country, DO.Incoterms, DO.Forwarder, DO.OIDVEND, VD.Name AS Vendor, SUM(DO.QuantityBox) AS SumQuantityBox, ");
            sbSQL.Append("       CASE WHEN DO.ETAWH IS NOT NULL THEN CONVERT(VARCHAR(10), DO.ETAWH, 103) ELSE '' END AS ETAWH, CASE WHEN DO.ContractedETD IS NOT NULL THEN CONVERT(VARCHAR(10), DO.ContractedETD, 103) ELSE '' END AS ContractedETD ");
            sbSQL.Append("FROM   CODO AS DO INNER JOIN ");
            sbSQL.Append("       ( ");
            sbSQL.Append(sbSTATUS);
            sbSQL.Append("       ) AS ST ON DO.DocStatus = ST.ID LEFT OUTER JOIN ");
            sbSQL.Append("       ItemCustomer AS IC ON DO.ItemCode = IC.OIDCSITEM LEFT OUTER JOIN ");
            sbSQL.Append("       PortAndCity AS PC ON DO.PortCode = PC.PortCode LEFT OUTER JOIN ");
            sbSQL.Append("       Vendor AS VD ON VD.VendorType = 6 AND DO.OIDVEND = VD.OIDVEND ");
            sbSQL.Append("GROUP BY DO.DONo, DO.OrderNo, DO.RevisionNo, DO.DocStatus, ST.Status, DO.ItemCode, IC.ItemName, DO.PatternDimensionCode,  ");
            sbSQL.Append("       DO.TransportationMethod, DO.PortCode, PC.PortName, PC.City, PC.Country, DO.Incoterms, DO.Forwarder, DO.OIDVEND, VD.Name, ");
            sbSQL.Append("       DO.ETAWH, DO.ContractedETD  ");
            sbSQL.Append("ORDER BY [DO No.], [PO Order No.], [Revision No.]");
            DataTable dtDO = new DBQuery(sbSQL).getDataTable();

            sbSQL.Clear();
            sbSQL.Append("SELECT DO.DONo AS [DO No.], DO.OrderNo AS [PO Order No.], DO.RevisionNo AS [Revision No.],  ");
            sbSQL.Append("       DO.SetCode, DO.QuantityBox, DO.OIDCOLOR, PC.ColorNo AS ColorCode, PC.ColorName AS Color, ");
            sbSQL.Append("       DO.OIDSIZE, PS.SizeNo AS SizeCode, PS.SizeName AS Size, DO.PatternDimensionCode, DO.QtyperSet, DO.PickingUnit ");
            sbSQL.Append("FROM   CODO AS DO LEFT OUTER JOIN ");
            sbSQL.Append("       ProductColor AS PC ON DO.OIDCOLOR = PC.OIDCOLOR LEFT OUTER JOIN ");
            sbSQL.Append("       ProductSize AS PS ON DO.OIDSIZE = PS.OIDSIZE ");
            sbSQL.Append("ORDER BY [DO No.], [PO Order No.], [Revision No.], DO.OIDDO");
            DataTable dtDO2 = new DBQuery(sbSQL).getDataTable();

            DataSet dsDOS = new DataSet();
            dsDOS.Tables.Add(dtDO);
            dsDOS.Tables.Add(dtDO2);

            dsDOS.Relations.Add("DO Detail",
               new DataColumn[] { dsDOS.Tables[0].Columns["DO No."], dsDOS.Tables[0].Columns["PO Order No."], dsDOS.Tables[0].Columns["Revision No."] },
               new DataColumn[] { dsDOS.Tables[1].Columns["DO No."], dsDOS.Tables[1].Columns["PO Order No."], dsDOS.Tables[1].Columns["Revision No."] }
               );

            gcSumDO.DataSource = dsDOS.Tables[0];

            gvSumDO.Columns["DO No."].Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
            gvSumDO.Columns["PO Order No."].Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
            gvSumDO.Columns["Revision No."].Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;

            gvSumDO.Columns["DocStatus"].Visible = false;
            gvSumDO.Columns["ItemCode"].Visible = false;
            gvSumDO.Columns["TransMethod"].Visible = false;
            gvSumDO.Columns["Ship to Port Code"].Visible = false;
            gvSumDO.Columns["OIDVEND"].Visible = false;
            gvSumDO.Columns["SumQuantityBox"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;

            gvSumDO.OptionsView.ColumnAutoWidth = false;
            gvSumDO.BestFitColumns();
            gvSumDO.Columns["DO No."].Width = 180;
            gvSumDO.Columns["PO Order No."].Width = 160;
        }

        private void LoadData()
        {
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT Code, Name AS Supplier, OIDVEND AS ID ");
            sbSQL.Append("FROM Vendor ");
            sbSQL.Append("WHERE (VendorType = 6) ");
            sbSQL.Append("ORDER BY Code ");
            new ObjDevEx.setSearchLookUpEdit(sluePoOIDBillto, sbSQL, "Supplier", "ID").getData();
            sluePoOIDBillto.Properties.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
            slueDoOIDVEND.Properties.DataSource = sluePoOIDBillto.Properties.DataSource;
            slueDoOIDVEND.Properties.DisplayMember = "Supplier";
            slueDoOIDVEND.Properties.ValueMember = "ID";
            slueDoOIDVEND.Properties.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;

            sbSQL.Clear();
            sbSQL.Append("SELECT ITC.ItemCode, ITC.ItemName, CUS.Name AS Customer, ITC.StyleNo AS [StyleNo.], ITC.Season, ITC.OIDCSITEM AS ID ");
            sbSQL.Append("FROM   ItemCustomer AS ITC LEFT OUTER JOIN ");
            sbSQL.Append("       Customer AS CUS ON ITC.OIDCUST = CUS.OIDCUST ");
            sbSQL.Append("ORDER BY ITC.ItemCode ");
            new ObjDevEx.setSearchLookUpEdit(sluePoItemCode, sbSQL, "ItemCode", "ID").getData();
            sluePoItemCode.Properties.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
            slueDoItemCode.Properties.DataSource = sluePoItemCode.Properties.DataSource;
            slueDoItemCode.Properties.DisplayMember = "ItemCode";
            slueDoItemCode.Properties.ValueMember = "ID";
            slueDoItemCode.Properties.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;

            sbSQL.Clear();
            sbSQL.Append("SELECT PortCode, PortName, City, Country, OIDPORT AS ID ");
            sbSQL.Append("FROM   PortAndCity ");
            sbSQL.Append("ORDER BY PortCode ");
            new ObjDevEx.setSearchLookUpEdit(slueDoPortCode, sbSQL, "PortCode", "ID").getData();
            slueDoPortCode.Properties.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;

            sbSQL.Clear();
            sbSQL.Append("SELECT Code, Name AS Customer, OIDCUST AS ID ");
            sbSQL.Append("FROM Customer ");
            sbSQL.Append("ORDER BY Code ");
            new ObjDevEx.setSearchLookUpEdit(sluePoOIDCUST, sbSQL, "Customer", "ID").getData();

            sbSQL.Clear();
            sbSQL.Append("SELECT DISTINCT BusinessUnit ");
            sbSQL.Append("FROM COPO ");
            sbSQL.Append("ORDER BY BusinessUnit");
            new ObjDevEx.setGridLookUpEdit(gluePoBusinessUnit, sbSQL, "BusinessUnit", "BusinessUnit").getData();

            sbSQL.Clear();
            sbSQL.Append("SELECT SeasonNo AS [Season No.], SeasonName AS [Season Name] ");
            sbSQL.Append("FROM Season ");
            sbSQL.Append("ORDER BY OIDSEASON");
            new ObjDevEx.setGridLookUpEdit(gluePoSeason, sbSQL, "Season No.", "Season No.").getData();

            sbSQL.Clear();
            DevExpress.XtraEditors.SearchLookUpEdit sluePoOIDCOLOR = new DevExpress.XtraEditors.SearchLookUpEdit();
            sbSQL.Append("SELECT ColorNo, ColorName, OIDCOLOR AS ID ");
            sbSQL.Append("FROM ProductColor ");
            sbSQL.Append("ORDER BY ColorNo");
            new ObjDevEx.setSearchLookUpEdit(sluePoOIDCOLOR, sbSQL, "ColorNo", "ID").getData();

            sbSQL.Clear();
            DevExpress.XtraEditors.SearchLookUpEdit sluePoOIDSIZE = new DevExpress.XtraEditors.SearchLookUpEdit();
            sbSQL.Append("SELECT SizeNo, SizeName, OIDSIZE AS ID ");
            sbSQL.Append("FROM ProductSize ");
            sbSQL.Append("ORDER BY SizeNo");
            new ObjDevEx.setSearchLookUpEdit(sluePoOIDSIZE, sbSQL, "SizeNo", "ID").getData();

            sbSQL.Clear();
            sbSQL.Append("SELECT Currency, OIDCURR AS ID ");
            sbSQL.Append("FROM Currency ");
            sbSQL.Append("ORDER BY OIDCURR");
            new ObjDevEx.setGridLookUpEdit(gluePoOIDCURR, sbSQL, "Currency", "ID").getData();

            sbSQL.Clear();
            sbSQL.Append("SELECT Name, Description, OIDPayment AS ID ");
            sbSQL.Append("FROM PaymentTerm ");
            sbSQL.Append("ORDER BY OIDPayment ");
            new ObjDevEx.setSearchLookUpEdit(sluePoPaymentTerms, sbSQL, "Name", "Name").getData();

            sbSQL.Clear();
            sbSQL.Append("SELECT DISTINCT PO.OrderNo + '_' + CONVERT(NVARCHAR, PO.RevisionNo) AS POID, PO.OrderNo AS [PO. Order No.], PO.RevisionNo AS [Revision No.], ST.Status, PO.Season, CUS.Name AS Customer, IC.ItemName ");
            sbSQL.Append("FROM   COPO AS PO INNER JOIN ");
            sbSQL.Append("       Customer AS CUS ON PO.OIDCUST = CUS.OIDCUST INNER JOIN ");
            sbSQL.Append("       ItemCustomer AS IC ON PO.ItemCode = IC.OIDCSITEM INNER JOIN ");
            sbSQL.Append("       (" + sbSTATUS.ToString() + ") AS ST ON PO.DocumentStatus = ST.ID ");
            sbSQL.Append("ORDER BY [PO. Order No.] ");
            new ObjDevEx.setSearchLookUpEdit(slueDoOrderNo, sbSQL, "PO. Order No.", "POID").getData();
            slueDoOrderNo.Properties.PopulateViewColumns();
            slueDoOrderNo.Properties.View.Columns["POID"].Visible = false;

            //*** SET GRIDCONTROL COLUMN *****
            //PO
            repositoryItemGridLookUpEdit1.DataSource = gluePoDocumentStatus.Properties.DataSource;
            repositoryItemGridLookUpEdit1.DisplayMember = "Status";
            repositoryItemGridLookUpEdit1.ValueMember = "ID";
            repositoryItemGridLookUpEdit1.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;

            repositoryItemSearchLookUpEdit1.DataSource = sluePoOIDCUST.Properties.DataSource;
            repositoryItemSearchLookUpEdit1.DisplayMember = "Customer";
            repositoryItemSearchLookUpEdit1.ValueMember = "ID";
            repositoryItemSearchLookUpEdit1.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;

            repositoryItemGridLookUpEdit2.DataSource = gluePoSeason.Properties.DataSource;
            repositoryItemGridLookUpEdit2.DisplayMember = "Season No.";
            repositoryItemGridLookUpEdit2.ValueMember = "Season No.";
            repositoryItemGridLookUpEdit2.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;

            repositoryItemGridLookUpEdit3.DataSource = gluePoBusinessUnit.Properties.DataSource;
            repositoryItemGridLookUpEdit3.DisplayMember = "BusinessUnit";
            repositoryItemGridLookUpEdit3.ValueMember = "BusinessUnit";
            repositoryItemGridLookUpEdit3.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;

            repositoryItemSearchLookUpEdit2.DataSource = sluePoItemCode.Properties.DataSource;
            repositoryItemSearchLookUpEdit2.DisplayMember = "ItemCode";
            repositoryItemSearchLookUpEdit2.ValueMember = "ID";
            repositoryItemSearchLookUpEdit2.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;

            repositoryItemSearchLookUpEdit3.DataSource = sluePoOIDCOLOR.Properties.DataSource;
            repositoryItemSearchLookUpEdit3.DisplayMember = "ColorNo";
            repositoryItemSearchLookUpEdit3.ValueMember = "ID";
            repositoryItemSearchLookUpEdit3.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;

            repositoryItemSearchLookUpEdit4.DataSource = sluePoOIDSIZE.Properties.DataSource;
            repositoryItemSearchLookUpEdit4.DisplayMember = "SizeNo";
            repositoryItemSearchLookUpEdit4.ValueMember = "ID";
            repositoryItemSearchLookUpEdit4.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;

            repositoryItemSearchLookUpEdit5.DataSource = sluePoOIDBillto.Properties.DataSource;
            repositoryItemSearchLookUpEdit5.DisplayMember = "Supplier";
            repositoryItemSearchLookUpEdit5.ValueMember = "ID";
            repositoryItemSearchLookUpEdit5.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;

            repositoryItemSearchLookUpEdit6.DataSource = sluePoPaymentTerms.Properties.DataSource;
            repositoryItemSearchLookUpEdit6.DisplayMember = "Name";
            repositoryItemSearchLookUpEdit6.ValueMember = "ID";
            repositoryItemSearchLookUpEdit6.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;

            repositoryItemGridLookUpEdit4.DataSource = gluePoOIDCURR.Properties.DataSource;
            repositoryItemGridLookUpEdit4.DisplayMember = "Currency";
            repositoryItemGridLookUpEdit4.ValueMember = "ID";
            repositoryItemGridLookUpEdit4.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;

            //DO
            repositoryItemGridLookUpEdit5.DataSource = glueDoDocStatus.Properties.DataSource;
            repositoryItemGridLookUpEdit5.DisplayMember = "Status";
            repositoryItemGridLookUpEdit5.ValueMember = "ID";
            repositoryItemGridLookUpEdit5.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;

            repositoryItemSearchLookUpEdit7.DataSource = slueDoItemCode.Properties.DataSource;
            repositoryItemSearchLookUpEdit7.DisplayMember = "ItemCode";
            repositoryItemSearchLookUpEdit7.ValueMember = "ID";
            repositoryItemSearchLookUpEdit7.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;

            repositoryItemGridLookUpEdit6.DataSource = glueDoTransportationMethod.Properties.DataSource;
            repositoryItemGridLookUpEdit6.DisplayMember = "TransportMethod";
            repositoryItemGridLookUpEdit6.ValueMember = "ID";
            repositoryItemGridLookUpEdit6.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;

            repositoryItemSearchLookUpEdit8.DataSource = slueDoPortCode.Properties.DataSource;
            repositoryItemSearchLookUpEdit8.DisplayMember = "PortCode";
            repositoryItemSearchLookUpEdit8.ValueMember = "ID";
            repositoryItemSearchLookUpEdit8.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;

            repositoryItemSearchLookUpEdit9.DataSource = slueDoOIDVEND.Properties.DataSource;
            repositoryItemSearchLookUpEdit9.DisplayMember = "Supplier";
            repositoryItemSearchLookUpEdit9.ValueMember = "ID";
            repositoryItemSearchLookUpEdit9.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;

            LoadSummary();
        }

        private void SetNewPO()
        {
            //*** PO ***
            gluePoDocumentStatus.EditValue = 0;

            //txePoOID.Text = new DBQuery("SELECT CASE WHEN ISNULL(MAX(OID), '') = '' THEN 1 ELSE MAX(OID) + 1 END AS NewNo FROM COPO").getString();
            txePoOrderNo.Text = "";
            spePoRevisionNo.Value = 0;
            txePoLot.Text = "";
            dtePoRevisedDate.EditValue = DBNull.Value;
            gluePoDocumentStatus.EditValue = 0;
            txePoOrderPlanNumber.Text = "";
            sluePoOIDCUST.EditValue = "";
            if (Convert.ToInt32(DateTime.Now.ToString("yyyy")) > 2500)
                spePoSeason.Value = Convert.ToInt32(DateTime.Now.ToString("yyyy")) - 543;
            else
                spePoSeason.Value = Convert.ToInt32(DateTime.Now.ToString("yyyy"));

            gluePoSeason.EditValue = "";
            gluePoBusinessUnit.EditValue = "";

            sluePoItemCode.EditValue = "";
            txePoItemName.Text = "";
            txePoFabricWidth.Text = "";
            txePoFBComposition.Text = "";

            txePoOriginalSalesPrice.Text = "";

            txePoApprover.Text = "";
            dtePoApprovalDate.EditValue = DBNull.Value;
            sluePoOIDBillto.EditValue = "";
            txePoAddress.Text = "";
            txePoTelephoneNo.Text = "";
            sluePoPaymentTerms.EditValue = "";
            gluePoOIDCURR.EditValue = "";
            txePoRemark.Text = "";
            txePoAllocationOrderNumber.Text = "";

            gcEntryPO.Enabled = true;
            sbPoClearTable.Enabled = true;
            for (int i = gvEntryPO.RowCount - 1; i >= 0; i--)
                gvEntryPO.DeleteRow(i);
            gvEntryPO.OptionsView.ColumnAutoWidth = false;
            gvEntryPO.BestFitColumns();
            //gvEntryPO.Columns["ProductionPlanID"].Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;


            pbcPOSave.Properties.Step = 1;
            pbcPOSave.Properties.PercentView = true;
            pbcPOSave.Properties.Minimum = 0;
            pbcPOSave.EditValue = 0;
            lciPregressPOSave.Visibility = LayoutVisibility.Never;

            txePoOrderNo.Focus();
        }

        private void SetNewDO()
        {
            //*** PO ***
            glueDoDocStatus.EditValue = 0;

            //*** DO ***
            slueDoOrderNo.EditValue = "";
            speDoRevisionNo.Value = 0;
            txeDoDONo.Text = "";
            slueDoDONo.EditValue = "";

            slueDoItemCode.EditValue = "";
            txeDoItemName.Text = "";
            glueDoTransportationMethod.EditValue = "";
            slueDoPortCode.EditValue = "";
            txeDoPortName.Text = "";
            txeDoWHName.Text = "";
            txeDoAddress.Text = "";
            txeDoTelephone.Text = "";
            txeDoPersonInCharge.Text = "";
            glueDoIncoterms.EditValue = "1";
            txeDoForwarder.Text = "";
            slueDoOIDVEND.EditValue = "";
            dteDoETAWH.EditValue = DBNull.Value;
            dteDoContractedETD.EditValue = DBNull.Value;

            txeCREATE.Text = "0";

            if (Convert.ToInt32(DateTime.Now.ToString("yyyy")) > 2500)
                txeDATE.Text = (Convert.ToInt32(DateTime.Now.ToString("yyyy")) - 543).ToString() + DateTime.Now.ToString("-MM-dd HH:mm:ss");
            else
                txeDATE.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");


            gcEntryDO.Enabled = true;
            sbDoClearTable.Enabled = true;
            for (int i = gvEntryDO.RowCount - 1; i >= 0; i--)
                gvEntryDO.DeleteRow(i);
            gvEntryDO.OptionsView.ColumnAutoWidth = false;
            gvEntryDO.BestFitColumns();
            //gvEntryDO.Columns["ProductionPlanID"].Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;

            pbcDOSave.Properties.Step = 1;
            pbcDOSave.Properties.PercentView = true;
            pbcDOSave.Properties.Minimum = 0;
            pbcDOSave.EditValue = 0;
            lciPregressDOSave.Visibility = LayoutVisibility.Never;

            txeDoDONo.Focus();
            ////Tab Import
            //txePoFilePath.Text = "";
            //spsPO.CloseCellEditor(DevExpress.XtraSpreadsheet.CellEditorEnterValueMode.Default);
            //spsPO.CreateNewDocument();
            //cbePoSheet.Properties.Items.Clear();
            //cbePoSheet.Text = "";

            //txeDoFilePath.Text = "";
            //spsDO.CloseCellEditor(DevExpress.XtraSpreadsheet.CellEditorEnterValueMode.Default);
            //spsDO.CreateNewDocument();
            //cbeDoSheet.Properties.Items.Clear();
            //cbeDoSheet.Text = "";
        }

        private void NewData()
        {
            //Tab Entry
            SetNewPO();
            SetNewDO();

            txePoFilePath.Text = "";
            spsPO.CloseCellEditor(DevExpress.XtraSpreadsheet.CellEditorEnterValueMode.Default);
            spsPO.CreateNewDocument();
            cbePoSheet.Properties.Items.Clear();
            cbePoSheet.Text = "";

            lciPregressPOSave.Visibility = LayoutVisibility.Never;
            pbcPOSave.Properties.Step = 1;
            pbcPOSave.Properties.PercentView = true;
            pbcPOSave.Properties.Maximum = 100;
            pbcPOSave.Properties.Minimum = 0;
            pbcPOSave.EditValue = 0;


            txeDoFilePath.Text = "";
            spsDO.CloseCellEditor(DevExpress.XtraSpreadsheet.CellEditorEnterValueMode.Default);
            spsDO.CreateNewDocument();
            cbeDoSheet.Properties.Items.Clear();
            cbeDoSheet.Text = "";

            lciPregressDOSave.Visibility = LayoutVisibility.Never;
            pbcDOSave.Properties.Step = 1;
            pbcDOSave.Properties.PercentView = true;
            pbcDOSave.Properties.Maximum = 100;
            pbcDOSave.Properties.Minimum = 0;
            pbcDOSave.EditValue = 0;

            txePoOrderNo.Focus();

        }

        private void bbiNew_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            LoadData();
            NewData();
        }

        private int FindStatus(string Status)
        {
            int retStatus = 0;
            Status = Status.ToUpper().Trim().Replace("'", "");
            if (Status != "")
            {
                switch (Status)
                {
                    case "NEW":
                        retStatus = 0;
                        break;
                    case "REVISE":
                        retStatus = 1;
                        break;
                    case "CHANGE":
                        retStatus = 2;
                        break;
                    case "CANCEL":
                        retStatus = 3;
                        break;
                    case "FINISHED":
                        retStatus = 9;
                        break;
                    default:
                        retStatus = 0;
                        break;
                }
            }
            return retStatus;
        }

        private int FindMethod(string Status)
        {
            int retStatus = 0;
            Status = Status.ToUpper().Trim().Replace("'", "");
            if (Status != "")
            {
                switch (Status)
                {
                    case "SHIP":
                        retStatus = 0;
                        break;
                    case "AIR":
                        retStatus = 1;
                        break;
                    default:
                        retStatus = 0;
                        break;
                }
            }
            return retStatus;
        }

        private void bbiSave_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {

            if (tabbedControlGroup3.SelectedTabPage == layoutControlGroup5) //Entry
            {
                if (tabbedControlGroup4.SelectedTabPage == layoutControlGroup8) //PO
                {
                    gvEntryPO.CloseEditor();
                    gvEntryPO.UpdateCurrentRow();

                    if (gvEntryPO.HasColumnErrors == true)
                    {
                        FUNC.msgError("Can not save. Because found error in table. Please check.");
                    }
                    else if (CountErrorPO() > 0)
                    {
                        FUNC.msgWarning("Please enter color, size and qty order in the table.");
                    }
                    else
                    {
                        bool chkPO = true;
                        if (gluePoDocumentStatus.EditValue.ToString() == "0") //New
                        {
                            if (txePoOrderNo.Text.Trim() == "")
                            {
                                chkPO = false;
                                txePoOrderNo.Focus();
                                FUNC.msgWarning("Please input PO Order No.");
                            }
                        }
                        else
                        {
                            if (sluePoOrderNo.Text.Trim() == "")
                            {
                                chkPO = false;
                                sluePoOrderNo.Focus();
                                FUNC.msgWarning("Please select PO Order No.");
                            }
                        }

                        if (chkPO == true)
                        {
                            if (gvEntryPO.RowCount == 1)
                            {
                                FUNC.msgWarning("Please input data.");
                            }
                            else
                            {
                                //string strCREATE = "0";
                                string strCREATE = txeCREATE.Text.Trim() != "" ? txeCREATE.Text.Trim() : "0";

                                //Form Input
                                if (gluePoDocumentStatus.Text.Trim() == "")
                                {
                                    FUNC.msgWarning("Please select document status.");
                                    gluePoDocumentStatus.Focus();
                                }
                                else if (sluePoOIDCUST.Text.Trim() == "")
                                {
                                    FUNC.msgWarning("Please select customer.");
                                    sluePoOIDCUST.Focus();
                                }
                                else if (sluePoItemCode.Text.Trim() == "")
                                {
                                    FUNC.msgWarning("Please select item code.");
                                    sluePoItemCode.Focus();
                                }
                                else if (gvEntryPO.RowCount == 1)
                                {
                                    FUNC.msgWarning("Please enter color, size and qty order in the table.");
                                    gvEntryPO.Focus();
                                }
                                else
                                {
                                    bool chkDup = false;
                                    string msgACT = "";
                                    if (gluePoDocumentStatus.EditValue.ToString() == "0") //New
                                    {
                                        chkDup = chkDuplicatePO(txePoOrderNo.Text);
                                        msgACT = "Save New #PO";
                                    }
                                    else if (gluePoDocumentStatus.EditValue.ToString() == "1") //Revise
                                    {
                                        chkDup = chkDuplicatePO(sluePoOrderNo.Text);
                                        msgACT = "Save Revise #PO";
                                    }
                                    else if (gluePoDocumentStatus.EditValue.ToString() == "2") //Change
                                    {
                                        chkDup = chkDuplicatePO(sluePoOrderNo.Text);
                                        msgACT = "Save Change #PO";
                                    }
                                    else if (gluePoDocumentStatus.EditValue.ToString() == "3") //Cancel
                                    {
                                        chkDup = chkDuplicatePO(sluePoOrderNo.Text);
                                        msgACT = "Save Cancel #PO";
                                    }
                                    else if (gluePoDocumentStatus.EditValue.ToString() == "9") //Finished
                                    {
                                        chkDup = chkDuplicatePO(sluePoOrderNo.Text);
                                        msgACT = "Save Finished #PO";
                                    }

                                    if (chkDup == true)
                                    {
                                        if (FUNC.msgQuiz("Confirm " + msgACT + " ?") == true)
                                        {
                                            string PONO = gluePoDocumentStatus.EditValue.ToString() == "0" ? txePoOrderNo.Text.ToUpper().Trim() : sluePoOrderNo.Text.Trim();
                                            string RevisionNo = spePoRevisionNo.Value.ToString() != "" ? spePoRevisionNo.Value.ToString() : "0";
                                            string LotNo = txePoLot.Text.ToUpper().Trim();
                                            string RevisedDate = dtePoRevisedDate.Text.Trim() != "" ? "'" + Convert.ToDateTime(dtePoRevisedDate.EditValue.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL";
                                            string Status = gluePoDocumentStatus.Text.Trim() != "" ? gluePoDocumentStatus.EditValue.ToString() : "0";
                                            string OrderPlanNo = txePoOrderPlanNumber.Text.Trim();
                                            string CUSTID = sluePoOIDCUST.Text.Trim() != "" ? sluePoOIDCUST.EditValue.ToString() : "0";
                                            string Season = spePoSeason.Value.ToString() + gluePoSeason.EditValue.ToString();
                                            string Unit = gluePoBusinessUnit.Text.Trim() != "" ? gluePoBusinessUnit.EditValue.ToString() : "";
                                            string ItemCode = sluePoItemCode.Text.Trim() != "" ? sluePoItemCode.EditValue.ToString() : "0";
                                            string Remark = txePoRemark.Text.Trim();
                                            string Allocation = txePoAllocationOrderNumber.Text.Trim();
                                            string SalesPrice = txePoOriginalSalesPrice.Text.Trim();
                                            string Approver = txePoApprover.Text.Trim();
                                            string ApprovalDate = dtePoApprovalDate.Text.Trim() != "" ? "'" + Convert.ToDateTime(dtePoApprovalDate.EditValue.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL";
                                            string BillTo = sluePoOIDBillto.Text.Trim() != "" ? sluePoOIDBillto.EditValue.ToString() : "0";
                                            string PaymentTerm = sluePoPaymentTerms.Text.Trim() != "" ? sluePoPaymentTerms.EditValue.ToString() : "";
                                            string Currency = gluePoOIDCURR.Text.Trim() != "" ? gluePoOIDCURR.EditValue.ToString() : "0";
                                            //string POID = txePoOID.Text.Trim();

                                            StringBuilder sbSQL = new StringBuilder();
                                            if (gvEntryPO.RowCount > 1)
                                            {
                                                DataTable dtPO = (DataTable)gcEntryPO.DataSource;
                                                if (dtPO.Rows.Count > 0)
                                                {
                                                    foreach (DataRow row in dtPO.Rows)
                                                    {
                                                        string OIDCOLOR = row["OIDCOLOR"].ToString() != "" ? row["OIDCOLOR"].ToString() : "0";
                                                        string OIDSIZE = row["OIDSIZE"].ToString() != "" ? row["OIDSIZE"].ToString() : "0";
                                                        string PatternDimensionCode = row["PatternDimensionCode"].ToString().ToUpper().Trim();
                                                        string SKUCode = row["SKUCode"].ToString().ToUpper().Trim();
                                                        string SampleCode = row["SampleCode"].ToString().ToUpper().Trim();
                                                        string OrderQtyPCS = row["OrderQtyPCS"].ToString() != "" ? row["OrderQtyPCS"].ToString() : "0";
                                                        string ID = row["OID"].ToString() != "" ? row["OID"].ToString() : "";


                                                        if (gluePoDocumentStatus.EditValue.ToString() == "0" || gluePoDocumentStatus.EditValue.ToString() == "1") //new & revise
                                                        {
                                                            sbSQL.Append("INSERT INTO COPO(OrderNo, Lot, RevisionNo, Season, BusinessUnit, OIDCUST, RevisedDate, DocumentStatus, OIDBillto, PaymentTerms, OIDCURR, OrderPlanNumber, ItemCode, OIDCOLOR, OIDSIZE, PatternDimensionCode, ");
                                                            sbSQL.Append("              SKUCode, SampleCode, OrderQtyPCS, Remark, AllocationOrderNumber, OriginalSalesPrice, Approver, ApprovalDate) ");
                                                            sbSQL.Append("  VALUES(N'" + PONO + "', N'" + LotNo + "', '" + RevisionNo + "', N'" + Season + "', N'" + Unit + "', '" + CUSTID + "', " + RevisedDate + ", '" + Status + "', '" + BillTo + "', N'" + PaymentTerm + "', '" + Currency + "', N'" + OrderPlanNo + "', '" + ItemCode + "', '" + OIDCOLOR + "', '" + OIDSIZE + "', N'" + PatternDimensionCode + "', N'" + SKUCode + "', N'" + SampleCode + "', '" + OrderQtyPCS + "', N'" + Remark + "', N'" + Allocation + "', N'" + SalesPrice + "', N'" + Approver + "', " + ApprovalDate + ")  ");
                                                        }
                                                        else if (gluePoDocumentStatus.EditValue.ToString() == "2") //change
                                                        {
                                                            if (ID == "")
                                                            {
                                                                sbSQL.Append("INSERT INTO COPO(OrderNo, Lot, RevisionNo, Season, BusinessUnit, OIDCUST, RevisedDate, DocumentStatus, OIDBillto, PaymentTerms, OIDCURR, OrderPlanNumber, ItemCode, OIDCOLOR, OIDSIZE, PatternDimensionCode, ");
                                                                sbSQL.Append("              SKUCode, SampleCode, OrderQtyPCS, Remark, AllocationOrderNumber, OriginalSalesPrice, Approver, ApprovalDate) ");
                                                                sbSQL.Append("  VALUES(N'" + PONO + "', N'" + LotNo + "', '" + RevisionNo + "', N'" + Season + "', N'" + Unit + "', '" + CUSTID + "', " + RevisedDate + ", '" + Status + "', '" + BillTo + "', N'" + PaymentTerm + "', '" + Currency + "', N'" + OrderPlanNo + "', '" + ItemCode + "', '" + OIDCOLOR + "', '" + OIDSIZE + "', N'" + PatternDimensionCode + "', N'" + SKUCode + "', N'" + SampleCode + "', '" + OrderQtyPCS + "', N'" + Remark + "', N'" + Allocation + "', N'" + SalesPrice + "', N'" + Approver + "', " + ApprovalDate + ")  ");
                                                            }
                                                            else
                                                            {
                                                                sbSQL.Append("UPDATE COPO SET ");
                                                                sbSQL.Append("  OrderNo=N'" + PONO + "', Lot=N'" + LotNo + "', RevisionNo='" + RevisionNo + "', Season=N'" + Season + "', BusinessUnit=N'" + Unit + "', OIDCUST='" + CUSTID + "', RevisedDate=" + RevisedDate + ", ");
                                                                sbSQL.Append("  DocumentStatus='" + Status + "', OIDBillto='" + BillTo + "', PaymentTerms=N'" + PaymentTerm + "', OIDCURR='" + Currency + "', OrderPlanNumber=N'" + OrderPlanNo + "', ItemCode='" + ItemCode + "', ");
                                                                sbSQL.Append("  OIDCOLOR='" + OIDCOLOR + "', OIDSIZE='" + OIDSIZE + "', PatternDimensionCode=N'" + PatternDimensionCode + "', SKUCode=N'" + SKUCode + "', SampleCode=N'" + SampleCode + "', OrderQtyPCSOld = OrderQtyPCS, OrderQtyPCS='" + OrderQtyPCS + "', ");
                                                                sbSQL.Append("  Remark=N'" + Remark + "', AllocationOrderNumber=N'" + Allocation + "', OriginalSalesPrice=N'" + SalesPrice + "', Approver=N'" + Approver + "', ApprovalDate=" + ApprovalDate + "  ");
                                                                sbSQL.Append("WHERE (OID='" + ID + "')  ");
                                                            }
                                                        }
                                                        else if (gluePoDocumentStatus.EditValue.ToString() == "3" || gluePoDocumentStatus.EditValue.ToString() == "9") //cancel & finished
                                                        {
                                                            sbSQL.Append("UPDATE COPO SET ");
                                                            sbSQL.Append("  DocumentStatus='" + Status + "' ");
                                                            sbSQL.Append("WHERE (OID='" + ID + "')  ");
                                                        }
                                                    }
                                                }
                                            }

                                            try
                                            {
                                                bool chkSAVE = new DBQuery(sbSQL).runSQL();
                                                if (chkSAVE == true)
                                                {
                                                    if (gluePoDocumentStatus.EditValue.ToString() == "0" || gluePoDocumentStatus.EditValue.ToString() == "1") //New & Revise
                                                    {
                                                        sbSQL.Clear();
                                                        sbSQL.Append("SELECT DISTINCT PO.OrderNo + '_' + CONVERT(NVARCHAR, PO.RevisionNo) AS POID, PO.OrderNo AS [PO. Order No.], PO.RevisionNo AS [Revision No.], ST.Status, PO.Season, CUS.Name AS Customer, IC.ItemName ");
                                                        sbSQL.Append("FROM   COPO AS PO INNER JOIN ");
                                                        sbSQL.Append("       Customer AS CUS ON PO.OIDCUST = CUS.OIDCUST INNER JOIN ");
                                                        sbSQL.Append("       ItemCustomer AS IC ON PO.ItemCode = IC.OIDCSITEM INNER JOIN ");
                                                        sbSQL.Append("       (" + sbSTATUS.ToString() + ") AS ST ON PO.DocumentStatus = ST.ID ");
                                                        sbSQL.Append("ORDER BY [PO. Order No.] ");
                                                        new ObjDevEx.setSearchLookUpEdit(slueDoOrderNo, sbSQL, "PO. Order No.", "POID").getData();
                                                        slueDoOrderNo.Properties.PopulateViewColumns();
                                                        slueDoOrderNo.Properties.View.Columns["POID"].Visible = false;
                                                        speDoRevisionNo.Value = 0;
                                                    }

                                                    SetNewPO();

                                                    FUNC.msgInfo(msgACT + " Complete.");
                                                }
                                            }
                                            catch (Exception)
                                            { }
                                        }
                                    }
                                    else
                                    {
                                        txePoOrderNo.Text = "";
                                        txePoOrderNo.Focus();
                                        FUNC.msgWarning("Duplicate PO Order No. !! Please Change.");
                                    }
                                }

                            }
                        }
                    }
                }
                else if (tabbedControlGroup4.SelectedTabPage == layoutControlGroup9) //DO
                {
                    gvEntryDO.CloseEditor();
                    gvEntryDO.UpdateCurrentRow();

                    if (gvEntryDO.HasColumnErrors == true)
                    {
                        FUNC.msgError("Can not save. Because found error in table. Please check.");
                    }
                    else
                    {
                        bool chkDO = true;
                        if (glueDoDocStatus.EditValue.ToString() == "0") //New
                        {
                            if (txeDoDONo.Text.Trim() == "")
                            {
                                chkDO = false;
                                txeDoDONo.Focus();
                                FUNC.msgWarning("Please input DO No.");
                            }
                        }
                        else
                        {
                            if (slueDoDONo.Text.Trim() == "")
                            {
                                chkDO = false;
                                slueDoDONo.Focus();
                                FUNC.msgWarning("Please select DO No.");
                            }
                        }

                        if (chkDO == true)
                        {
                            if (slueDoOrderNo.Text.Trim() == "")
                            {
                                slueDoOrderNo.Focus();
                                FUNC.msgWarning("Please select PO Order No.");
                            }
                            else if (glueDoTransportationMethod.Text.Trim() == "")
                            {
                                glueDoTransportationMethod.Focus();
                                FUNC.msgWarning("Please select transportation method.");
                            }
                            else if (slueDoPortCode.Text.Trim() == "")
                            {
                                slueDoPortCode.Focus();
                                FUNC.msgWarning("Please select Ship to port code.");
                            }
                            else
                            {
                                //string strCREATE = "0";
                                string strCREATE = txeCREATE.Text.Trim() != "" ? txeCREATE.Text.Trim() : "0";

                                bool chkDup = false;
                                string msgACT = "";
                                if (glueDoDocStatus.EditValue.ToString() == "0") //New
                                {
                                    chkDup = chkDuplicateDO(txeDoDONo.Text);
                                    msgACT = "Save New #DO";
                                }
                                else if (glueDoDocStatus.EditValue.ToString() == "1") //Revise
                                {
                                    chkDup = chkDuplicatePO(slueDoDONo.Text);
                                    msgACT = "Save Revise #DO";
                                }
                                else if (glueDoDocStatus.EditValue.ToString() == "2") //Change
                                {
                                    chkDup = chkDuplicatePO(slueDoDONo.Text);
                                    msgACT = "Save Change #DO";
                                }
                                else if (glueDoDocStatus.EditValue.ToString() == "3") //Cancel
                                {
                                    chkDup = chkDuplicatePO(slueDoDONo.Text);
                                    msgACT = "Save Cancel #DO";
                                }
                                else if (glueDoDocStatus.EditValue.ToString() == "9") //Finished
                                {
                                    chkDup = chkDuplicatePO(slueDoDONo.Text);
                                    msgACT = "Save Finished #DO";
                                }

                                if (chkDup == true)
                                {
                                    if (FUNC.msgQuiz("Confirm " + msgACT + " ?") == true)
                                    {
                                        string DONO = glueDoDocStatus.EditValue.ToString() == "0" ? txeDoDONo.Text.ToUpper().Trim() : slueDoDONo.Text.Trim();

                                        string PONO = slueDoOrderNo.Text.Trim() != "" ? slueDoOrderNo.Text.ToUpper().Trim() : "";
                                        string RevisionNo = speDoRevisionNo.Value.ToString() != "" ? speDoRevisionNo.Value.ToString() : "0";
                                        string ItemCode = slueDoItemCode.Text.Trim() != "" ? slueDoItemCode.EditValue.ToString() : "";

                                        string TransportationMethod = glueDoTransportationMethod.Text.Trim() != "" ? glueDoTransportationMethod.EditValue.ToString() : "";
                                        string PortCode = slueDoPortCode.Text.Trim() != "" ? slueDoPortCode.EditValue.ToString() : "";

                                        string Incoterms = glueDoIncoterms.EditValue.ToString();
                                        string Forwarder = txeDoForwarder.Text.Trim();

                                        string OIDVEND = slueDoOIDVEND.Text.Trim() != "" ? slueDoOIDVEND.EditValue.ToString() : "";

                                        string ETAWH = dteDoETAWH.Text.Trim() != "" ? "'" + Convert.ToDateTime(dteDoETAWH.EditValue.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL";
                                        string ContractedETD = dteDoContractedETD.Text.Trim() != "" ? "'" + Convert.ToDateTime(dteDoContractedETD.EditValue.ToString()).ToString("yyyy-MM-dd") + "'" : "NULL";

                                        string Status = glueDoDocStatus.Text.Trim() != "" ? glueDoDocStatus.EditValue.ToString() : "0";

                                        StringBuilder sbSQL = new StringBuilder();
                                        if (gvEntryDO.RowCount > 1)
                                        {
                                            DataTable dtDO = (DataTable)gcEntryDO.DataSource;
                                            if (dtDO.Rows.Count > 0)
                                            {
                                                string chkPO = PONO;
                                                string chkREVISE = RevisionNo;
                                                string[] arrChk = new DBQuery("SELECT TOP(1) OrderNo, RevisionNo FROM CODO WHERE (DONo = N'" + DONO + "')").getMultipleValue();
                                                if (arrChk.Length > 0)
                                                {
                                                    chkPO = arrChk[0];
                                                    chkREVISE = arrChk[1];
                                                }
                                                bool chkChangePO = false;
                                                if (chkPO != PONO || chkREVISE != RevisionNo)
                                                {
                                                    sbSQL.Append("DELETE FROM CODO WHERE (DONo = N'" + DONO + "') ");
                                                    chkChangePO = true;
                                                }

                                                foreach (DataRow row in dtDO.Rows)
                                                {
                                                    string OIDCOLOR = row["OIDCOLOR"].ToString() != "" ? row["OIDCOLOR"].ToString() : "0";
                                                    string OIDSIZE = row["OIDSIZE"].ToString() != "" ? row["OIDSIZE"].ToString() : "0";
                                                    string PatternDimensionCode = row["PatternDimensionCode"].ToString().ToUpper().Trim();
                                                    string SetCode = row["SetCode"].ToString().ToUpper().Trim();
                                                    string QtyBox = row["QtyBox"].ToString() != "" ? row["QtyBox"].ToString() : "0";
                                                    string QtySet = row["QtySet"].ToString() != "" ? row["QtySet"].ToString() : "0";
                                                    string PickUnit = row["PickUnit"].ToString() != "" ? row["PickUnit"].ToString() : "0";
                                                    string ID = row["OID"].ToString() != "" ? row["OID"].ToString() : "";

                                                    if (ID == "" && QtyBox == "0")
                                                        continue;
                                                    else
                                                    {
                                                        if (QtyBox == "0")
                                                        {
                                                            sbSQL.Append("DELETE FROM CODO WHRE (OIDDO='" + ID + "') ");
                                                        }
                                                        else
                                                        {
                                                            if (glueDoDocStatus.EditValue.ToString() == "0" || glueDoDocStatus.EditValue.ToString() == "1") //new & revise
                                                            {
                                                                sbSQL.Append("INSERT INTO CODO(OrderNo, RevisionNo, DONo, DocStatus, ItemCode, ETAWH, TransportationMethod, PortCode, Incoterms, ContractedETD, Forwarder, OIDVEND, SetCode, QuantityBox, ");
                                                                sbSQL.Append("          OIDCOLOR, OIDSIZE, PatternDimensionCode, QtyperSet, PickingUnit, UpdatedBy, UpdatedDate) ");
                                                                sbSQL.Append("  VALUES(N'" + PONO + "', '" + RevisionNo + "', N'" + DONO + "', '" + Status + "', '" + ItemCode + "', " + ETAWH + ", '" + TransportationMethod + "', '" + PortCode + "', '" + Incoterms + "', " + ContractedETD + ", N'" + Forwarder + "', '" + OIDVEND + "', N'" + SetCode + "', '" + QtyBox + "', '" + OIDCOLOR + "', '" + OIDSIZE + "', N'" + PatternDimensionCode + "', '" + QtySet + "', '" + PickUnit + "', '" + strCREATE + "', GETDATE())  ");
                                                            }
                                                            else if (glueDoDocStatus.EditValue.ToString() == "2") //change
                                                            {
                                                                if (chkChangePO == true)
                                                                {
                                                                    sbSQL.Append("INSERT INTO CODO(OrderNo, RevisionNo, DONo, DocStatus, ItemCode, ETAWH, TransportationMethod, PortCode, Incoterms, ContractedETD, Forwarder, OIDVEND, SetCode, QuantityBox, ");
                                                                    sbSQL.Append("          OIDCOLOR, OIDSIZE, PatternDimensionCode, QtyperSet, PickingUnit, UpdatedBy, UpdatedDate) ");
                                                                    sbSQL.Append("  VALUES(N'" + PONO + "', '" + RevisionNo + "', N'" + DONO + "', '" + Status + "', '" + ItemCode + "', " + ETAWH + ", '" + TransportationMethod + "', '" + PortCode + "', '" + Incoterms + "', " + ContractedETD + ", N'" + Forwarder + "', '" + OIDVEND + "', N'" + SetCode + "', '" + QtyBox + "', '" + OIDCOLOR + "', '" + OIDSIZE + "', N'" + PatternDimensionCode + "', '" + QtySet + "', '" + PickUnit + "', '" + strCREATE + "', GETDATE())  ");
                                                                }
                                                                else
                                                                {
                                                                    if (ID == "")
                                                                    {
                                                                        sbSQL.Append("INSERT INTO CODO(OrderNo, RevisionNo, DONo, DocStatus, ItemCode, ETAWH, TransportationMethod, PortCode, Incoterms, ContractedETD, Forwarder, OIDVEND, SetCode, QuantityBox, ");
                                                                        sbSQL.Append("          OIDCOLOR, OIDSIZE, PatternDimensionCode, QtyperSet, PickingUnit, UpdatedBy, UpdatedDate) ");
                                                                        sbSQL.Append("  VALUES(N'" + PONO + "', '" + RevisionNo + "', N'" + DONO + "', '" + Status + "', '" + ItemCode + "', " + ETAWH + ", '" + TransportationMethod + "', '" + PortCode + "', '" + Incoterms + "', " + ContractedETD + ", N'" + Forwarder + "', '" + OIDVEND + "', N'" + SetCode + "', '" + QtyBox + "', '" + OIDCOLOR + "', '" + OIDSIZE + "', N'" + PatternDimensionCode + "', '" + QtySet + "', '" + PickUnit + "', '" + strCREATE + "', GETDATE())  ");
                                                                    }
                                                                    else
                                                                    {
                                                                        sbSQL.Append("UPDATE CODO SET ");
                                                                        sbSQL.Append("  OrderNo=N'" + PONO + "', RevisionNo='" + RevisionNo + "', DONo=N'" + DONO + "', DocStatus='" + Status + "', ItemCode='" + ItemCode + "', ");
                                                                        sbSQL.Append("  ETAWHOld=ETAWH, ETAWH=" + ETAWH + ", TransportationMethod='" + TransportationMethod + "', PortCode='" + PortCode + "', Incoterms='" + Incoterms + "', ");
                                                                        sbSQL.Append("  ContractedETDOld=ContractedETD, ContractedETD=" + ContractedETD + ", Forwarder=N'" + Forwarder + "', OIDVEND='" + OIDVEND + "', ");
                                                                        sbSQL.Append("  SetCode=N'" + SetCode + "', QuantityBoxOld=QuantityBox, QuantityBox='" + QtyBox + "', OIDCOLOR='" + OIDCOLOR + "', OIDSIZE='" + OIDSIZE + "', ");
                                                                        sbSQL.Append("  PatternDimensionCode=N'" + PatternDimensionCode + "', QtyperSet='" + QtySet + "', PickingUnit='" + PickUnit + "', ");
                                                                        sbSQL.Append("  UpdatedBy='" + strCREATE + "', UpdatedDate=GETDATE()  ");
                                                                        sbSQL.Append("WHERE (OIDDO='" + ID + "')  ");
                                                                    }
                                                                }
                                                            }
                                                            else if (glueDoDocStatus.EditValue.ToString() == "3" || glueDoDocStatus.EditValue.ToString() == "9") //cancel & finished
                                                            {
                                                                sbSQL.Append("UPDATE CODO SET ");
                                                                sbSQL.Append("  DocStatus='" + Status + "' ");
                                                                sbSQL.Append("WHERE (OIDDO='" + ID + "')  ");
                                                            }
                                                        }
                                                      // MessageBox.Show(sbSQL.ToString());
                                                    }
                                                }
                                            }
                                        }

                                        try
                                        {
                                            bool chkSAVE = new DBQuery(sbSQL).runSQL();
                                            if (chkSAVE == true)
                                            {
                                                SetNewDO();
                                                txeDoDONo.Focus();
                                                FUNC.msgInfo(msgACT + " Complete.");
                                            }
                                        }
                                        catch (Exception)
                                        { }
                                    }
                                }
                                else
                                {
                                    txeDoDONo.Text = "";
                                    txeDoDONo.Focus();
                                    FUNC.msgWarning("Duplicate DO No. !! Please Change.");
                                }

                            }
                        }
                    }
                }
            }
            else if (tabbedControlGroup3.SelectedTabPage == layoutControlGroup12) //Import File
            {
                if (tabbedControlGroup5.SelectedTabPage == layoutControlGroup13) //PO
                {
                    if (txePoFilePath.Text.Trim() == "")
                    {
                        FUNC.msgWarning("Please select excel file.");
                        txePoFilePath.Focus();
                    }
                    else if (cbePoSheet.Text.Trim() == "")
                    {
                        FUNC.msgWarning("Please select excel sheet.");
                        cbePoSheet.Focus();
                    }
                    else
                    {
                        if (FUNC.msgQuiz("Confirm save excel file import data ?") == true)
                        {
                            StringBuilder sbSQL = new StringBuilder();

                            bool chkSAVE = false;

                            IWorkbook workbook = spsPO.Document;
                            Worksheet WSHEETPO = workbook.Worksheets[0];

                            lciPregressPOSave.Visibility = LayoutVisibility.Always;
                            pbcPOSave.Properties.Step = 1;
                            pbcPOSave.Properties.PercentView = true;
                            pbcPOSave.Properties.Maximum = WSHEETPO.GetDataRange().RowCount;
                            pbcPOSave.Properties.Minimum = 0;
                            pbcPOSave.EditValue = 0;

                            string Customer = "";
                            string OIDCUST = "";

                            string Billto = "";
                            string OIDBillto = "";

                            //string PaymentTerms = "";
                            //string OIDPaymentTerms = "";

                            string Currency = "";
                            string OIDCURR = "";

                            string ColorCode = "";
                            string ColorName = "";
                            string OIDCOLOR = "";

                            string SizeCode = "";
                            string SizeName = "";
                            string OIDSIZE = "";

                            string ItemCode = "";
                            string OIDITEM = "";

                            for (int i = 3; i < WSHEETPO.GetDataRange().RowCount; i++)
                            {
                                string OrderNo = WSHEETPO.Rows[i][0].DisplayText.ToString().Trim();

                                if (OrderNo != "")
                                {
                                    string Lot = "";
                                    string RevisionNo = WSHEETPO.Rows[i][1].DisplayText.ToString().Trim();
                                    RevisionNo = IsNumeric(RevisionNo) == true ? RevisionNo : "0";
                                    string Season = WSHEETPO.Rows[i][2].DisplayText.ToString().Trim() + WSHEETPO.Rows[i][3].DisplayText.ToString().Trim().ToUpper().Trim();
                                    Season = Season.Length > 6 ? Season.Substring(0, 6) : Season;
                                    string BusinessUnit = WSHEETPO.Rows[i][5].DisplayText.ToString().ToUpper().Trim();
                                    if (BusinessUnit.Length > 30)
                                        BusinessUnit = BusinessUnit.Substring(0, 30);
                                    string strCustomer = WSHEETPO.Rows[i][6].DisplayText.ToString().ToUpper().Trim().Replace("'", "''");
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
                                        OIDCUST = new DBQuery(sbCUST).getString();
                                    }
                                    string RevisedDate = WSHEETPO.Rows[i][8].DisplayText.ToString().Trim();
                                    RevisedDate = RevisedDate != "" ? "'" + Convert.ToDateTime(RevisedDate).ToString("yyyy-MM-dd") + "'" : "NULL";
                                    string documentStatus = WSHEETPO.Rows[i][9].DisplayText.ToString().Trim();
                                    int DocumentStatus = FindStatus(documentStatus);
                                    string strBillto = WSHEETPO.Rows[i][17].DisplayText.ToString().Trim().Replace("'", "''");
                                    if (Billto != strBillto.Replace(" ", "").Replace(".", "").Replace(",", ""))
                                    {
                                        Billto = strBillto.Replace(" ", "").Replace(".", "").Replace(",", "");
                                        string VendCode = Billto.Length > 20 ? Billto.Substring(0, 20) : Billto;
                                        StringBuilder sbCUST = new StringBuilder();
                                        sbCUST.Append("IF NOT EXISTS(SELECT OIDVEND FROM Vendor WHERE (REPLACE(REPLACE(REPLACE(Name, ' ', ''), '.', ''), ',', '') LIKE N'" + Billto + "%')) ");
                                        sbCUST.Append(" BEGIN ");
                                        sbCUST.Append("   INSERT INTO Vendor(Code, Name, VendorType) VALUES(N'" + VendCode + "', N'" + strBillto + "', 6) ");
                                        sbCUST.Append(" END ");
                                        sbCUST.Append("SELECT TOP(1) OIDVEND FROM Vendor WHERE (REPLACE(REPLACE(REPLACE(Name, ' ', ''), '.', ''), ',', '') LIKE N'" + Billto + "%') ");
                                        OIDBillto = new DBQuery(sbCUST).getString();
                                    }

                                    string strPaymentTerms = WSHEETPO.Rows[i][23].DisplayText.ToString().ToUpper().Trim().Replace("'", "''");
                                    //if (PaymentTerms != strPaymentTerms.Replace(" ", "").Replace(".", "").Replace(",", "").Replace("/", ""))
                                    //{
                                    //    PaymentTerms = strPaymentTerms.Replace(" ", "").Replace(".", "").Replace(",", "").Replace("/", "");
                                    //    string TermsName = strPaymentTerms.Length > 50 ? strPaymentTerms.Substring(0, 50) : strPaymentTerms;
                                    //    TermsName = TermsName.Replace("'", "''");
                                    //    StringBuilder sbTERMS = new StringBuilder();
                                    //    sbTERMS.Append("IF NOT EXISTS(SELECT OIDPayment FROM PaymentTerm WHERE (REPLACE(REPLACE(REPLACE(REPLACE(Name, ' ', ''), '.', ''), ',', ''), '/', '') LIKE N'" + PaymentTerms + "%') ");
                                    //    sbTERMS.Append(" BEGIN ");
                                    //    sbTERMS.Append("   INSERT INTO PaymentTerm(Name) VALUES(N'" + TermsName + "') ");
                                    //    sbTERMS.Append(" END ");
                                    //    sbTERMS.Append("SELECT TOP(1) OIDPayment FROM PaymentTerm WHERE (REPLACE(REPLACE(REPLACE(REPLACE(Name, ' ', ''), '.', ''), ',', ''), '/', '') LIKE N'" + PaymentTerms + "%') ");
                                    //    OIDPaymentTerms = new DBQuery(sbTERMS).getString();
                                    //}

                                    string strCurrency = WSHEETPO.Rows[i][24].DisplayText.ToString().ToUpper().Trim().Replace("'", "''");
                                    if (Currency != strCurrency.Replace(" ", "").Replace(".", "").Replace(",", "").Replace("/", ""))
                                    {
                                        Currency = strCurrency.Replace(" ", "").Replace(".", "").Replace(",", "").Replace("/", "");
                                        string CurrName = strCurrency.Length > 10 ? strCurrency.Substring(0, 10) : strCurrency;
                                        CurrName = CurrName.Replace("'", "''");
                                        StringBuilder sbCURR = new StringBuilder();
                                        sbCURR.Append("IF NOT EXISTS(SELECT OIDCURR FROM Currency WHERE (REPLACE(REPLACE(REPLACE(REPLACE(Currency, ' ', ''), '.', ''), ',', ''), '/', '') LIKE N'" + Currency + "%')) ");
                                        sbCURR.Append(" BEGIN ");
                                        sbCURR.Append("   INSERT INTO Currency(Currency) VALUES(N'" + CurrName + "') ");
                                        sbCURR.Append(" END ");
                                        sbCURR.Append("SELECT TOP(1) OIDCURR FROM Currency WHERE (REPLACE(REPLACE(REPLACE(REPLACE(Currency, ' ', ''), '.', ''), ',', ''), '/', '') LIKE N'" + Currency + "%') ");
                                        OIDCURR = new DBQuery(sbCURR).getString();
                                    }

                                    string OrderPlanNumber = WSHEETPO.Rows[i][26].DisplayText.ToString().ToUpper().Trim().Replace("'", "''");
                                    if (OrderPlanNumber.Length > 30)
                                        OrderPlanNumber = OrderPlanNumber.Substring(0, 30);
                                    string TrimOrderNo = "";

                                    string strItemCode = WSHEETPO.Rows[i][27].DisplayText.ToString().Trim().Replace("'", "''");
                                    string strItemName = WSHEETPO.Rows[i][28].DisplayText.ToString().Trim().Replace("'", "''");
                                    string strStyleNo = WSHEETPO.Rows[i][29].DisplayText.ToString().Trim().Replace("'", "''");

                                    string strStyle = strStyleNo.Replace(Regex.Match(strStyleNo, @"\d+([,\.]\d+)?").Value, "");
                                    StringBuilder sbSTYLE = new StringBuilder();
                                    sbSTYLE.Append("IF NOT EXISTS(SELECT OIDSTYLE FROM ProductStyle WHERE StyleName = N'" + strStyle + "') ");
                                    sbSTYLE.Append("  BEGIN ");
                                    sbSTYLE.Append("       INSERT INTO ProductStyle(StyleName) VALUES(N'" + strStyle + "') ");
                                    sbSTYLE.Append("  END ");
                                    sbSTYLE.Append("SELECT OIDSTYLE FROM ProductStyle WHERE(StyleName = N'" + strStyle + "') ");
                                    string OIDSTYLE = new DBQuery(sbSTYLE).getString();

                                    if (ItemCode != strItemCode)
                                    {
                                        ItemCode = strItemCode;
                                        strItemCode = strItemCode.Length > 20 ? strItemCode.Substring(0, 20) : strItemCode;
                                        strStyleNo = strStyleNo.Length > 10 ? strStyleNo.Substring(0, 10) : strStyleNo;
                                        StringBuilder sbITEM = new StringBuilder();
                                        string FabricWidth = WSHEETPO.Rows[i][38].DisplayText.ToString().Trim().Replace("'", "''");
                                        string FBComposition = WSHEETPO.Rows[i][40].DisplayText.ToString().Trim().Replace("'", "''");

                                        sbITEM.Append("IF NOT EXISTS(SELECT OIDCSITEM FROM ItemCustomer WHERE (OIDCUST='" + OIDCUST + "') AND (ItemCode = N'" + strItemCode + "')) ");
                                        sbITEM.Append(" BEGIN ");
                                        sbITEM.Append("   INSERT INTO ItemCustomer(OIDCUST, ItemCode, ItemName, OIDSTYLE, Season, FabricWidth, FBComposition, StyleNo) VALUES('" + OIDCUST + "', N'" + strItemCode + "', N'" + strItemName + "', '" + OIDSTYLE + "', N'" + Season + "', N'" + FabricWidth + "', N'" + FBComposition + "', N'" + strStyleNo + "') ");
                                        sbITEM.Append(" END ");
                                        sbITEM.Append("ELSE ");
                                        sbITEM.Append(" BEGIN ");
                                        sbITEM.Append("   UPDATE ItemCustomer SET  ");
                                        sbITEM.Append("     ItemName=N'" + strItemName + "', OIDSTYLE='" + OIDSTYLE + "', Season=N'" + Season + "', FabricWidth=N'" + FabricWidth + "', FBComposition=N'" + FBComposition + "', StyleNo=N'" + strStyleNo + "'  ");
                                        sbITEM.Append("   WHERE (OIDCUST='" + OIDCUST + "') AND (ItemCode = N'" + strItemCode + "')  ");
                                        sbITEM.Append(" END ");
                                        sbITEM.Append("SELECT TOP(1) OIDCSITEM FROM ItemCustomer WHERE (OIDCUST='" + OIDCUST + "') AND (ItemCode = N'" + strItemCode + "') ");
                                        OIDITEM = new DBQuery(sbITEM).getString();
                                    }

                                    string strColorCode = WSHEETPO.Rows[i][47].DisplayText.ToString().ToUpper().Trim().Replace("'", "''");
                                    string strColorName = WSHEETPO.Rows[i][48].DisplayText.ToString().ToUpper().Trim().Replace("'", "''");
                                    if (ColorCode != strColorCode.Replace(" ", "").Replace(".", "").Replace(",", "").Replace("/", ""))
                                    {
                                        ColorCode = strColorCode.Replace(" ", "").Replace(".", "").Replace(",", "").Replace("/", "");
                                        string CCode = strColorCode.Length > 20 ? strColorCode.Substring(0, 20) : strColorCode;
                                        CCode = CCode.Replace("'", "''");

                                        ColorName = strColorName.Replace(" ", "").Replace(".", "").Replace(",", "").Replace("/", "");
                                        string CName = strColorName.Length > 50 ? strColorName.Substring(0, 50) : strColorName;
                                        CName = CName.Replace("'", "''");

                                        StringBuilder sbCOLOR = new StringBuilder();
                                        sbCOLOR.Append("IF NOT EXISTS(SELECT OIDCOLOR FROM ProductColor WHERE (REPLACE(REPLACE(REPLACE(REPLACE(ColorNo, ' ', ''), '.', ''), ',', ''), '/', '') LIKE N'" + CCode + "%')) ");
                                        sbCOLOR.Append(" BEGIN ");
                                        sbCOLOR.Append("   INSERT INTO ProductColor(ColorNo, ColorName) VALUES(N'" + CCode + "', N'" + CName + "') ");
                                        sbCOLOR.Append(" END ");
                                        sbCOLOR.Append("SELECT TOP(1) OIDCOLOR FROM ProductColor WHERE (REPLACE(REPLACE(REPLACE(REPLACE(ColorNo, ' ', ''), '.', ''), ',', ''), '/', '') LIKE N'" + CCode + "%') ");
                                        OIDCOLOR = new DBQuery(sbCOLOR).getString();
                                    }


                                    string strSizeCode = WSHEETPO.Rows[i][49].DisplayText.ToString().ToUpper().Trim().Replace("'", "''");
                                    string strSizeName = WSHEETPO.Rows[i][50].DisplayText.ToString().ToUpper().Trim().Replace("'", "''");
                                    if (SizeCode != strSizeCode.Replace(" ", "").Replace(".", "").Replace(",", "").Replace("/", ""))
                                    {
                                        SizeCode = strSizeCode.Replace(" ", "").Replace(".", "").Replace(",", "").Replace("/", "");
                                        string SCode = strSizeCode.Length > 10 ? strSizeCode.Substring(0, 10) : strSizeCode;
                                        SCode = SCode.Replace("'", "''");

                                        SizeName = strSizeName.Replace(" ", "").Replace(".", "").Replace(",", "").Replace("/", "");
                                        string SName = strSizeName.Length > 50 ? strSizeName.Substring(0, 50) : strSizeName;
                                        SName = SName.Replace("'", "''");

                                        StringBuilder sbCOLOR = new StringBuilder();
                                        sbCOLOR.Append("IF NOT EXISTS(SELECT OIDSIZE FROM ProductSize WHERE (REPLACE(REPLACE(REPLACE(REPLACE(SizeNo, ' ', ''), '.', ''), ',', ''), '/', '') LIKE N'" + SCode + "%')) ");
                                        sbCOLOR.Append(" BEGIN ");
                                        sbCOLOR.Append("   INSERT INTO ProductSize(SizeNo, SizeName) VALUES(N'" + SCode + "', N'" + SName + "') ");
                                        sbCOLOR.Append(" END ");
                                        sbCOLOR.Append("SELECT TOP(1) OIDSIZE FROM ProductSize WHERE (REPLACE(REPLACE(REPLACE(REPLACE(SizeNo, ' ', ''), '.', ''), ',', ''), '/', '') LIKE N'" + SCode + "%') ");
                                        OIDSIZE = new DBQuery(sbCOLOR).getString();
                                    }

                                    string PatternDimensionCode = WSHEETPO.Rows[i][51].DisplayText.ToString().ToUpper().Trim().Replace("'", "''");
                                    if (PatternDimensionCode.Length > 20)
                                        PatternDimensionCode = PatternDimensionCode.Substring(0, 20);
                                    string SKUCode = WSHEETPO.Rows[i][52].DisplayText.ToString().ToUpper().Trim().Replace("'", "''");
                                    if (SKUCode.Length > 30)
                                        SKUCode = SKUCode.Substring(0, 30);
                                    string SampleCode = WSHEETPO.Rows[i][54].DisplayText.ToString().ToUpper().Trim().Replace("'", "''");
                                    if (SampleCode.Length > 30)
                                        SampleCode = SampleCode.Substring(0, 30);
                                    string OrderQtyPCS = WSHEETPO.Rows[i][55].DisplayText.ToString().Trim().Replace("'", "''").Replace(",", "").Replace(" ", "");
                                    string Remark = WSHEETPO.Rows[i][56].DisplayText.ToString().Trim().Replace("'", "''");
                                    if (Remark.Length > 150)
                                        Remark = Remark.Substring(0, 150);
                                    string AllocationOrderNumber = WSHEETPO.Rows[i][57].DisplayText.ToString().ToUpper().Trim().Replace("'", "''");
                                    if (AllocationOrderNumber.Length > 50)
                                        AllocationOrderNumber = AllocationOrderNumber.Substring(0, 50);
                                    string OriginalSalesPrice = WSHEETPO.Rows[i][58].DisplayText.ToString().ToUpper().Trim().Replace("'", "''");
                                    if (OriginalSalesPrice.Length > 150)
                                        OriginalSalesPrice = OriginalSalesPrice.Substring(0, 150);
                                    string CostSheetNo = "";
                                    string Approver = WSHEETPO.Rows[i][10].DisplayText.ToString().Trim().Replace("'", "''");
                                    if (Approver.Length > 30)
                                        Approver = Approver.Substring(0, 30);
                                    string ApprovalDate = WSHEETPO.Rows[i][11].DisplayText.ToString().Trim();
                                    ApprovalDate = ApprovalDate != "" ? "'" + Convert.ToDateTime(ApprovalDate).ToString("yyyy-MM-dd") + "'" : "NULL";


                                    sbSQL.Clear();
                                    sbSQL.Append("IF NOT EXISTS(SELECT OID FROM COPO WHERE OrderNo = N'" + OrderNo + "' AND RevisionNo = '" + RevisionNo + "' AND Season = N'" + Season + "' AND OIDCUST = N'" + OIDCUST + "' AND OIDCOLOR = '" + OIDCOLOR + "' AND OIDSIZE = '" + OIDSIZE + "') ");
                                    sbSQL.Append(" BEGIN ");
                                    sbSQL.Append("   INSERT INTO COPO(OrderNo, Lot, RevisionNo, Season, BusinessUnit, OIDCUST, RevisedDate, DocumentStatus, OIDBillto, PaymentTerms, OIDCURR, OrderPlanNumber, TrimOrderNo, ItemCode, OIDCOLOR, OIDSIZE, ");
                                    sbSQL.Append("                      PatternDimensionCode, SKUCode, SampleCode, OrderQtyPCS, Remark, AllocationOrderNumber, OriginalSalesPrice, CostSheetNo, Approver, ApprovalDate) ");
                                    sbSQL.Append("   VALUES(N'" + OrderNo + "', N'" + Lot + "', '" + RevisionNo + "', N'" + Season + "', N'" + BusinessUnit + "', '" + OIDCUST + "', " + RevisedDate + ", '" + DocumentStatus + "', '" + OIDBillto + "', N'" + strPaymentTerms + "', '" + OIDCURR + "', N'" + OrderPlanNumber + "', N'" + TrimOrderNo + "', '" + OIDITEM + "', '" + OIDCOLOR + "', '" + OIDSIZE + "', ");
                                    sbSQL.Append("          N'" + PatternDimensionCode + "', N'" + SKUCode + "', N'" + SampleCode + "', '" + OrderQtyPCS + "', N'" + Remark + "', N'" + AllocationOrderNumber + "', N'" + OriginalSalesPrice + "', N'" + CostSheetNo + "', N'" + Approver + "', " + ApprovalDate + ")  ");
                                    sbSQL.Append(" END ");
                                    sbSQL.Append("ELSE ");
                                    sbSQL.Append(" BEGIN ");
                                    sbSQL.Append("   UPDATE COPO SET ");
                                    sbSQL.Append("      Lot=N'" + Lot + "', ");
                                    sbSQL.Append("      BusinessUnit = N'" + BusinessUnit + "', ");
                                    sbSQL.Append("      RevisedDate = " + RevisedDate + ", ");
                                    sbSQL.Append("      DocumentStatus = '" + DocumentStatus + "', ");
                                    sbSQL.Append("      OIDBillto = '" + OIDBillto + "', ");
                                    sbSQL.Append("      PaymentTerms = N'" + strPaymentTerms + "', ");
                                    sbSQL.Append("      OIDCURR = '" + OIDCURR + "', ");
                                    sbSQL.Append("      OrderPlanNumber = N'" + OrderPlanNumber + "', ");
                                    sbSQL.Append("      TrimOrderNo = N'" + TrimOrderNo + "', ");
                                    sbSQL.Append("      ItemCode = '" + OIDITEM + "', ");
                                    sbSQL.Append("      PatternDimensionCode = N'" + PatternDimensionCode + "', ");
                                    sbSQL.Append("      SKUCode = N'" + SKUCode + "', ");
                                    sbSQL.Append("      SampleCode = N'" + SampleCode + "', ");
                                    sbSQL.Append("      OrderQtyPCSOld = OrderQtyPCS, ");
                                    sbSQL.Append("      OrderQtyPCS = '" + OrderQtyPCS + "', ");
                                    sbSQL.Append("      Remark = N'" + Remark + "', ");
                                    sbSQL.Append("      AllocationOrderNumber = N'" + AllocationOrderNumber + "', ");
                                    sbSQL.Append("      OriginalSalesPrice = N'" + OriginalSalesPrice + "', ");
                                    sbSQL.Append("      CostSheetNo = N'" + CostSheetNo + "', ");
                                    sbSQL.Append("      Approver = N'" + Approver + "', ");
                                    sbSQL.Append("      ApprovalDate = " + ApprovalDate + " ");
                                    sbSQL.Append("    WHERE (OrderNo = N'" + OrderNo + "') AND (RevisionNo = '" + RevisionNo + "') AND (Season = N'" + Season + "') AND (OIDCUST = N'" + OIDCUST + "') AND (OIDCOLOR = '" + OIDCOLOR + "') AND (OIDSIZE = '" + OIDSIZE + "') ");
                                    sbSQL.Append(" END   ");

                                    ////memoEdit1.EditValue = sbSQL.ToString();
                                    ////break;
                                    //MessageBox.Show(sbSQL.ToString());
                                    try
                                    {
                                        chkSAVE = new DBQuery(sbSQL).runSQL();
                                        if (chkSAVE == false)
                                        {
                                            break;
                                        }
                                        else
                                        {
                                            pbcPOSave.PerformStep();
                                            pbcPOSave.Update();
                                        }
                                    }
                                    catch (Exception)
                                    { }

                                }

                            }

                            if (chkSAVE == true) //Save PO Complete
                            {
                                tabbedControlGroup5.SelectedTabPage = layoutControlGroup14; //SELECT DO TAB
                                if (cbePoSheet.Text == "2 First Sheet (PO & DO)") //มี DO ด้วย
                                    chkSAVE = SAVEDO(spsDO);
                            }

                            if (chkSAVE == true)
                            {
                                lciPregressPOSave.Visibility = LayoutVisibility.Never;
                                tabbedControlGroup5.SelectedTabPage = layoutControlGroup13;
                                FUNC.msgInfo("Save PO complete.");
                                bbiNew.PerformClick();
                            }

                        }
                    }
                }
                else if (tabbedControlGroup5.SelectedTabPage == layoutControlGroup14) //DO
                {
                    if (txeDoFilePath.Text.Trim() == "")
                    {
                        FUNC.msgWarning("Please select excel file.");
                        txeDoFilePath.Focus();
                    }
                    else if (cbeDoSheet.Text.Trim() == "")
                    {
                        FUNC.msgWarning("Please select excel sheet.");
                        cbeDoSheet.Focus();
                    }
                    else
                    {
                        if (FUNC.msgQuiz("Confirm save excel file import data ?") == true)
                        {
                            StringBuilder sbSQL = new StringBuilder();

                            bool chkSAVE = SAVEDO(spsDO);

                            if (chkSAVE == true)
                            {
                                lciPregressDOSave.Visibility = LayoutVisibility.Never;
                                FUNC.msgInfo("Save DO complete.");
                                bbiNew.PerformClick();
                            }

                        }
                    }
                }
            }


            LoadSummary();
        }

        private bool SAVEDO(DevExpress.XtraSpreadsheet.SpreadsheetControl spsDES)
        {
            bool chkSAVE = false;
            string ItemCode = "";
            string OIDITEM = "";

            string xCOLOR = "";
            string OIDCOLOR = "";
            string xSIZE = "";
            string OIDSIZE = "";
            StringBuilder sbSQL = new StringBuilder();

            IWorkbook workbookDO = spsDES.Document;
            Worksheet WSHEETDO = workbookDO.Worksheets[0];

            lciPregressDOSave.Visibility = LayoutVisibility.Always;
            pbcDOSave.Properties.Step = 1;
            pbcDOSave.Properties.PercentView = true;
            pbcDOSave.Properties.Maximum = WSHEETDO.GetDataRange().RowCount;
            pbcDOSave.Properties.Minimum = 0;
            pbcDOSave.EditValue = 0;

            for (int i = 3; i < WSHEETDO.GetDataRange().RowCount; i++)
            {
                string OrderNo = WSHEETDO.Rows[i][0].DisplayText.ToString().Trim();

                if (OrderNo != "")
                {
                    string RevisionNo = WSHEETDO.Rows[i][1].DisplayText.ToString().Trim();
                    RevisionNo = IsNumeric(RevisionNo) == true ? RevisionNo : "0";
                    string DONo = WSHEETDO.Rows[i][2].DisplayText.ToString().ToUpper().Trim();
                    string docStatus = WSHEETDO.Rows[i][3].DisplayText.ToString().Trim();
                    int DocStatus = FindStatus(docStatus);
                    string strItemCode = WSHEETDO.Rows[i][4].DisplayText.ToString().Trim().Replace("'", "''");
                    if (ItemCode != strItemCode)
                    {
                        ItemCode = strItemCode;
                        strItemCode = strItemCode.Length > 20 ? strItemCode.Substring(0, 20) : strItemCode;
                        StringBuilder sbITEM = new StringBuilder();
                        sbITEM.Append("SELECT TOP(1) OIDCSITEM FROM ItemCustomer WHERE (ItemCode = N'" + strItemCode + "') ");
                        OIDITEM = new DBQuery(sbITEM).getString();
                    }
                    string ETAWH = WSHEETDO.Rows[i][6].DisplayText.ToString().Trim();
                    ETAWH = ETAWH != "" ? "'" + Convert.ToDateTime(ETAWH).ToString("yyyy-MM-dd") + "'" : "NULL";
                    string transportationMethod = WSHEETDO.Rows[i][7].DisplayText.ToString().Trim();
                    int TransportationMethod = FindMethod(transportationMethod);
                    string PortCode = WSHEETDO.Rows[i][10].DisplayText.ToString().Trim().Replace(" ", "").Replace(",", "");
                    PortCode = PortCode.Length > 4 ? PortCode.Substring(0, 4) : PortCode;
                    PortCode = IsNumeric(PortCode) == true ? PortCode : "0";

                    string Incoterms = WSHEETDO.Rows[i][14].DisplayText.ToString().Trim(); 
                    Incoterms = Incoterms != "" ? "1" : "0"; 

                    string ContractedETD = WSHEETDO.Rows[i][15].DisplayText.ToString().Trim();
                    ContractedETD = ContractedETD != "" ? "'" + Convert.ToDateTime(ContractedETD).ToString("yyyy-MM-dd") + "'" : "NULL";
                    string Forwarder = WSHEETDO.Rows[i][16].DisplayText.ToString().Trim();

                    string OIDVEND = WSHEETDO.Rows[i][17].DisplayText.ToString().Trim();
                    OIDVEND = WSHEETDO.Rows[i][18].DisplayText.ToString().Trim();
                    OIDVEND = OIDVEND.Replace(" ", "").Replace(".", "").Replace(",", "");
                    OIDVEND = new DBQuery("SELECT OIDVEND FROM Vendor WHERE(REPLACE(REPLACE(REPLACE(Name, ' ', ''), '.', ''), ',', '') = '" + OIDVEND + "')").getString();

                    string SetCode = WSHEETDO.Rows[i][23].DisplayText.ToString().Trim();
                    string QuantityBox = WSHEETDO.Rows[i][24].DisplayText.ToString().Trim();

                    string strCOLOR = WSHEETDO.Rows[i][25].DisplayText.ToString().Trim();
                    if (xCOLOR != strCOLOR)
                    {
                        xCOLOR = strCOLOR;
                        string strCOLORNAME = WSHEETDO.Rows[i][26].DisplayText.ToString().Trim();
                        StringBuilder sbCOLOR = new StringBuilder();
                        sbCOLOR.Append("IF NOT EXISTS(SELECT OIDCOLOR FROM ProductColor WHERE (ColorNo = N'" + strCOLOR + "')) ");
                        sbCOLOR.Append(" BEGIN ");
                        sbCOLOR.Append("   INSERT INTO ProductColor(ColorNo, ColorName) VALUES(N'" + strCOLOR + "', N'" + strCOLORNAME + "') ");
                        sbCOLOR.Append(" END ");
                        sbCOLOR.Append("SELECT TOP(1) OIDCOLOR FROM ProductColor WHERE (ColorNo = N'" + strCOLOR + "') ");
                        OIDCOLOR = new DBQuery(sbCOLOR).getString();
                    }

                    string strSIZE = WSHEETDO.Rows[i][28].DisplayText.ToString().Trim();
                    if (xSIZE != strSIZE)
                    {
                        xSIZE = strSIZE;
                        string strSIZENAME = WSHEETDO.Rows[i][29].DisplayText.ToString().Trim();
                        StringBuilder sbSIZE = new StringBuilder();
                        sbSIZE.Append("IF NOT EXISTS(SELECT OIDSIZE FROM ProductSize WHERE (SizeNo = N'" + strSIZE + "')) ");
                        sbSIZE.Append(" BEGIN ");
                        sbSIZE.Append("   INSERT INTO ProductSize(SizeNo, SizeName) VALUES(N'" + strSIZE + "', N'" + strSIZENAME + "') ");
                        sbSIZE.Append(" END ");
                        sbSIZE.Append("SELECT TOP(1) OIDSIZE FROM ProductSize WHERE (SizeNo = N'" + strSIZE + "') ");
                        OIDSIZE = new DBQuery(sbSIZE).getString();
                    }

                    string PatternDimensionCode = WSHEETDO.Rows[i][27].DisplayText.ToString().Trim();
                    string QtyperSet = WSHEETDO.Rows[i][30].DisplayText.ToString().Trim();
                    QtyperSet = IsNumeric(QtyperSet) == true ? QtyperSet : "0";

                    string PickingUnit = WSHEETDO.Rows[i][31].DisplayText.ToString().Trim();
                    PickingUnit = IsNumeric(PickingUnit) == true ? PickingUnit : "0";
                    string CartonType = "0";

                    sbSQL.Clear();
                    sbSQL.Append("IF NOT EXISTS(SELECT OIDDO FROM CODO WHERE DONo = N'" + DONo + "' AND OrderNo = N'" + OrderNo + "' AND RevisionNo = '" + RevisionNo + "' AND OIDCOLOR = '" + OIDCOLOR + "' AND OIDSIZE = '" + OIDSIZE + "') ");
                    sbSQL.Append(" BEGIN ");
                    sbSQL.Append("   INSERT INTO CODO(OrderNo, RevisionNo, DONo, DocStatus, ItemCode, ETAWH, TransportationMethod, PortCode, Incoterms, ContractedETD, Forwarder, OIDVEND, SetCode, QuantityBox, ");
                    sbSQL.Append("                      OIDCOLOR, OIDSIZE, PatternDimensionCode, QtyperSet, PickingUnit, CartonType, UpdatedBy, UpdatedDate) ");
                    sbSQL.Append("   VALUES(N'" + OrderNo + "', '" + RevisionNo + "', N'" + DONo + "', '" + DocStatus + "', '" + OIDITEM + "', " + ETAWH + ", '" + TransportationMethod + "', '" + PortCode + "', '" + Incoterms + "', " + ContractedETD + ", '" + Forwarder + "', '" + OIDVEND + "', N'" + SetCode + "', '" + QuantityBox + "', ");
                    sbSQL.Append("          '" + OIDCOLOR + "', '" + OIDSIZE + "', N'" + PatternDimensionCode + "', '" + QtyperSet + "', '" + PickingUnit + "', '" + CartonType + "', '0', GETDATE())  ");
                    sbSQL.Append(" END ");
                    sbSQL.Append("ELSE ");
                    sbSQL.Append(" BEGIN ");
                    sbSQL.Append("   UPDATE CODO SET ");
                    sbSQL.Append("      DocStatus='" + DocStatus + "', ");
                    sbSQL.Append("      ItemCode = '" + OIDITEM + "', ");
                    sbSQL.Append("      ETAWHOld = ETAWH, ");
                    sbSQL.Append("      ETAWH = " + ETAWH + ", ");
                    sbSQL.Append("      TransportationMethod = '" + TransportationMethod + "', ");
                    sbSQL.Append("      PortCode = '" + PortCode + "', ");
                    sbSQL.Append("      Incoterms = '" + Incoterms + "', ");
                    sbSQL.Append("      ContractedETDOld = ContractedETD, ");
                    sbSQL.Append("      ContractedETD = " + ContractedETD + ", ");
                    sbSQL.Append("      Forwarder = '" + Forwarder + "', ");
                    sbSQL.Append("      OIDVEND = '" + OIDVEND + "', ");
                    sbSQL.Append("      SetCode = N'" + SetCode + "', ");
                    sbSQL.Append("      QuantityBoxOld = QuantityBox, ");
                    sbSQL.Append("      QuantityBox = '" + QuantityBox + "', ");
                    sbSQL.Append("      PatternDimensionCode = N'" + PatternDimensionCode + "', ");
                    sbSQL.Append("      QtyperSet = '" + QtyperSet + "', ");
                    sbSQL.Append("      PickingUnit = '" + PickingUnit + "', ");
                    sbSQL.Append("      CartonType = '" + CartonType + "', ");
                    sbSQL.Append("      UpdatedBy = '0', ");
                    sbSQL.Append("      UpdatedDate = GETDATE() ");
                    sbSQL.Append("    WHERE (DONo = N'" + DONo + "') AND (OrderNo = N'" + OrderNo + "') AND (RevisionNo = '" + RevisionNo + "') AND (OIDCOLOR = '" + OIDCOLOR + "') AND (OIDSIZE = '" + OIDSIZE + "') ");
                    sbSQL.Append(" END   ");

                    //MessageBox.Show(sbSQL.ToString());
                    try
                    {
                        chkSAVE = new DBQuery(sbSQL).runSQL();
                        if (chkSAVE == false)
                        {
                            break;
                        }
                        else
                        {
                            pbcDOSave.PerformStep();
                            pbcDOSave.Update();
                        }
                    }
                    catch (Exception)
                    { }


                }
            }
            lciPregressDOSave.Visibility = LayoutVisibility.Never;
            return chkSAVE;
        }

        public static bool IsNumeric(string Expression)
        {
            double retNum;

            bool isNum = Double.TryParse(Expression.ToString(), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
            return isNum;
        }

        private void bbiExcel_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (tabbedControlGroup2.SelectedTabPage == layoutControlGroup2) //PO
            {
                string pathFile = new ObjSet.Folder(@"C:\MDS\Export\").GetPath() + "SummaryPOList_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
                gvSumPO.ExportToXlsx(pathFile);
                System.Diagnostics.Process.Start(pathFile);
            }
            else if (tabbedControlGroup2.SelectedTabPage == layoutControlGroup3) //DO
            {
                string pathFile = new ObjSet.Folder(@"C:\MDS\Export\").GetPath() + "SummaryDOList_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
                gvSumDO.ExportToXlsx(pathFile);
                System.Diagnostics.Process.Start(pathFile);
            }
        }

        private void bbiPrintPreview_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (tabbedControlGroup2.SelectedTabPage == layoutControlGroup2) //PO
                gcSumPO.ShowPrintPreview();
            else if (tabbedControlGroup2.SelectedTabPage == layoutControlGroup3) //DO
                gcSumDO.ShowPrintPreview();
        }

        private void bbiPrint_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (tabbedControlGroup2.SelectedTabPage == layoutControlGroup2) //PO
                gcSumPO.Print();
            else if (tabbedControlGroup2.SelectedTabPage == layoutControlGroup3) //DO
                gcSumDO.Print();
        }

        // This event is generated by Data Source Configuration Wizard
        void unboundSource1_ValueNeeded(object sender, DevExpress.Data.UnboundSourceValueNeededEventArgs e)
        {
            // Handle this event to obtain data from your data source
            // e.Value = something /* TODO: Assign the real data here.*/
        }

        // This event is generated by Data Source Configuration Wizard
        void unboundSource1_ValuePushed(object sender, DevExpress.Data.UnboundSourceValuePushedEventArgs e)
        {
            // Handle this event to save modified data back to your data source
            // something = e.Value; /* TODO: Propagate the value into the storage.*/
        }

        private void tabbedControlGroup3_SelectedPageChanged(object sender, DevExpress.XtraLayout.LayoutTabPageChangedEventArgs e)
        {
            if (tabbedControlGroup3.SelectedTabPage == layoutControlGroup4) //Summary
            {
                bbiNew.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
                bbiEdit.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
                bbiSave.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
                bbiDelete.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;

                bbiRefresh.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
                ribbonPageGroup2.Visible = true;
                ribbonPageGroup5.Visible = true;

                tabbedControlGroup2.SelectedTabPage = layoutControlGroup2;
            }
            else if (tabbedControlGroup3.SelectedTabPage == layoutControlGroup5) //Entry
            {
                bbiNew.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
                bbiEdit.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
                bbiSave.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
                bbiDelete.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;

                bbiRefresh.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
                ribbonPageGroup2.Visible = false;
                ribbonPageGroup5.Visible = false;

                tabbedControlGroup4.SelectedTabPage = layoutControlGroup8;
            }
            else if (tabbedControlGroup3.SelectedTabPage == layoutControlGroup12) //Import
            {
                bbiNew.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
                bbiEdit.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
                bbiSave.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
                bbiDelete.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;

                bbiRefresh.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
                ribbonPageGroup2.Visible = false;
                ribbonPageGroup5.Visible = false;

                tabbedControlGroup5.SelectedTabPage = layoutControlGroup13;
            }
        }

        private void gvSumPO_RowCellClick(object sender, RowCellClickEventArgs e)
        {
            if (gvSumPO.IsFilterRow(e.RowHandle)) return;
        }

        private void gvSumDO_RowCellClick(object sender, RowCellClickEventArgs e)
        {
            if (gvSumDO.IsFilterRow(e.RowHandle)) return;
        }

        private void gvEntryPO_RowCellClick(object sender, RowCellClickEventArgs e)
        {
            if (gvEntryPO.IsFilterRow(e.RowHandle)) return;
        }

        private void gvEntryDO_RowCellClick(object sender, RowCellClickEventArgs e)
        {
            if (gvEntryDO.IsFilterRow(e.RowHandle)) return;
        }

        private void sbPoBrowse_Click(object sender, EventArgs e)
        {
            cbePoSheet.Properties.Items.Clear();
            cbePoSheet.Text = "";

            ofdPO.Filter = "Excel files |*.xlsx;*.xls;*.csv";
            ofdPO.FileName = "";
            ofdPO.Title = "Select Excel File";

            if (ofdPO.ShowDialog() == DialogResult.OK)
            {
                txePoFilePath.Text = ofdPO.FileName;
                DevExpress.XtraSpreadsheet.SpreadsheetControl xss = new DevExpress.XtraSpreadsheet.SpreadsheetControl();
                IWorkbook workbook = xss.Document;
                using (FileStream stream = new FileStream(txePoFilePath.Text, FileMode.Open))
                {
                    string ext = Path.GetExtension(txePoFilePath.Text);
                    if (ext == ".xlsx")
                        workbook.LoadDocument(stream, DocumentFormat.Xlsx);
                    else if (ext == ".xls")
                        workbook.LoadDocument(stream, DocumentFormat.Xls);
                    else if (ext == ".csv")
                        workbook.LoadDocument(stream, DocumentFormat.Csv);
                }
                WorksheetCollection worksheets = workbook.Worksheets;

                bool chkPO = false;
                bool chkDO = false;

                for (int i = 0; i < worksheets.Count; i++)
                    if (worksheets[i].Name.IndexOf("PO") > -1)
                    {
                        cbePoSheet.Properties.Items.Add(worksheets[i].Name);
                        chkPO = true;
                    }
                    else if (worksheets[i].Name.IndexOf("DO") > -1)
                        chkDO = true;

                if (chkPO == true && chkDO == true)
                    cbePoSheet.Properties.Items.Insert(0, "2 First Sheet (PO & DO)");

            }
        }

        private void sbDoBrowse_Click(object sender, EventArgs e)
        {
            cbeDoSheet.Properties.Items.Clear();
            cbeDoSheet.Text = "";

            ofdDO.Filter = "Excel files |*.xlsx;*.xls;*.csv";
            ofdDO.FileName = "";
            ofdDO.Title = "Select Excel File";

            if (ofdDO.ShowDialog() == DialogResult.OK)
            {
                txeDoFilePath.Text = ofdDO.FileName;
                DevExpress.XtraSpreadsheet.SpreadsheetControl xss = new DevExpress.XtraSpreadsheet.SpreadsheetControl();
                IWorkbook workbook = xss.Document;
                using (FileStream stream = new FileStream(txeDoFilePath.Text, FileMode.Open))
                {
                    string ext = Path.GetExtension(txeDoFilePath.Text);
                    if (ext == ".xlsx")
                        workbook.LoadDocument(stream, DocumentFormat.Xlsx);
                    else if (ext == ".xls")
                        workbook.LoadDocument(stream, DocumentFormat.Xls);
                    else if (ext == ".csv")
                        workbook.LoadDocument(stream, DocumentFormat.Csv);
                }
                WorksheetCollection worksheets = workbook.Worksheets;
                for (int i = 0; i < worksheets.Count; i++)
                    if (worksheets[i].Name.IndexOf("DO") > -1)
                        cbeDoSheet.Properties.Items.Add(worksheets[i].Name);
            }
        }

        private bool chkDuplicatePO(string PO)
        {
            PO = PO.ToUpper().Trim().Replace("'", "''");
            bool chkDup = true;
            if (PO != "")
            {
                if (gluePoDocumentStatus.EditValue.ToString() == "0") //New
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT TOP(1) OrderNo FROM COPO WHERE (OrderNo = N'" + PO + "') ");
                    if (new DBQuery(sbSQL).getString() != "")
                    {
                        chkDup = false;
                    }
                }
                else if (gluePoDocumentStatus.EditValue.ToString() == "1") //Revise
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT TOP(1) OrderNo FROM COPO WHERE (OrderNo = N'" + PO + "') AND (RevisionNo = '" + spePoRevisionNo.Value.ToString() + "') ");
                    if (new DBQuery(sbSQL).getString() != "")
                    {
                        chkDup = false;
                    }
                }
                else 
                {
                    chkDup = true;
                }
               
            }
            return chkDup;
        }

        private void txePoOrderNo_Leave(object sender, EventArgs e)
        {
            txePoOrderNo.Text = txePoOrderNo.Text.ToUpper().Trim();
            if (txePoOrderNo.Text.Trim() != "")
            {
                bool chkDup = chkDuplicatePO(txePoOrderNo.Text);
                if (chkDup == false) //Load PO
                {
                    txePoOrderNo.Text = "";
                    txePoOrderNo.Focus();
                    FUNC.msgWarning("Duplicate PO Order No. !! Please Change.");
                }
            }
        }

        private void txePoOrderNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txePoOrderNo.Text = txePoOrderNo.Text.ToUpper().Trim();
                spePoRevisionNo.Focus();
            }
        }

        private void sbPoClearTable_Click(object sender, EventArgs e)
        {
            gvEntryPO.CloseEditor();
            gvEntryPO.UpdateCurrentRow();
            for (int i = gvEntryPO.RowCount - 1; i >= 0; i--)
                gvEntryPO.DeleteRow(i);
            gvEntryPO.OptionsView.ColumnAutoWidth = false;
            gvEntryPO.BestFitColumns();
            //gvEntryPO.Columns["ProductionPlanID"].Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
        }

        private void sbDoClearTable_Click(object sender, EventArgs e)
        {
            gvEntryDO.CloseEditor();
            gvEntryDO.UpdateCurrentRow();
            for (int i = gvEntryDO.RowCount - 1; i >= 0; i--)
                gvEntryDO.DeleteRow(i);
            gvEntryDO.OptionsView.ColumnAutoWidth = false;
            gvEntryDO.BestFitColumns();
            //gvEntryDO.Columns["ProductionPlanID"].Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
        }

        private void gvEntryPO_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {
            if (e.Column.FieldName == "OrderQtyPCS"
                || e.Column.FieldName == "OIDCOLOR"
                || e.Column.FieldName == "OIDSIZE")
            {
                DevExpress.XtraEditors.ViewInfo.BaseEditViewInfo info = ((DevExpress.XtraGrid.Views.Grid.ViewInfo.GridCellInfo)e.Cell).ViewInfo;
                string error = GetError(e.CellValue, e.RowHandle, e.Column);
                if (e.CellValue == null || String.IsNullOrEmpty(e.CellValue.ToString()) == true)
                {
                    SetError(info, error);
                }
            }
        }

        private void gvEntryPO_InitNewRow(object sender, InitNewRowEventArgs e)
        {
            gvEntryPO.SetRowCellValue(e.RowHandle, "OrderQtyPCS", 0);
        }

        private void gvEntryDO_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {
           
        }

        private void gvEntryDO_InitNewRow(object sender, InitNewRowEventArgs e)
        {
            gvEntryDO.SetRowCellValue(e.RowHandle, "QtyBox", 0);
            gvEntryDO.SetRowCellValue(e.RowHandle, "QtySet", 0);
            gvEntryDO.SetRowCellValue(e.RowHandle, "PickUnit", 0);
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

        private int CountErrorPO()
        {
            int CError = 0;
            DataTable dtError = (DataTable)gcEntryPO.DataSource;
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
                            if (c.ColumnName == "OIDCOLOR" || c.ColumnName == "OIDSIZE" || c.ColumnName == "OrderQtyPCS")
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

        private bool chkDupPO(string PONO, int rowIndex)
        {
            gvEntryPO.CloseEditor();
            gvEntryPO.UpdateCurrentRow();

            PONO = PONO.ToUpper().Trim();
            bool chkDup = true;

            if (PONO != "")
            {
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT COUNT(OID) AS COUNT_ID ");
                sbSQL.Append("FROM COPO ");
                sbSQL.Append("WHERE (OrderNo = N'" + PONO + "') ");
                if (new DBQuery(sbSQL).getInt() > 0)
                    chkDup = false;
                else
                {
                    int countPlan = 0;
                    DataTable dtFind = (DataTable)gcEntryPO.DataSource;
                    int xRow = 0;
                    foreach (DataRow row in dtFind.Rows)
                    {
                        string chkPONO = row["OrderNo"].ToString().ToUpper().Trim();
                        if (chkPONO == PONO && xRow != rowIndex)
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

        private bool chkDupDO(string PONO, string DONO, int rowIndex)
        {
            gvEntryDO.CloseEditor();
            gvEntryDO.UpdateCurrentRow();

            PONO = PONO.ToUpper().Trim();
            DONO = DONO.ToUpper().Trim();

            bool chkDup = true;

            if (PONO != "")
            {
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT COUNT(OIDDO) AS COUNT_ID ");
                sbSQL.Append("FROM CODO ");
                sbSQL.Append("WHERE (OrderNo = N'" + PONO + "') AND (DONo = N'" + DONO + "') ");
                if (new DBQuery(sbSQL).getInt() > 0)
                    chkDup = false;
                else
                {
                    int countPlan = 0;
                    DataTable dtFind = (DataTable)gcEntryDO.DataSource;
                    int xRow = 0;
                    foreach (DataRow row in dtFind.Rows)
                    {
                        string chkPONO = row["OrderNo"].ToString().ToUpper().Trim();
                        string chkDONO = row["DONo"].ToString().ToUpper().Trim();
                        if (chkPONO == PONO && chkDONO == DONO && xRow != rowIndex)
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

        private void setDefaultStatusPO()
        {
            layoutControlItem5.Visibility = LayoutVisibility.Always; //textEdit
            layoutControlItem14.Visibility = LayoutVisibility.Never; //searchLookUpEdit

            txePoOrderNo.Text = "";
            sluePoOrderNo.EditValue = "";
            spePoRevisionNo.EditValue = "";
            txePoLot.Text = "";

            dtePoRevisedDate.EditValue = null;
            dtePoRevisedDate.ReadOnly = false;

            txePoOrderPlanNumber.Text = "";
            sluePoOIDCUST.EditValue = "";
            spePoSeason.Value = 0;
            gluePoSeason.EditValue = "";
            gluePoBusinessUnit.EditValue = "";
            sluePoItemCode.EditValue = "";
            txePoRemark.Text = "";
            txePoAllocationOrderNumber.Text = "";
            txePoOriginalSalesPrice.Text = "";
            txePoApprover.Text = "";
            dtePoApprovalDate.EditValue = null;
            sluePoOIDBillto.EditValue = "";
            sluePoPaymentTerms.EditValue = "";
            gluePoOIDCURR.EditValue = "";

            dtPOEntry.Rows.Clear();
            gcEntryPO.DataSource = dtPOEntry;

            txePoItemName.Text = "";
            txePoFabricWidth.Text = "";
            txePoFBComposition.Text = "";
            txePoAddress.Text = "";
            txePoTelephoneNo.Text = "";

            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT DISTINCT PO.OrderNo + '_' + CONVERT(NVARCHAR, PO.RevisionNo) AS POID, PO.OrderNo AS [PO. Order No.], PO.RevisionNo AS [Revision No.], ST.Status, PO.Season, CUS.Name AS Customer, IC.ItemName ");
            sbSQL.Append("FROM   COPO AS PO INNER JOIN ");
            sbSQL.Append("       Customer AS CUS ON PO.OIDCUST = CUS.OIDCUST INNER JOIN ");
            sbSQL.Append("       ItemCustomer AS IC ON PO.ItemCode = IC.OIDCSITEM INNER JOIN ");
            sbSQL.Append("       (" + sbSTATUS.ToString() + ") AS ST ON PO.DocumentStatus = ST.ID ");
            sbSQL.Append("ORDER BY [PO. Order No.] ");

            if (gluePoDocumentStatus.Text.ToString() == "") //null
            {
                layoutControlItem5.Visibility = LayoutVisibility.Always; //textEdit
                layoutControlItem14.Visibility = LayoutVisibility.Never; //searchLookUpEdit

                spePoRevisionNo.ReadOnly = true;
                txePoLot.ReadOnly = true;

                dtePoRevisedDate.EditValue = null;
                dtePoRevisedDate.ReadOnly = false;

                txePoOrderPlanNumber.ReadOnly = true;
                sluePoOIDCUST.ReadOnly = true;
                spePoSeason.ReadOnly = true;
                gluePoSeason.ReadOnly = true;
                gluePoBusinessUnit.ReadOnly = true;
                sluePoItemCode.ReadOnly = true;
                txePoRemark.ReadOnly = true;
                txePoAllocationOrderNumber.ReadOnly = true;
                txePoOriginalSalesPrice.ReadOnly = true;
                txePoApprover.ReadOnly = true;
                dtePoApprovalDate.ReadOnly = true;
                sluePoOIDBillto.ReadOnly = true;
                sluePoPaymentTerms.ReadOnly = true;
                gluePoOIDCURR.ReadOnly = true;

                gvEntryPO.OptionsBehavior.ReadOnly = true;
                gvEntryPO.OptionsBehavior.Editable = false;
            }
            else
            {
                if (gluePoDocumentStatus.EditValue.ToString() == "0") //New
                {
                    layoutControlItem5.Visibility = LayoutVisibility.Always; //textEdit
                    layoutControlItem14.Visibility = LayoutVisibility.Never; //searchLookUpEdit

                    spePoRevisionNo.Value = 0;
                    txePoOrderNo.ReadOnly = false;
                    spePoRevisionNo.ReadOnly = true;
                    txePoLot.ReadOnly = false;

                    dtePoRevisedDate.ReadOnly = true;
                    dtePoRevisedDate.EditValue = null;

                    txePoOrderPlanNumber.ReadOnly = false;
                    sluePoOIDCUST.ReadOnly = false;
                    spePoSeason.ReadOnly = false;
                    gluePoSeason.ReadOnly = false;
                    gluePoBusinessUnit.ReadOnly = false;
                    sluePoItemCode.ReadOnly = false;
                    txePoRemark.ReadOnly = false;
                    txePoAllocationOrderNumber.ReadOnly = false;
                    txePoOriginalSalesPrice.ReadOnly = false;
                    txePoApprover.ReadOnly = false;
                    dtePoApprovalDate.ReadOnly = false;
                    sluePoOIDBillto.ReadOnly = false;
                    sluePoPaymentTerms.ReadOnly = false;
                    gluePoOIDCURR.ReadOnly = false;

                    gvEntryPO.OptionsBehavior.ReadOnly = false;
                    gvEntryPO.OptionsBehavior.Editable = true;
                    gvEntryPO.OptionsBehavior.AllowAddRows = DevExpress.Utils.DefaultBoolean.True;
                    gvEntryPO.OptionsView.NewItemRowPosition = NewItemRowPosition.Bottom;
                }
                else
                {
                    new ObjDevEx.setSearchLookUpEdit(sluePoOrderNo, sbSQL, "PO. Order No.", "POID").getData();

                    sluePoOrderNo.Properties.PopulateViewColumns();
                    sluePoOrderNo.Properties.View.Columns["POID"].Visible = false;

                    if (gluePoDocumentStatus.EditValue.ToString() == "1") //Revise
                    {
                        layoutControlItem5.Visibility = LayoutVisibility.Never; //textEdit
                        layoutControlItem14.Visibility = LayoutVisibility.Always; //searchLookUpEdit

                        txePoOrderNo.ReadOnly = false;
                        spePoRevisionNo.ReadOnly = true;
                        txePoLot.ReadOnly = false;

                        dtePoRevisedDate.ReadOnly = false;
                        dtePoRevisedDate.EditValue = DateTime.Now;

                        txePoOrderPlanNumber.ReadOnly = false;
                        sluePoOIDCUST.ReadOnly = false;
                        spePoSeason.ReadOnly = false;
                        gluePoSeason.ReadOnly = false;
                        gluePoBusinessUnit.ReadOnly = false;
                        sluePoItemCode.ReadOnly = false;
                        txePoRemark.ReadOnly = false;
                        txePoAllocationOrderNumber.ReadOnly = false;
                        txePoOriginalSalesPrice.ReadOnly = false;
                        txePoApprover.ReadOnly = false;
                        dtePoApprovalDate.ReadOnly = false;
                        sluePoOIDBillto.ReadOnly = false;
                        sluePoPaymentTerms.ReadOnly = false;
                        gluePoOIDCURR.ReadOnly = false;

                        gvEntryPO.OptionsBehavior.ReadOnly = false;
                        gvEntryPO.OptionsBehavior.Editable = true;
                        gvEntryPO.OptionsBehavior.AllowAddRows = DevExpress.Utils.DefaultBoolean.False;
                        gvEntryPO.OptionsView.NewItemRowPosition = NewItemRowPosition.None;
                    }
                    else if (gluePoDocumentStatus.EditValue.ToString() == "2") //Change
                    {
                        layoutControlItem5.Visibility = LayoutVisibility.Never; //textEdit
                        layoutControlItem14.Visibility = LayoutVisibility.Always; //searchLookUpEdit

                        spePoRevisionNo.ReadOnly = true;
                        txePoLot.ReadOnly = false;
                        
                        dtePoRevisedDate.EditValue = null;
                        dtePoRevisedDate.ReadOnly = false;

                        txePoOrderPlanNumber.ReadOnly = false;
                        sluePoOIDCUST.ReadOnly = false;
                        spePoSeason.ReadOnly = false;
                        gluePoSeason.ReadOnly = false;
                        gluePoBusinessUnit.ReadOnly = false;
                        sluePoItemCode.ReadOnly = false;
                        txePoRemark.ReadOnly = false;
                        txePoAllocationOrderNumber.ReadOnly = false;
                        txePoOriginalSalesPrice.ReadOnly = false;
                        txePoApprover.ReadOnly = false;
                        dtePoApprovalDate.ReadOnly = false;
                        sluePoOIDBillto.ReadOnly = false;
                        sluePoPaymentTerms.ReadOnly = false;
                        gluePoOIDCURR.ReadOnly = false;

                        gvEntryPO.OptionsBehavior.ReadOnly = false;
                        gvEntryPO.OptionsBehavior.Editable = true;
                        gvEntryPO.OptionsBehavior.AllowAddRows = DevExpress.Utils.DefaultBoolean.True;
                        gvEntryPO.OptionsView.NewItemRowPosition = NewItemRowPosition.Bottom;
                    }
                    else if (gluePoDocumentStatus.EditValue.ToString() == "3") //Cancel
                    {
                        layoutControlItem5.Visibility = LayoutVisibility.Never; //textEdit
                        layoutControlItem14.Visibility = LayoutVisibility.Always; //searchLookUpEdit

                        spePoRevisionNo.ReadOnly = true;
                        txePoLot.ReadOnly = true;

                        dtePoRevisedDate.EditValue = null;
                        dtePoRevisedDate.ReadOnly = false;

                        txePoOrderPlanNumber.ReadOnly = true;
                        sluePoOIDCUST.ReadOnly = true;
                        spePoSeason.ReadOnly = true;
                        gluePoSeason.ReadOnly = true;
                        gluePoBusinessUnit.ReadOnly = true;
                        sluePoItemCode.ReadOnly = true;
                        txePoRemark.ReadOnly = true;
                        txePoAllocationOrderNumber.ReadOnly = true;
                        txePoOriginalSalesPrice.ReadOnly = true;
                        txePoApprover.ReadOnly = true;
                        dtePoApprovalDate.ReadOnly = true;
                        sluePoOIDBillto.ReadOnly = true;
                        sluePoPaymentTerms.ReadOnly = true;
                        gluePoOIDCURR.ReadOnly = true;

                        gvEntryPO.OptionsBehavior.ReadOnly = true;
                        gvEntryPO.OptionsBehavior.Editable = false;
                        gvEntryPO.OptionsBehavior.AllowAddRows = DevExpress.Utils.DefaultBoolean.False;
                        gvEntryPO.OptionsView.NewItemRowPosition = NewItemRowPosition.None;
                    }
                    else if (gluePoDocumentStatus.EditValue.ToString() == "9") //Finished
                    {
                        layoutControlItem5.Visibility = LayoutVisibility.Never; //textEdit
                        layoutControlItem14.Visibility = LayoutVisibility.Always; //searchLookUpEdit

                        spePoRevisionNo.ReadOnly = true;
                        txePoLot.ReadOnly = true;

                        dtePoRevisedDate.EditValue = null;
                        dtePoRevisedDate.ReadOnly = false;

                        txePoOrderPlanNumber.ReadOnly = true;
                        sluePoOIDCUST.ReadOnly = true;
                        spePoSeason.ReadOnly = true;
                        gluePoSeason.ReadOnly = true;
                        gluePoBusinessUnit.ReadOnly = true;
                        sluePoItemCode.ReadOnly = true;
                        txePoRemark.ReadOnly = true;
                        txePoAllocationOrderNumber.ReadOnly = true;
                        txePoOriginalSalesPrice.ReadOnly = true;
                        txePoApprover.ReadOnly = true;
                        dtePoApprovalDate.ReadOnly = true;
                        sluePoOIDBillto.ReadOnly = true;
                        sluePoPaymentTerms.ReadOnly = true;
                        gluePoOIDCURR.ReadOnly = true;

                        gvEntryPO.OptionsBehavior.ReadOnly = true;
                        gvEntryPO.OptionsBehavior.Editable = false;
                        gvEntryPO.OptionsBehavior.AllowAddRows = DevExpress.Utils.DefaultBoolean.False;
                        gvEntryPO.OptionsView.NewItemRowPosition = NewItemRowPosition.None;
                    }
                }
            }
        }

        private void setDefaultStatusDO()
        {
            layoutControlItem39.Visibility = LayoutVisibility.Always; //textEdit
            layoutControlItem17.Visibility = LayoutVisibility.Never; //searchLookUpEdit

            slueDoOrderNo.EditValue = "";
            speDoRevisionNo.EditValue = "";

            txeDoDONo.Text = "";
            slueDoDONo.EditValue = "";

            slueDoItemCode.EditValue = "";
            txeDoItemName.Text = "";

            glueDoTransportationMethod.EditValue = "";
            slueDoPortCode.EditValue = "";

            txeDoPortName.Text = "";
            txeDoWHName.Text = "";
            txeDoAddress.Text = "";
            txeDoTelephone.Text = "";
            txeDoPersonInCharge.Text = "";

            glueDoIncoterms.EditValue = "1";
            txeDoForwarder.Text = "";
            slueDoOIDVEND.EditValue = "";
            dteDoETAWH.EditValue = null;
            dteDoContractedETD.EditValue = null;

            gvEntryDO.OptionsBehavior.AllowAddRows = DevExpress.Utils.DefaultBoolean.False;
            gvEntryDO.OptionsView.NewItemRowPosition = NewItemRowPosition.None;

            txeDATE.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT DISTINCT DO.DONo AS [DO No.], DO.OrderNo AS [PO. Order No.], DO.RevisionNo AS [Revision No.], ST.Status, IC.ItemName ");
            sbSQL.Append("FROM   CODO AS DO INNER JOIN ");
            sbSQL.Append("       ItemCustomer AS IC ON DO.ItemCode = IC.OIDCSITEM INNER JOIN ");
            sbSQL.Append("       (" + sbSTATUS.ToString() + ") AS ST ON DO.DocStatus = ST.ID ");
            sbSQL.Append("ORDER BY [DO No.] ");

            if (glueDoDocStatus.Text.ToString() == "") //null
            {
                layoutControlItem39.Visibility = LayoutVisibility.Always; //textEdit
                layoutControlItem17.Visibility = LayoutVisibility.Never; //searchLookUpEdit

                slueDoOrderNo.ReadOnly = true;

                txeDoDONo.ReadOnly = true;
                slueDoDONo.ReadOnly = true;
                glueDoTransportationMethod.ReadOnly = true;
                slueDoPortCode.ReadOnly = true;
                glueDoIncoterms.ReadOnly = true;
                txeDoForwarder.ReadOnly = true;
                slueDoOIDVEND.ReadOnly = true;
                dteDoETAWH.ReadOnly = true;
                dteDoContractedETD.ReadOnly = true;

                gvEntryDO.OptionsBehavior.ReadOnly = true;
                gvEntryDO.OptionsBehavior.Editable = false;
            }
            else
            {
                if (glueDoDocStatus.EditValue.ToString() == "0") //New
                {
                    layoutControlItem39.Visibility = LayoutVisibility.Always; //textEdit
                    layoutControlItem17.Visibility = LayoutVisibility.Never; //searchLookUpEdit

                    slueDoOrderNo.ReadOnly = false;

                    txeDoDONo.ReadOnly = false;
                    slueDoDONo.ReadOnly = false;
                    glueDoTransportationMethod.ReadOnly = false;
                    slueDoPortCode.ReadOnly = false;
                    glueDoIncoterms.ReadOnly = false;
                    txeDoForwarder.ReadOnly = false;
                    slueDoOIDVEND.ReadOnly = false;
                    dteDoETAWH.ReadOnly = false;
                    dteDoContractedETD.ReadOnly = false;

                    gvEntryDO.OptionsBehavior.ReadOnly = false;
                    gvEntryDO.OptionsBehavior.Editable = true;

                }
                else
                {
                    new ObjDevEx.setSearchLookUpEdit(slueDoDONo, sbSQL, "DO No.", "DO No.").getData();

                    if (glueDoDocStatus.EditValue.ToString() == "1") //Revise
                    {
                        layoutControlItem39.Visibility = LayoutVisibility.Never; //textEdit
                        layoutControlItem17.Visibility = LayoutVisibility.Always; //searchLookUpEdit

                        slueDoOrderNo.ReadOnly = true;

                        txeDoDONo.ReadOnly = false;
                        slueDoDONo.ReadOnly = false;
                        glueDoTransportationMethod.ReadOnly = false;
                        slueDoPortCode.ReadOnly = false;
                        glueDoIncoterms.ReadOnly = false;
                        txeDoForwarder.ReadOnly = false;
                        slueDoOIDVEND.ReadOnly = false;
                        dteDoETAWH.ReadOnly = false;
                        dteDoContractedETD.ReadOnly = false;

                        gvEntryDO.OptionsBehavior.ReadOnly = false;
                        gvEntryDO.OptionsBehavior.Editable = true;
                        
                    }
                    else if (glueDoDocStatus.EditValue.ToString() == "2") //Change
                    {
                        layoutControlItem39.Visibility = LayoutVisibility.Never; //textEdit
                        layoutControlItem17.Visibility = LayoutVisibility.Always; //searchLookUpEdit

                        slueDoOrderNo.ReadOnly = false;

                        txeDoDONo.ReadOnly = false;
                        slueDoDONo.ReadOnly = false;
                        glueDoTransportationMethod.ReadOnly = false;
                        slueDoPortCode.ReadOnly = false;
                        glueDoIncoterms.ReadOnly = false;
                        txeDoForwarder.ReadOnly = false;
                        slueDoOIDVEND.ReadOnly = false;
                        dteDoETAWH.ReadOnly = false;
                        dteDoContractedETD.ReadOnly = false;

                        gvEntryDO.OptionsBehavior.ReadOnly = false;
                        gvEntryDO.OptionsBehavior.Editable = true;

                    }
                    else if (glueDoDocStatus.EditValue.ToString() == "3") //Cancel
                    {
                        layoutControlItem39.Visibility = LayoutVisibility.Never; //textEdit
                        layoutControlItem17.Visibility = LayoutVisibility.Always; //searchLookUpEdit

                        slueDoOrderNo.ReadOnly = true;

                        txeDoDONo.ReadOnly = false;
                        slueDoDONo.ReadOnly = false;
                        glueDoTransportationMethod.ReadOnly = true;
                        slueDoPortCode.ReadOnly = true;
                        glueDoIncoterms.ReadOnly = true;
                        txeDoForwarder.ReadOnly = true;
                        slueDoOIDVEND.ReadOnly = true;
                        dteDoETAWH.ReadOnly = true;
                        dteDoContractedETD.ReadOnly = true;

                        gvEntryDO.OptionsBehavior.ReadOnly = true;
                        gvEntryDO.OptionsBehavior.Editable = false;

                    }
                    else if (glueDoDocStatus.EditValue.ToString() == "9") //Finished
                    {
                        layoutControlItem39.Visibility = LayoutVisibility.Never; //textEdit
                        layoutControlItem17.Visibility = LayoutVisibility.Always; //searchLookUpEdit

                        slueDoOrderNo.ReadOnly = true;

                        txeDoDONo.ReadOnly = false;
                        slueDoDONo.ReadOnly = false;
                        glueDoTransportationMethod.ReadOnly = true;
                        slueDoPortCode.ReadOnly = true;
                        glueDoIncoterms.ReadOnly = true;
                        txeDoForwarder.ReadOnly = true;
                        slueDoOIDVEND.ReadOnly = true;
                        dteDoETAWH.ReadOnly = true;
                        dteDoContractedETD.ReadOnly = true;

                        gvEntryDO.OptionsBehavior.ReadOnly = true;
                        gvEntryDO.OptionsBehavior.Editable = false;

                    }
                }
            }
        }

        private void spePoRevisionNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txePoLot.Focus();
        }

        private void txePoLot_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                dtePoRevisedDate.Focus();
        }

        private void dtePoRevisedDate_EditValueChanged(object sender, EventArgs e)
        {
            gluePoDocumentStatus.Focus();
        }

        private void gluePoDocumentStatus_EditValueChanged(object sender, EventArgs e)
        {
            //txePoOrderPlanNumber.Focus();
            setDefaultStatusPO();
        }

        private void txePoOrderPlanNumber_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                sluePoOIDCUST.Focus();
        }

        private void sluePoOIDCUST_EditValueChanged(object sender, EventArgs e)
        {
            if (sluePoOIDCUST.Text.Trim() != "")
            {
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT CT.PaymentTerm, CC.OIDCURR ");
                sbSQL.Append("FROM Customer AS CT LEFT OUTER JOIN ");
                sbSQL.Append("     Currency AS CC ON CT.PaymentCurrency = CC.Currency ");
                sbSQL.Append("WHERE (CT.OIDCUST = '" + sluePoOIDCUST.EditValue.ToString() + "') ");
                string[] arrCUS = new DBQuery(sbSQL).getMultipleValue();
                if (arrCUS.Length > 0)
                {
                    sluePoPaymentTerms.EditValue = arrCUS[0];
                    gluePoOIDCURR.EditValue = arrCUS[1];
                }
            }
            spePoSeason.Focus();
        }

        private void spePoSeason_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                gluePoSeason.Focus();
        }

        private void gluePoSeason_EditValueChanged(object sender, EventArgs e)
        {
            sluePoItemCode.Focus();
        }

        private void gluePoBusinessUnit_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txePoApprover.Focus();
        }

        private void txePoPatternDimensionCode_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                sluePoItemCode.Focus();
        }

        private void sluePoItemCode_EditValueChanged(object sender, EventArgs e)
        {
            if (sluePoItemCode.Text.Trim() != "")
            {
                txePoItemName.Text = "";
                txePoFabricWidth.Text = "";
                txePoFBComposition.Text = "";
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT ItemCode, ItemName, FabricWidth, FBComposition ");
                sbSQL.Append("FROM   ItemCustomer ");
                sbSQL.Append("WHERE (OIDCSITEM = '" + sluePoItemCode.EditValue.ToString() + "') ");
                string[] arrITEM = new DBQuery(sbSQL).getMultipleValue();
                if (arrITEM.Length > 0)
                {
                    txePoItemName.Text = arrITEM[1];
                    txePoFabricWidth.Text = arrITEM[2];
                    txePoFBComposition.Text = arrITEM[3];
                }
            }
            txePoRemark.Focus();
        }

        private void txePoOrderQtyPCS_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txePoOriginalSalesPrice.Focus();
        }

        private void txePoOriginalSalesPrice_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                gluePoBusinessUnit.Focus();
        }

        private void txePoApprover_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                dtePoApprovalDate.Focus();
        }

        private void dtePoApprovalDate_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                sluePoOIDBillto.Focus();
        }

        private void sluePoOIDBillto_EditValueChanged(object sender, EventArgs e)
        {
            if (sluePoOIDBillto.Text.Trim() != "")
            {
                txePoAddress.Text = "";
                txePoTelephoneNo.Text = "";
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT Address1 + ' ' + Address2 + ' ' + Address3 + ' ' + City + ' ' + Country AS Address, TelephoneNo ");
                sbSQL.Append("FROM   Vendor ");
                sbSQL.Append("WHERE (VendorType = 6) AND (OIDVEND = '" + sluePoOIDBillto.EditValue.ToString() + "') ");
                string[] arrVEND = new DBQuery(sbSQL).getMultipleValue();
                if (arrVEND.Length > 0)
                {
                    txePoAddress.Text = arrVEND[0];
                    txePoTelephoneNo.Text = arrVEND[1];
                }
            }
            sluePoPaymentTerms.Focus();
        }

        private void sluePoPaymentTerms_EditValueChanged(object sender, EventArgs e)
        {
            gluePoOIDCURR.Focus();
        }

        private void gluePoOIDCURR_EditValueChanged(object sender, EventArgs e)
        {
            txePoRemark.Focus();
        }

        private void txePoRemark_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txePoAllocationOrderNumber.Focus();
        }

        private void txePoAllocationOrderNumber_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txePoOriginalSalesPrice.Focus();
        }

        private void LoadPODetail(string PONO)
        {
            gcEntryDO.DataSource = null;
            PONO = PONO.ToUpper().Trim();
            if (PONO != "")
            {
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT COPO.OIDCOLOR, PC.ColorNo, PC.ColorName, COPO.OIDSIZE, PS.SizeNo, PS.SizeName, COPO.PatternDimensionCode, '' AS SetCode, 0 AS QtyBox, 0 AS QtySet, 0 AS PickUnit ");
                sbSQL.Append("FROM   COPO LEFT OUTER JOIN ");
                sbSQL.Append("       ProductColor AS PC ON COPO.OIDCOLOR = PC.OIDCOLOR LEFT OUTER JOIN ");
                sbSQL.Append("       ProductSize AS PS ON COPO.OIDSIZE = PS.OIDSIZE ");
                sbSQL.Append("WHERE (COPO.OrderNo = N'" + PONO + "') ");
                sbSQL.Append("ORDER BY COPO.OID ");
                new ObjDevEx.setGridControl(gcEntryDO, gvEntryDO, sbSQL).getData(false, false, false, true);
            }
        }

        private void speDoRevisionNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txeDoDONo.Focus();
        }

        private void txeDoDONo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeDoDONo.Text = txeDoDONo.Text.ToUpper().Trim();
                if (slueDoOrderNo.Text.Trim() != "")
                {
                    string DONO = txeDoDONo.Text.ToUpper().Trim();
                    LOADDO(DONO);
                }
                slueDoOrderNo.Focus();
            }

        }

        private void glueDoDocStatus_EditValueChanged(object sender, EventArgs e)
        {
            setDefaultStatusDO();
            slueDoItemCode.Focus();
        }

        private void slueDoItemCode_EditValueChanged(object sender, EventArgs e)
        {
            glueDoTransportationMethod.Focus();
        }

        private void glueDoTransportationMethod_EditValueChanged(object sender, EventArgs e)
        {
            slueDoPortCode.Focus();
        }

        private void LoadPortCode(string PortID)
        {
            PortID = PortID.Trim();
            txeDoPortName.Text = "";
            txeDoWHName.Text = "";
            txeDoAddress.Text = "";
            txeDoTelephone.Text = "";
            txeDoPersonInCharge.Text = "";
            if (PortID != "")
            {
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT PortCode, PortName, WHName, Address, Telephone, PersonInCharge ");
                sbSQL.Append("FROM   PortAndCity ");
                sbSQL.Append("WHERE (OIDPORT = '" + PortID + "') ");
                string[] arrPort = new DBQuery(sbSQL).getMultipleValue();
                if (arrPort.Length > 0)
                {
                    txeDoPortName.Text = arrPort[1];
                    txeDoWHName.Text = arrPort[2];
                    txeDoAddress.Text = arrPort[3];
                    txeDoTelephone.Text = arrPort[4];
                    txeDoPersonInCharge.Text = arrPort[5];
                }
            }
        }

        private void slueDoPortCode_EditValueChanged(object sender, EventArgs e)
        {
            if(slueDoPortCode.Text.Trim() != "")
                LoadPortCode(slueDoPortCode.EditValue.ToString());
            glueDoIncoterms.Focus();
        }

        private void txeDoIncoterms_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txeDoForwarder.Focus();
        }

        private void txeDoForwarder_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                slueDoOIDVEND.Focus();
        }

        private void slueDoOIDVEND_EditValueChanged(object sender, EventArgs e)
        {
            dteDoETAWH.Focus();
        }

        private void dteDoETAWH_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                dteDoContractedETD.Focus();
        }

        private void dteDoContractedETD_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txeCREATE.Focus();
        }

        private void cbePoSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txePoFilePath.Text.Trim() != "" && cbePoSheet.Text.Trim() != "")
            {
                if (txeDoFilePath.Text == "")
                {
                    cbeDoSheet.Properties.Items.Clear();
                    cbeDoSheet.Text = "";
                    spsDO.CloseCellEditor(DevExpress.XtraSpreadsheet.CellEditorEnterValueMode.Default);
                    spsDO.CreateNewDocument();
                }

                IWorkbook workbook = spsPO.Document;

                try
                {
                    using (FileStream stream = new FileStream(txePoFilePath.Text, FileMode.Open))
                    {
                        // workbook.CalculateFull();
                        string ext = Path.GetExtension(txePoFilePath.Text);
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
                            if (cbePoSheet.Text == "2 First Sheet (PO & DO)")
                            {
                                //PO
                                for (int i = workbook.Worksheets.Count - 1; i >= 0; i--)
                                {
                                    if (workbook.Worksheets[i].Name.IndexOf("PO") == -1)
                                        workbook.Worksheets.RemoveAt(i);
                                }
                            }
                            else
                            {
                                for (int i = workbook.Worksheets.Count - 1; i >= 0; i--)
                                {
                                    if (workbook.Worksheets[i].Name != cbePoSheet.Text)
                                        workbook.Worksheets.RemoveAt(i);
                                }
                            }

                        }
                    }

                    if (cbePoSheet.Text == "2 First Sheet (PO & DO)")
                    {
                        txeDoFilePath.Text = "";
                        cbeDoSheet.Properties.Items.Clear();
                        cbeDoSheet.Text = "";
                        spsDO.CloseCellEditor(DevExpress.XtraSpreadsheet.CellEditorEnterValueMode.Default);
                        spsDO.CreateNewDocument();

                        IWorkbook workbookDO = spsDO.Document;
                        using (FileStream stream = new FileStream(txePoFilePath.Text, FileMode.Open))
                        {
                            // workbook.CalculateFull();
                            string ext = Path.GetExtension(txePoFilePath.Text);
                            if (ext == ".xlsx")
                                workbookDO.LoadDocument(stream, DocumentFormat.Xlsx);
                            else if (ext == ".xls")
                                workbookDO.LoadDocument(stream, DocumentFormat.Xls);
                            else if (ext == ".csv")
                                workbookDO.LoadDocument(stream, DocumentFormat.Csv);
                            //workbook.Worksheets.ActiveWorksheet = workbook.Worksheets[0];

                            //***Delete sheet
                            if (workbookDO.Worksheets.Count > 0)
                            {
                                //DO
                                for (int i = workbookDO.Worksheets.Count - 1; i >= 0; i--)
                                {
                                    if (workbookDO.Worksheets[i].Name.IndexOf("DO") == -1)
                                        workbookDO.Worksheets.RemoveAt(i);
                                }
                            }
                        }
                    }
                }
                catch (Exception)
                {
                    FUNC.msgWarning("Please close excel file before import.");
                    txePoFilePath.Text = "";
                }
            }
        }

        private void cbeDoSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txeDoFilePath.Text.Trim() != "" && cbeDoSheet.Text.Trim() != "")
            {
                IWorkbook workbook = spsDO.Document;

                try
                {
                    using (FileStream stream = new FileStream(txeDoFilePath.Text, FileMode.Open))
                    {
                        // workbook.CalculateFull();
                        string ext = Path.GetExtension(txeDoFilePath.Text);
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
                            if (cbeDoSheet.Text == "2 First Sheet (Do & DO)")
                            {
                                //Do
                                for (int i = workbook.Worksheets.Count - 1; i >= 0; i--)
                                {
                                    if (workbook.Worksheets[i].Name.IndexOf("Do") == -1)
                                        workbook.Worksheets.RemoveAt(i);
                                }
                            }
                            else
                            {
                                for (int i = workbook.Worksheets.Count - 1; i >= 0; i--)
                                {
                                    if (workbook.Worksheets[i].Name != cbeDoSheet.Text)
                                        workbook.Worksheets.RemoveAt(i);
                                }
                            }

                        }
                    }
                }
                catch (Exception)
                {
                    FUNC.msgWarning("Please close excel file before imDort.");
                    txeDoFilePath.Text = "";
                }
            }
        }

        private void gvEntryPO_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {

        }

        private void repositoryItemSearchLookUpEdit2_EditValueChanged(object sender, EventArgs e)
        {

        }

        private void repositoryItemSearchLookUpEdit3_EditValueChanged(object sender, EventArgs e)
        {
            DevExpress.XtraEditors.TextEdit textEditor = (DevExpress.XtraEditors.TextEdit)sender;
            string strID = textEditor.EditValue.ToString();
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT ColorName ");
            sbSQL.Append("FROM   ProductColor ");
            sbSQL.Append("WHERE (OIDCOLOR = '" + strID + "') ");
            string[] arrITEM = new DBQuery(sbSQL).getMultipleValue();

            int rowHandle = gvEntryPO.FocusedRowHandle;
            if (arrITEM.Length > 0)
            {
                gvEntryPO.SetRowCellValue(rowHandle, "ColorName", arrITEM[0]);
            }
            else
            {
                gvEntryPO.SetRowCellValue(rowHandle, "ColorName", "");
            }
        }

        private void repositoryItemSearchLookUpEdit4_EditValueChanged(object sender, EventArgs e)
        {
            DevExpress.XtraEditors.TextEdit textEditor = (DevExpress.XtraEditors.TextEdit)sender;
            string strID = textEditor.EditValue.ToString();
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT SizeName ");
            sbSQL.Append("FROM   ProductSize ");
            sbSQL.Append("WHERE (OIDSIZE = '" + strID + "') ");
            string[] arrITEM = new DBQuery(sbSQL).getMultipleValue();

            int rowHandle = gvEntryPO.FocusedRowHandle;
            if (arrITEM.Length > 0)
            {
                gvEntryPO.SetRowCellValue(rowHandle, "SizeName", arrITEM[0]);
            }
            else
            {
                gvEntryPO.SetRowCellValue(rowHandle, "SizeName", "");
            }
        }

        private void repositoryItemSearchLookUpEdit5_EditValueChanged(object sender, EventArgs e)
        {
            //DevExpress.XtraEditors.TextEdit textEditor = (DevExpress.XtraEditors.TextEdit)sender;
            //string strID = textEditor.EditValue.ToString();
            //StringBuilder sbSQL = new StringBuilder();
            //sbSQL.Append("SELECT Address1 + ' ' + Address2 + ' ' + Address3 + ' ' + City + ' ' + Country AS Address, TelephoneNo AS Telephone ");
            //sbSQL.Append("FROM   Vendor ");
            //sbSQL.Append("WHERE (VendorType = 6) AND (OIDVEND = '" + strID + "') ");
            //string[] arrITEM = new DBQuery(sbSQL).getMultipleValue();

            //int rowHandle = gvEntryPO.FocusedRowHandle;
            //if (arrITEM.Length > 0)
            //{
            //    gvEntryPO.SetRowCellValue(rowHandle, "Address", arrITEM[0]);
            //    gvEntryPO.SetRowCellValue(rowHandle, "TelephoneNo", arrITEM[1]);

            //}
            //else
            //{
            //    gvEntryPO.SetRowCellValue(rowHandle, "Address", "");
            //    gvEntryPO.SetRowCellValue(rowHandle, "TelephoneNo", "");
            //}

        }

        private bool chkDupColorSize(string COLOR, string SIZE, int rowIndex)
        {
            //gvEntryPO.CloseEditor();
            //gvEntryPO.UpdateCurrentRow();

            COLOR = COLOR.ToUpper().Trim();
            SIZE = SIZE.ToUpper().Trim();

            bool chkDup = true;

            if (COLOR != "" && SIZE != "")
            {
                int countPlan = 0;
                DataTable dtFind = (DataTable)gcEntryPO.DataSource;
                int xRow = 0;
                foreach (DataRow row in dtFind.Rows)
                {
                    string chkCOLOR = row["OIDCOLOR"].ToString().ToUpper().Trim();
                    string chkSIZE = row["OIDSIZE"].ToString().ToUpper().Trim();
                    if (chkCOLOR == COLOR && chkSIZE == SIZE && xRow != rowIndex)
                        countPlan++;
                    xRow++;
                }
                //MessageBox.Show(countPlan.ToString());
                if (countPlan > 0)
                    chkDup = false;
            }
            return chkDup;
        }

        private void gvEntryPO_ValidateRow(object sender, DevExpress.XtraGrid.Views.Base.ValidateRowEventArgs e)
        {
            GridView view = sender as GridView;
            DevExpress.XtraGrid.Columns.GridColumn COLORCol = view.Columns["OIDCOLOR"];
            DevExpress.XtraGrid.Columns.GridColumn SIZECol = view.Columns["OIDSIZE"];
            string strCOLOR = Convert.ToString((Int32)view.GetRowCellValue(e.RowHandle, COLORCol));
            string strSIZE = Convert.ToString((Int32)view.GetRowCellValue(e.RowHandle, SIZECol));
            //MessageBox.Show(strCOLOR + ", " + strSIZE);
            bool chkPlan = chkDupColorSize(strCOLOR, strSIZE, e.RowHandle);
            //Validity criterion
            if (chkPlan == false)
            {
                e.Valid = false;
                //Set errors with specific descriptions for the columns
                view.SetColumnError(COLORCol, "Duplicate color & size. !! Please change.");
                view.SetColumnError(SIZECol, "Duplicate color & size. !! Please change.");
            }
        }

        private void gvEntryPO_InvalidRowException(object sender, DevExpress.XtraGrid.Views.Base.InvalidRowExceptionEventArgs e)
        {
            e.ExceptionMode = DevExpress.XtraEditors.Controls.ExceptionMode.NoAction;
        }

        //DO ENTRY
        private void repositoryItemSearchLookUpEdit7_EditValueChanged(object sender, EventArgs e)
        {

        }

        private void repositoryItemSearchLookUpEdit8_EditValueChanged(object sender, EventArgs e)
        {

        }

        private void gvEntryDO_ValidateRow(object sender, DevExpress.XtraGrid.Views.Base.ValidateRowEventArgs e)
        {

        }

        private void gvEntryDO_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator) e.Info.DisplayText = (e.RowHandle + 1).ToString();
        }

        private void slueDoOrderNo_EditValueChanged(object sender, EventArgs e)
        {
            string DONO = "";
            if (slueDoOrderNo.Text.Trim() != "")
            {
                GridView view = slueDoOrderNo.Properties.View;
                int rowHandle = view.FocusedRowHandle;
                if (rowHandle >= 0)
                {
                    string fieldName = "Revision No."; // or other field name  
                    object value = view.GetRowCellValue(rowHandle, fieldName);
                    speDoRevisionNo.Value = Convert.ToInt32(value);
                    if (glueDoDocStatus.Text.Trim() != "")
                    {
                        if (glueDoDocStatus.EditValue.ToString() == "0") //new
                            DONO = txeDoDONo.Text.ToUpper().Trim();
                        else
                            DONO = slueDoDONo.EditValue.ToString().ToUpper().Trim();
                    }
                }
                else if(slueDoDONo.Text.Trim() != "")
                {
                    speDoRevisionNo.Value = new DBQuery("SELECT TOP(1) RevisionNo FROm CODO WHERE (DONo = N'" + slueDoDONo.Text.Trim() + "')").getInt();
                    DONO = slueDoDONo.EditValue.ToString().ToUpper().Trim();
                }
                LoadDOPO(slueDoOrderNo.Text, speDoRevisionNo.Value.ToString(), DONO);
            }
            glueDoTransportationMethod.Focus();

        }

        private void sluePoOrderNo_EditValueChanged(object sender, EventArgs e)
        {
            if (sluePoOrderNo.Text.Trim() != "")
            {
                GridView view = sluePoOrderNo.Properties.View;
                int rowHandle = view.FocusedRowHandle;
                string fieldName = "Revision No."; // or other field name  
                object value = view.GetRowCellValue(rowHandle, fieldName);
                spePoRevisionNo.Value = Convert.ToInt32(value);
                if (gluePoDocumentStatus.EditValue.ToString() != "1") //not revise
                    LoadPORevision(sluePoOrderNo.Text, spePoRevisionNo.Value.ToString());
                else
                {
                    //spePoRevisionNo.Value = Convert.ToInt32(value) + 1;
                    spePoRevisionNo.Value = new DBQuery("SELECT MAX(RevisionNo)+1 AS Revise FROM COPO WHERE (OrderNo = N'" + sluePoOrderNo.Text.Trim() + "') ").getInt();
                    dtePoRevisedDate.EditValue = DateTime.Now;
                    LoadPORevision(sluePoOrderNo.Text);
                }
            }
        }

        private void LoadPORevision(string PONO, string REVISE = "")
        {
            PONO = PONO.ToUpper().Trim();
            REVISE = REVISE.Trim();
            if (PONO != "")
            {
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT DISTINCT Lot, RevisedDate, OrderPlanNumber, OIDCUST, Season, BusinessUnit, ItemCode, Remark, AllocationOrderNumber, OriginalSalesPrice, ");
                sbSQL.Append("       Approver, ApprovalDate, OIDBillto, PaymentTerms, OIDCURR ");
                sbSQL.Append("FROM  COPO ");
                sbSQL.Append("WHERE (OrderNo = N'" + PONO + "') ");
                if (REVISE != "")
                    sbSQL.Append("AND (RevisionNo = '" + REVISE + "') ");
                else
                    sbSQL.Append("AND (RevisionNo = (SELECT MAX(RevisionNo) AS RevisionNo FROM COPO WHERE (OrderNo = N'" + PONO + "'))) ");
                string[] arrPO = new DBQuery(sbSQL).getMultipleValue();
                if (arrPO.Length > 0)
                {
                    txePoLot.Text = arrPO[0];


                    if (REVISE != "")
                    {
                        if (arrPO[1] != "")
                            dtePoRevisedDate.EditValue = Convert.ToDateTime(arrPO[1]);
                        else
                            dtePoRevisedDate.EditValue = null;
                    }
                    else
                        dtePoRevisedDate.EditValue = DateTime.Now;

                    txePoOrderPlanNumber.Text = arrPO[2];
                    sluePoOIDCUST.EditValue = arrPO[3];

                    string Season = arrPO[4];
                    spePoSeason.Value = Convert.ToInt32(Regex.Match(Season, @"\d+([,\.]\d+)?").Value);
                    gluePoSeason.EditValue = arrPO[4].Replace(Regex.Match(Season, @"\d+([,\.]\d+)?").Value, "");

                    gluePoBusinessUnit.EditValue = arrPO[5];
                    sluePoItemCode.EditValue = arrPO[6];
                    txePoRemark.Text = arrPO[7];
                    txePoAllocationOrderNumber.Text = arrPO[8];
                    txePoOriginalSalesPrice.Text = arrPO[9];
                    txePoApprover.Text = arrPO[10];

                    if (arrPO[11] != "")
                        dtePoApprovalDate.EditValue = Convert.ToDateTime(arrPO[11]);
                    else
                        dtePoApprovalDate.EditValue = null;

                    sluePoOIDBillto.EditValue = arrPO[12];
                    //sluePoPaymentTerms.EditValue = arrPO[13];
                    //gluePoOIDCURR.EditValue = arrPO[14];

                    sbSQL.Clear();
                    sbSQL.Append("SELECT PO.OIDCOLOR, PC.ColorName, PO.OIDSIZE, PS.SizeName, PO.PatternDimensionCode, PO.SKUCode, PO.SampleCode, PO.OrderQtyPCS, PO.OID ");
                    sbSQL.Append("FROM COPO AS PO LEFT OUTER JOIN ");
                    sbSQL.Append("     ProductColor AS PC ON PO.OIDCOLOR = PC.OIDCOLOR LEFT OUTER JOIN ");
                    sbSQL.Append("     ProductSize AS PS ON PO.OIDSIZE = PS.OIDSIZE ");
                    sbSQL.Append("WHERE (PO.OrderNo = N'" + PONO + "') ");
                    if (REVISE != "")
                        sbSQL.Append("AND (PO.RevisionNo = '" + REVISE + "') ");
                    else
                        sbSQL.Append("AND (PO.RevisionNo = (SELECT MAX(RevisionNo) AS RevisionNo FROM COPO WHERE (OrderNo = N'" + PONO + "'))) ");
                    sbSQL.Append("ORDER BY PO.OID ");
                    new ObjDevEx.setGridControl(gcEntryPO, gvEntryPO, sbSQL).getData(false, false, false, true);
                }
            }
        }

        private void LOADDO(string DONO)
        {
            glueDoTransportationMethod.Text = "";
            glueDoIncoterms.EditValue = "0";
            txeDoForwarder.Text = "";
            slueDoOIDVEND.EditValue = "";
            dteDoETAWH.EditValue = null;
            dteDoContractedETD.EditValue = null;
            txeCREATE.Text = "";
            txeDATE.Text = "";

            DONO = DONO.ToUpper().Trim();
            if (DONO != "")
            {
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("IF EXISTS(SELECT OIDDO FROM CODO WHERE DONo = N'" + DONO + "') ");
                sbSQL.Append(" BEGIN ");
                sbSQL.Append("  SELECT TOP(1) OrderNo + '_' + CONVERT(NVARCHAR, RevisionNo) AS POID, RevisionNo, ItemCode, TransportationMethod, PortCode, ");
                sbSQL.Append("         Incoterms, Forwarder, OIDVEND, ETAWH, ContractedETD, UpdatedBy, UpdatedDate ");
                sbSQL.Append("  FROM CODO ");
                sbSQL.Append("  WHERE (DONo = N'" + DONO + "') ");
                sbSQL.Append(" END ");
                string[] arrDO = new DBQuery(sbSQL).getMultipleValue();
                //MessageBox.Show(arrDO.Length.ToString());
                if (arrDO.Length > 0)
                {
                    slueDoOrderNo.EditValue = arrDO[0];
                    speDoRevisionNo.Value = Convert.ToInt32(arrDO[1]);
                    glueDoTransportationMethod.Text = arrDO[3];
                    slueDoPortCode.EditValue = arrDO[4];
                    glueDoIncoterms.EditValue = arrDO[5];
                    txeDoForwarder.Text = arrDO[6];
                    slueDoOIDVEND.EditValue = arrDO[7];

                    if (arrDO[8] != "")
                        dteDoETAWH.EditValue = Convert.ToDateTime(arrDO[8]);
                    else
                        dteDoETAWH.EditValue = null;

                    if (arrDO[9] != "")
                        dteDoContractedETD.EditValue = Convert.ToDateTime(arrDO[9]);
                    else
                        dteDoContractedETD.EditValue = null;

                    txeCREATE.Text = arrDO[10];
                    txeDATE.Text = arrDO[11];
                }
            }

        }

        private void LoadDOPO(string PONO, string REVISE, string DONO)
        {
            PONO = PONO.ToUpper().Trim();
            REVISE = REVISE.Trim();
            DONO = DONO.ToUpper().Trim();

            slueDoItemCode.EditValue = "";
            txeDoItemName.Text = "";
            if (PONO != "")
            {
                StringBuilder sbPO = new StringBuilder();
                sbPO.Append("SELECT TOP(1) PO.ItemCode, IC.ItemName FROM COPO AS PO LEFT OUTER JOIN ItemCustomer AS IC ON PO.ItemCode = IC.OIDCSITEM WHERE (PO.OrderNo = N'" + PONO + "') AND (PO.RevisionNo = '" + REVISE + "') ");
                string[] arrPO = new DBQuery(sbPO).getMultipleValue();
                if (arrPO.Length > 0)
                {
                    slueDoItemCode.EditValue = arrPO[0];
                    txeDoItemName.Text = arrPO[1];
                }
            }

            if (PONO != "" || DONO != "")
            {
                StringBuilder sbSQL = new StringBuilder();
                if (PONO != "" && DONO != "")
                {
                    sbSQL.Append("SELECT PO.OIDCOLOR, PC.ColorName, PO.OIDSIZE, PS.SizeName, CASE WHEN ISNULL(DO.PatternDimensionCode, N'') = N'' THEN PO.PatternDimensionCode ELSE DO.PatternDimensionCode END AS PatternDimensionCode, ");
                    sbSQL.Append("       CASE WHEN ISNULL(DO.SetCode, N'') = N'' THEN N'' ELSE DO.SetCode END AS SetCode, CASE WHEN DO.QuantityBox IS NULL THEN 0 ELSE DO.QuantityBox END AS QtyBox,  ");
                    sbSQL.Append("       CASE WHEN DO.QtyperSet IS NULL THEN 0 ELSE DO.QtyperSet END AS QtySet, CASE WHEN DO.PickingUnit IS NULL THEN 0 ELSE DO.PickingUnit END AS PickUnit, CASE WHEN DO.OIDDO IS NULL THEN '' ELSE CONVERT(VARCHAR, DO.OIDDO) END AS OID ");
                    sbSQL.Append("FROM   COPO AS PO LEFT OUTER JOIN ");
                    sbSQL.Append("       ProductColor AS PC ON PO.OIDCOLOR = PC.OIDCOLOR LEFT OUTER JOIN ");
                    sbSQL.Append("       ProductSize AS PS ON PO.OIDSIZE = PS.OIDSIZE LEFT OUTER JOIN ");
                    sbSQL.Append("       (SELECT OrderNo, RevisionNo, OIDCOLOR, OIDSIZE, PatternDimensionCode, SetCode, QuantityBox, QtyperSet, PickingUnit, OIDDO ");
                    sbSQL.Append("        FROM   CODO ");
                    sbSQL.Append("        WHERE (DONo = N'" + DONO + "')) AS DO ON PO.OrderNo = DO.OrderNo AND PO.RevisionNo = DO.RevisionNo AND PO.OIDCOLOR = DO.OIDCOLOR AND PO.OIDSIZE = DO.OIDSIZE ");
                    sbSQL.Append("WHERE (PO.OrderNo = N'" + PONO + "') AND(PO.RevisionNo = '" + REVISE + "') ");
                    sbSQL.Append("ORDER BY PO.OID ");
                }
                else if (PONO != "")
                {
                    sbSQL.Append("SELECT PO.OIDCOLOR, PC.ColorName, PO.OIDSIZE, PS.SizeName, PO.PatternDimensionCode, ");
                    sbSQL.Append("       N'' AS SetCode, 0 AS QtyBox,  ");
                    sbSQL.Append("       0 AS QtySet, 0 AS PickUnit, '' AS OID ");
                    sbSQL.Append("FROM   COPO AS PO LEFT OUTER JOIN ");
                    sbSQL.Append("       ProductColor AS PC ON PO.OIDCOLOR = PC.OIDCOLOR LEFT OUTER JOIN ");
                    sbSQL.Append("       ProductSize AS PS ON PO.OIDSIZE = PS.OIDSIZE ");
                    sbSQL.Append("WHERE(PO.OrderNo = N'" + PONO + "') AND(PO.RevisionNo = '" + REVISE + "') ");
                    sbSQL.Append("ORDER BY PO.OID ");
                }
                //MessageBox.Show(sbSQL.ToString());
                new ObjDevEx.setGridControl(gcEntryDO, gvEntryDO, sbSQL).getData(false, false, false, true);
            }
        }

        private void gluePoBusinessUnit_EditValueChanged(object sender, EventArgs e)
        {
            
        }

        private void slueDoDONo_EditValueChanged(object sender, EventArgs e)
        {
            if (slueDoDONo.Text.Trim() != "")
            {
                string DONO = slueDoDONo.EditValue.ToString().ToUpper().Trim();
                LOADDO(DONO);
            }
            slueDoOrderNo.Focus();
        }

        private void txePoOrderNo_EditValueChanged(object sender, EventArgs e)
        {

        }

        private void txeDoDONo_Leave(object sender, EventArgs e)
        {
            txeDoDONo.Text = txeDoDONo.Text.ToUpper().Trim();
            if (txeDoDONo.Text.Trim() != "")
            {
                bool chkDup = chkDuplicateDO(txeDoDONo.Text);
                if (chkDup == false) //Load PO
                {
                    txeDoDONo.Text = "";
                    txeDoDONo.Focus();
                    FUNC.msgWarning("Duplicate DO No. !! Please Change.");
                }
            }
        }

        private bool chkDuplicateDO(string DO)
        {
            DO = DO.ToUpper().Trim().Replace("'", "''");
            bool chkDup = true;
            if (DO != "")
            {
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT TOP(1) DONo FROM CODO WHERE (DONo = N'" + DO + "') ");
                if (new DBQuery(sbSQL).getString() != "")
                {
                    chkDup = false;
                }
            }
            return chkDup;
        }

        private void gvSumPO_DoubleClick(object sender, EventArgs e)
        {
            GridView view = (GridView)sender;
            Point pt = view.GridControl.PointToClient(Control.MousePosition);
            DevExpress.XtraGrid.Views.Grid.ViewInfo.GridHitInfo info = view.CalcHitInfo(pt);
            if (info.InRow || info.InRowCell)
            {
                DataTable dtPO = (DataTable)gcSumPO.DataSource;
                if (dtPO.Rows.Count > 0)
                {
                    DataRow drPO = dtPO.Rows[info.RowHandle];
                    string PONO = drPO["PO Order No."].ToString();
                    string REVISE = drPO["Revision No."].ToString();
                    tabbedControlGroup3.SelectedTabPage = layoutControlGroup5;
                    tabbedControlGroup4.SelectedTabPage = layoutControlGroup8;
                    gluePoDocumentStatus.EditValue = 2;
                    string POID = PONO + "_" + REVISE;
                    sluePoOrderNo.EditValue = POID;
                    spePoRevisionNo.Value = Convert.ToInt32(REVISE);
                    LoadPORevision(PONO, REVISE);
                }
            }
        }


        private void gvSumDO_DoubleClick(object sender, EventArgs e)
        {
            GridView view = (GridView)sender;
            Point pt = view.GridControl.PointToClient(Control.MousePosition);
            DevExpress.XtraGrid.Views.Grid.ViewInfo.GridHitInfo info = view.CalcHitInfo(pt);
            if (info.InRow || info.InRowCell)
            {
                DataTable dtDO = (DataTable)gcSumDO.DataSource;
                if (dtDO.Rows.Count > 0)
                {
                    DataRow drDO = dtDO.Rows[info.RowHandle];
                    string DOID = drDO["OID DO"].ToString();
                    tabbedControlGroup3.SelectedTabPage = layoutControlGroup5;
                    tabbedControlGroup4.SelectedTabPage = layoutControlGroup9;
                    glueDoDocStatus.EditValue = 2;
                    //LoadForecast("Edit", ID);
                }
            }
        }

        private void bbiRefresh_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            LoadSummary();
        }

        private void gvSumDO_MasterRowExpanded(object sender, CustomMasterRowEventArgs e)
        {
            DevExpress.XtraGrid.Views.Grid.GridView masterView = (DevExpress.XtraGrid.Views.Grid.GridView)sender;
            DevExpress.XtraGrid.Views.Grid.GridView detailView = (DevExpress.XtraGrid.Views.Grid.GridView)masterView.GetDetailView(e.RowHandle, e.RelationIndex);

            detailView.Columns["DO No."].Visible = false;
            detailView.Columns["PO Order No."].Visible = false;
            detailView.Columns["Revision No."].Visible = false;
            detailView.Columns["OIDCOLOR"].Visible = false;
            detailView.Columns["OIDSIZE"].Visible = false;
            detailView.Columns["QuantityBox"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            detailView.Columns["QtyperSet"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            detailView.Columns["PickingUnit"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;

            detailView.OptionsView.ColumnAutoWidth = false;
            detailView.BestFitColumns();
        }

        private void gvSumPO_MasterRowExpanded(object sender, CustomMasterRowEventArgs e)
        {
            DevExpress.XtraGrid.Views.Grid.GridView masterView = (DevExpress.XtraGrid.Views.Grid.GridView)sender;
            DevExpress.XtraGrid.Views.Grid.GridView detailView = (DevExpress.XtraGrid.Views.Grid.GridView)masterView.GetDetailView(e.RowHandle, e.RelationIndex);

            detailView.Columns["PO Order No."].Visible = false;
            detailView.Columns["Revision No."].Visible = false;
            detailView.Columns["OIDCOLOR"].Visible = false;
            detailView.Columns["OIDSIZE"].Visible = false;
            detailView.Columns["Order Qty. (Pcs)"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;

            detailView.OptionsView.ColumnAutoWidth = false;
            detailView.BestFitColumns();
        }

        private void gvSumPO_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator) e.Info.DisplayText = (e.RowHandle + 1).ToString();
        }

        private void gvSumDO_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator) e.Info.DisplayText = (e.RowHandle + 1).ToString();
        }

        private void ribbonControl_Click(object sender, EventArgs e)
        {

        }
    }

    public class DocumentStatus
    { 
        public int ID { get; set; }
        public string Status { get; set; }
    }

    public class TransportationMethod
    {
        public int ID { get; set; }
        public string TransportMethod { get; set; }
    }
}