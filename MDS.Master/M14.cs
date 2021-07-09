using System;
using System.Text;
using DBConnect;
using System.Windows.Forms;
using DevExpress.LookAndFeel;
using DevExpress.Utils.Extensions;
using System.Drawing;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraEditors;
using DevExpress.XtraPrinting;
using TheepClass;
using DevExpress.Utils;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;

namespace MDS.Master
{
    public partial class M14 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private Functionality.Function FUNC = new Functionality.Function();
        bool chkShow = false;
        public LogIn UserLogin { get; set; }

        public string ConnectionString { get; set; }
        string CONNECT_STRING = "";
        DatabaseConnect DBC;

        public M14()
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
            sbSQL.Append("SELECT TOP (1) ReadWriteStatus FROM FunctionAccess WHERE (OIDUser = '" + UserLogin.OIDUser + "') AND(FunctionNo = 'M14') ");
            int chkReadWrite = this.DBC.DBQuery(sbSQL).getInt();
            if (chkReadWrite == 0)
                ribbonPageGroup1.Visible = false;

            sbSQL.Clear();
            sbSQL.Append("SELECT FullName, OIDUSER FROM Users ORDER BY OIDUSER ");
            new ObjDE.setGridLookUpEdit(glueCREATE, sbSQL, "FullName", "OIDUSER").getData();
            new ObjDE.setGridLookUpEdit(glueUPDATE, sbSQL, "FullName", "OIDUSER").getData();

            glueCREATE.EditValue = UserLogin.OIDUser;
            glueUPDATE.EditValue = UserLogin.OIDUser;

            bbiNew.PerformClick();
        }

        private void LoadData()
        {
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT Code, Name, ShortName, OIDCUST AS ID ");
            sbSQL.Append("FROM  Customer ");
            sbSQL.Append("ORDER BY Name ");
            new ObjDE.setSearchLookUpEdit(slueCode, sbSQL, "Name", "ID").getData(true);

            sbSQL.Clear();
            sbSQL.Append("SELECT CUS.Code, CUS.Name, ADDR.Code AS DestinationCode, ADDR.ShipToName, ADDR.ShipToAddress1, ADDR.ShipToAddress2, ADDR.ShipToAddress3, ADDR.Country, ADDR.PostCode, ADDR.TelephoneNo, ADDR.FaxNo, ");
            sbSQL.Append("       ADDR.CreatedBy, ADDR.CreatedDate, ADDR.UpdatedBy, ADDR.UpdatedDate, ADDR.OIDCUST AS CUSID, ADDR.OIDCUSTAdd AS ID ");
            sbSQL.Append("FROM   Customer AS CUS INNER JOIN ");
            sbSQL.Append("       CustomerAddress AS ADDR ON CUS.OIDCUST = ADDR.OIDCUST ");
            sbSQL.Append("ORDER BY CUS.Code, DestinationCode ");
            new ObjDE.setGridControl(gcCustDes, gvCustDes, sbSQL).getDataShowOrder(false, false, false, true);
            gvCustDes.Columns["NO"].Visible = false;
            gvCustDes.Columns["CreatedBy"].Visible = false;
            gvCustDes.Columns["CreatedDate"].Visible = false;
            gvCustDes.Columns["UpdatedBy"].Visible = false;
            gvCustDes.Columns["UpdatedDate"].Visible = false;
            gvCustDes.Columns["CUSID"].Visible = false;
            gvCustDes.Columns["ID"].Visible = false;
        }

        private void NewData()
        {
            slueCode.EditValue = "";
            lblStatus.Text = "* Add Destination";
            lblStatus.ForeColor = Color.Green;

            txeID.Text = this.DBC.DBQuery("SELECT CASE WHEN ISNULL(MAX(OIDCUSTAdd), '') = '' THEN 1 ELSE MAX(OIDCUSTAdd) + 1 END AS NewNo FROM CustomerAddress").getString();
            txeCustCode.Text = "";
            txeDes.Text = "";
            txeShip.Text = "";
            txeAddr1.Text = "";
            txeAddr2.Text = "";
            txeAddr3.Text = "";
            txeCountry.Text = "";
            txePostCode.Text = "";
            txeTelNo.Text = "";
            txeFaxNo.Text = "";

            glueCREATE.EditValue = UserLogin.OIDUser;
            txeCDATE.Text = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
            glueUPDATE.EditValue = UserLogin.OIDUser;
            txeUDATE.Text = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");

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

        private void slueCode_EditValueChanged(object sender, EventArgs e)
        {
            if (chkShow == false)
            {
                lblStatus.Text = "* Add Destination";
                lblStatus.ForeColor = Color.Green;

                txeID.Text = this.DBC.DBQuery("SELECT CASE WHEN ISNULL(MAX(OIDCUSTAdd), '') = '' THEN 1 ELSE MAX(OIDCUSTAdd) + 1 END AS NewNo FROM CustomerAddress").getString();
                object xcode = slueCode.Properties.View.GetFocusedRowCellValue("Code");
                if (xcode != null)
                {
                    txeCustCode.Text = xcode.ToString();
                }
                else
                {
                    txeCustCode.Text = "";
                }

                txeDes.Text = "";
                txeShip.Text = "";
                txeAddr1.Text = "";
                txeAddr2.Text = "";
                txeAddr3.Text = "";
                txeCountry.Text = "";
                txePostCode.Text = "";
                txeTelNo.Text = "";
                txeFaxNo.Text = "";

                glueCREATE.EditValue = UserLogin.OIDUser;
                txeCDATE.Text = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
                glueUPDATE.EditValue = UserLogin.OIDUser;
                txeUDATE.Text = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");


                StringBuilder sbSQL = new StringBuilder();
                if (slueCode.Text.Trim() != "")
                {
                    sbSQL.Append("SELECT CUS.Code, CUS.Name, ADDR.Code AS DestinationCode, ADDR.ShipToName, ADDR.ShipToAddress1, ADDR.ShipToAddress2, ADDR.ShipToAddress3, ADDR.Country, ADDR.PostCode, ADDR.TelephoneNo, ADDR.FaxNo, ");
                    sbSQL.Append("       ADDR.CreatedBy, ADDR.CreatedDate, ADDR.UpdatedBy, ADDR.UpdatedDate, ADDR.OIDCUST AS CUSID, ADDR.OIDCUSTAdd AS ID ");
                    sbSQL.Append("FROM   Customer AS CUS INNER JOIN ");
                    sbSQL.Append("       CustomerAddress AS ADDR ON CUS.OIDCUST = ADDR.OIDCUST ");
                    sbSQL.Append("WHERE (ADDR.OIDCUST = '" + slueCode.EditValue.ToString() + "') ");
                    sbSQL.Append("ORDER BY CUS.Code, DestinationCode ");
                }
                else
                {
                    sbSQL.Append("SELECT CUS.Code, CUS.Name, ADDR.Code AS DestinationCode, ADDR.ShipToName, ADDR.ShipToAddress1, ADDR.ShipToAddress2, ADDR.ShipToAddress3, ADDR.Country, ADDR.PostCode, ADDR.TelephoneNo, ADDR.FaxNo, ");
                    sbSQL.Append("       ADDR.CreatedBy, ADDR.CreatedDate, ADDR.UpdatedBy, ADDR.UpdatedDate, ADDR.OIDCUST AS CUSID, ADDR.OIDCUSTAdd AS ID ");
                    sbSQL.Append("FROM   Customer AS CUS INNER JOIN ");
                    sbSQL.Append("       CustomerAddress AS ADDR ON CUS.OIDCUST = ADDR.OIDCUST ");
                    sbSQL.Append("ORDER BY CUS.Code, DestinationCode ");
                }
                new ObjDE.setGridControl(gcCustDes, gvCustDes, sbSQL).getDataShowOrder(false, false, false, true);
                chkShow = false;
                txeDes.Focus();
            }
        }


        private void gvCustDes_RowStyle(object sender, RowStyleEventArgs e)
        {
            
        }

        private void txeDes_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeShip.Focus();
            }
        }

        private void txeDes_LostFocus(object sender, EventArgs e)
        {
            
        }

        private void txeShip_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeAddr1.Focus();
            }
        }

        private void txeAddr1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeAddr2.Focus();
            }
        }

        private void txeAddr2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeAddr3.Focus();
            }
        }

        private void txeAddr3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeCountry.Focus();
            }
        }

        private void txeCountry_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txePostCode.Focus();
            }
        }

        private void txePostCode_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeTelNo.Focus();
            }
        }

        private void txeTelNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeFaxNo.Focus();
            }
        }

        private void txeFaxNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                glueCREATE.Focus();
            }
        }

        private void bbiSave_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (slueCode.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input customer.");
                slueCode.Focus();
            }
            else if (txeDes.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input destination code.");
                txeDes.Focus();
            }
            else if(txeShip.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input ship to name.");
                txeShip.Focus();
            }
            else
            {
                if (FUNC.msgQuiz("Confirm save data ?") == true)
                {
                    StringBuilder sbSQL = new StringBuilder();
                    string strCREATE = UserLogin.OIDUser.ToString() != "" ? UserLogin.OIDUser.ToString() : "0";

                    string strUPDATE = UserLogin.OIDUser.ToString() != "" ? UserLogin.OIDUser.ToString() : "0";

                    bool chkGMP = chkDuplicate();
                    if (chkGMP == true)
                    {
                        if (lblStatus.Text == "* Add Destination")
                        {
                            sbSQL.Append("  INSERT INTO CustomerAddress(OIDCUST, Code, ShipToName, ShipToAddress1, ShipToAddress2, ShipToAddress3, Country, PostCode, TelephoneNo, FaxNo, CreatedBy, CreatedDate, UpdatedBy, UpdatedDate) ");
                            sbSQL.Append("  VALUES('" + slueCode.EditValue.ToString() + "', N'" + txeDes.Text.Trim().Replace("'", "''") + "', N'" + txeShip.Text.Trim().Replace("'", "''") + "', N'" + txeAddr1.Text.Trim().Replace("'", "''") + "', N'" + txeAddr2.Text.Trim().Replace("'", "''") + "', N'" + txeAddr3.Text.Trim().Replace("'", "''") + "', N'" + txeCountry.Text.Trim() + "', N'" + txePostCode.Text.Trim() + "', N'" + txeTelNo.Text.Trim() + "', ");
                            sbSQL.Append("         N'" + txeFaxNo.Text.Trim() + "', '" + strCREATE + "', GETDATE(), '" + strUPDATE + "', GETDATE()) ");
                        }
                        else if (lblStatus.Text == "* Edit Destination")
                        {
                            sbSQL.Append("  UPDATE CustomerAddress SET ");
                            sbSQL.Append("      OIDCUST='" + slueCode.EditValue.ToString() + "', Code=N'" + txeDes.Text.Trim().Replace("'", "''") + "', ");
                            sbSQL.Append("      ShipToName=N'" + txeShip.Text.Trim().Replace("'", "''") + "', ShipToAddress1=N'" + txeAddr1.Text.Trim().Replace("'", "''") + "', ");
                            sbSQL.Append("      ShipToAddress2=N'" + txeAddr2.Text.Trim().Replace("'", "''") + "', ShipToAddress3=N'" + txeAddr3.Text.Trim().Replace("'", "''") + "', ");
                            sbSQL.Append("      Country=N'" + txeCountry.Text.Trim() + "', PostCode=N'" + txePostCode.Text.Trim() + "', TelephoneNo=N'" + txeTelNo.Text.Trim() + "', ");
                            sbSQL.Append("      FaxNo=N'" + txeFaxNo.Text.Trim() + "', UpdatedBy='" + strUPDATE + "', UpdatedDate=GETDATE() ");
                            sbSQL.Append("  WHERE(OIDCUSTAdd = '" + txeID.Text.Trim() + "') ");
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
                else
                {
                    txeDes.Text = "";
                    txeDes.Focus();
                    FUNC.msgWarning("Duplicate destination code. !! Please Change.");
                }
            }
        }


        private bool chkDuplicate()
        {
            bool chkDup = true;
            if (txeDes.Text != "")
            {
                txeDes.Text = txeDes.Text.Trim();

                string strCUSID = "";
                if (slueCode.Text.Trim() != "")
                {
                    strCUSID = slueCode.EditValue.ToString();
                }

                if (txeDes.Text.Trim() != "" && lblStatus.Text == "* Add Destination")
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT TOP(1) Code FROM CustomerAddress WHERE (OIDCUST = '" + strCUSID + "') AND (Code = N'" + txeDes.Text.Trim() + "') ");
                    if (this.DBC.DBQuery(sbSQL).getString() != "")
                    {
                        chkDup = false;
                    }
                }
                else if (txeDes.Text.Trim() != "" && lblStatus.Text == "* Edit Destination")
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT TOP(1) OIDCUSTAdd ");
                    sbSQL.Append("FROM CustomerAddress ");
                    sbSQL.Append("WHERE (OIDCUST = '" + strCUSID + "') AND (Code = N'" + txeDes.Text.Trim() + "') ");
                    string strCHK = this.DBC.DBQuery(sbSQL).getString();
                    if (strCHK != "" && strCHK != txeID.Text.Trim())
                    {
                        chkDup = false;
                    }
                }
            }
            return chkDup;
        }

        private void bbiExcel_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            string pathFile = new ObjSet.Folder(@"C:\MDS\Export\").GetPath() + "CustomerDestinationList_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
            gvCustDes.ExportToXlsx(pathFile);
            System.Diagnostics.Process.Start(pathFile);
        }

        private void gvCustDes_RowClick(object sender, RowClickEventArgs e)
        {
            if (gvCustDes.IsFilterRow(e.RowHandle)) return;
            
        }

        private void gvCustDes_DoubleClick(object sender, EventArgs e)
        {
            DXMouseEventArgs ea = e as DXMouseEventArgs;
            GridView view = sender as GridView;
            GridHitInfo info = view.CalcHitInfo(ea.Location);
            if (info.InRow || info.InRowCell)
            {
                GridView gv = gvCustDes;
                lblStatus.Text = "* Edit Destination";
                lblStatus.ForeColor = Color.Red;
                chkShow = true;
                txeID.Text = gv.GetFocusedRowCellValue("ID").ToString();
                slueCode.EditValue = gv.GetFocusedRowCellValue("CUSID").ToString();
                txeCustCode.Text = gv.GetFocusedRowCellValue("Code").ToString();
                txeDes.Text = gv.GetFocusedRowCellValue("DestinationCode").ToString();
                txeShip.Text = gv.GetFocusedRowCellValue("ShipToName").ToString();
                txeAddr1.Text = gv.GetFocusedRowCellValue("ShipToAddress1").ToString();
                txeAddr2.Text = gv.GetFocusedRowCellValue("ShipToAddress2").ToString();
                txeAddr3.Text = gv.GetFocusedRowCellValue("ShipToAddress3").ToString();
                txeCountry.Text = gv.GetFocusedRowCellValue("Country").ToString();
                txePostCode.Text = gv.GetFocusedRowCellValue("PostCode").ToString();
                txeTelNo.Text = gv.GetFocusedRowCellValue("TelephoneNo").ToString();
                txeFaxNo.Text = gv.GetFocusedRowCellValue("FaxNo").ToString();

                string CreatedBy = gv.GetFocusedRowCellValue("CreatedBy").ToString() == null ? "" : gv.GetFocusedRowCellValue("CreatedBy").ToString();
                glueCREATE.EditValue = CreatedBy;
                txeCDATE.Text = gv.GetFocusedRowCellValue("CreatedDate").ToString();

                string UpdatedBy = gv.GetFocusedRowCellValue("UpdatedBy").ToString() == null ? "" : gv.GetFocusedRowCellValue("UpdatedBy").ToString();
                glueUPDATE.EditValue = UpdatedBy;
                txeUDATE.Text = gv.GetFocusedRowCellValue("UpdatedDate").ToString();
            }
        }

        private void bbiPrintPreview_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            gcCustDes.ShowPrintPreview();
        }

        private void bbiPrint_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            gcCustDes.Print();
        }

        private void txeDes_Leave(object sender, EventArgs e)
        {
            if (txeDes.Text.Trim() != "")
            {
                txeDes.Text = txeDes.Text.ToUpper().Trim();
                bool chkDup = chkDuplicate();
                if (chkDup == false)
                {
                    txeDes.Text = "";
                    txeDes.Focus();
                    FUNC.msgWarning("Duplicate destination code. !! Please Change.");
                }
                else
                {
                    txeShip.Focus();
                }
            }
        }

        private void txeTelNo_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = char.IsLetter(e.KeyChar);
        }

        private void txeFaxNo_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = char.IsLetter(e.KeyChar);
        }

        private void gvCustDes_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator) e.Info.DisplayText = (e.RowHandle + 1).ToString();
            gvCustDes.IndicatorWidth = 40;
        }

      
    }
}