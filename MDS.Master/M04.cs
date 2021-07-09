using System;
using System.Text;
using DBConnect;
using System.Windows.Forms;
using DevExpress.LookAndFeel;
using DevExpress.Utils.Extensions;
using System.Drawing;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraEditors;
using TheepClass;
using DevExpress.Utils;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;

namespace MDS.Master
{
    public partial class M04 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private Functionality.Function FUNC = new Functionality.Function();
        private string selCode = "";
        public LogIn UserLogin { get; set; }

        public string ConnectionString { get; set; }
        string CONNECT_STRING = "";
        DatabaseConnect DBC;

        public M04()
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
            sbSQL.Append("SELECT TOP (1) ReadWriteStatus FROM FunctionAccess WHERE (OIDUser = '" + UserLogin.OIDUser + "') AND(FunctionNo = 'M04') ");
            int chkReadWrite = this.DBC.DBQuery(sbSQL).getInt();
            if (chkReadWrite == 0)
                ribbonPageGroup1.Visible = false;

            sbSQL.Clear();
            sbSQL.Append("SELECT FullName, OIDUSER FROM Users ORDER BY OIDUSER ");
            new ObjDE.setGridLookUpEdit(glueCREATE, sbSQL, "FullName", "OIDUSER").getData();
            new ObjDE.setGridLookUpEdit(glueUPDATE, sbSQL, "FullName", "OIDUSER").getData();

            glueCREATE.EditValue = UserLogin.OIDUser;
            glueUPDATE.EditValue = UserLogin.OIDUser;

            glueCode.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard;
            glueCode.Properties.AcceptEditorTextAsNewValue = DevExpress.Utils.DefaultBoolean.True;

            bbiNew.PerformClick();
        }

        private void NewData()
        {
            txeID.Text = this.DBC.DBQuery("SELECT CASE WHEN ISNULL(MAX(OIDCUST), '') = '' THEN 1 ELSE MAX(OIDCUST) + 1 END AS NewNo FROM Customer").getString();
            lblStatus.Text = "* Add Customer";
            lblStatus.ForeColor = Color.Green;
            glueCode.EditValue = "";
            txeName.Text = "";
            txeShortName.Text = "";
            txeContacts.Text = "";
            txeEmail.Text = "";
            txeAddr1.Text = "";
            txeAddr2.Text = "";
            txeAddr3.Text = "";
            txeCountry.Text = "";
            txePostCode.Text = "";
            txeTelNo.Text = "";
            txeFaxNo.Text = "";
            glueCustType.EditValue = "";
            glueSection.EditValue = "";
            glueTerm.EditValue = "";
            glueCurrency.EditValue = "";
            glueCalendar.EditValue = "";
            txeEval.Text = "";
            txeOthContract.Text = "";
            txeOthAddr1.Text = "";
            txeOthAddr2.Text = "";
            txeOthAddr3.Text = "";
            glueCREATE.EditValue = UserLogin.OIDUser;
            txeCDATE.Text = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
            glueUPDATE.EditValue = UserLogin.OIDUser;
            txeUDATE.Text = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
            selCode = "";
        }

        private void LoadData()
        {
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT Code, Name, ShortName ");
            sbSQL.Append("FROM  Customer ");
            sbSQL.Append("ORDER BY Code, Name ");
            new ObjDE.setGridLookUpEdit(glueCode, sbSQL, "Code", "Code").getData(true);

            //Customer Type
            sbSQL.Clear();
            sbSQL.Append("SELECT '0' AS ID, 'Customer' AS Type ");
            sbSQL.Append("UNION ALL ");
            sbSQL.Append("SELECT '1' AS ID, 'Sub Contract' AS Type ");
            new ObjDE.setGridLookUpEdit(glueCustType, sbSQL, "Type", "ID").getData(true);

            //Sales Section
            sbSQL.Clear();
            sbSQL.Append("SELECT OIDDEPT AS ID, Name AS Department ");
            sbSQL.Append("FROM Departments ");
            sbSQL.Append("ORDER BY OIDDEPT ");
            new ObjDE.setGridLookUpEdit(glueSection, sbSQL, "Department", "Department").getData(true);

            //Payment Term
            sbSQL.Clear();
            sbSQL.Append("SELECT Name, Description ");
            sbSQL.Append("FROM PaymentTerm ");
            sbSQL.Append("ORDER BY OIDPayment ");
            new ObjDE.setGridLookUpEdit(glueTerm, sbSQL, "Name", "Name").getData(true);

            //Payment Currency
            sbSQL.Clear();
            sbSQL.Append("SELECT OIDCURR AS ID, Currency ");
            sbSQL.Append("FROM Currency ");
            sbSQL.Append("ORDER BY OIDCURR ");
            new ObjDE.setGridLookUpEdit(glueCurrency, sbSQL, "Currency", "Currency").getData(true);

            //Calendar No.
            sbSQL.Clear();
            sbSQL.Append("SELECT CompanyName AS CustomerName, Year, OIDCALENDAR AS No  ");
            sbSQL.Append("FROM CalendarMaster A  ");
            sbSQL.Append("CROSS APPLY(SELECT ShortName AS CompanyName FROM Customer WHERE OIDCUST = A.OIDCompany) B  ");
            sbSQL.Append("WHERE CompanyType = 1  ");
            sbSQL.Append("ORDER BY Year DESC, CompanyName, OIDCALENDAR  ");
            new ObjDE.setGridLookUpEdit(glueCalendar, sbSQL, "CustomerName", "No").getData(true);

            //All Customer
            sbSQL.Clear();
            sbSQL.Append("SELECT OIDCUST AS No, Code AS Customer, Name AS CustomerName, ShortName, Contacts AS ContactName, Email, Address1, Address2, Address3, Country, PostCode, TelephoneNo, ");
            sbSQL.Append("       FaxNo, CustomerType, CASE WHEN CustomerType=0 THEN 'Customer' ELSE CASE WHEN CustomerType=1 THEN 'Sub Contract' ELSE '' END END AS [Customer-Type], SalesSection, PaymentTerm, PaymentCurrency, CalendarNo, EvalutionPoint AS CustomerEvalutionPoint, OtherContact AS OtherContactName, ");
            sbSQL.Append("       OtherAddress1, OtherAddress2, OtherAddress3, CreatedBy, CreatedDate, UpdatedBy, UpdatedDate ");
            sbSQL.Append("FROM Customer ");
            sbSQL.Append("ORDER BY CustomerType, Code ");
            new ObjDE.setGridControl(gcCustomer, gvCustomer, sbSQL).getData(false, false, false, true);
            gvCustomer.Columns[0].Visible = false;
            gvCustomer.Columns[23].Visible = false;
            gvCustomer.Columns[24].Visible = false;
            gvCustomer.Columns[25].Visible = false;
            gvCustomer.Columns[26].Visible = false;
            gvCustomer.Columns["CustomerType"].Visible = false;
        }

        private void bbiNew_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            NewData();
            LoadData();
        }

        private void bbiSave_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (glueCode.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input customer code.");
                glueCode.Focus();
            }
            else if (txeName.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input customer name.");
                txeName.Focus();
            }
            else if (txeShortName.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input short name.");
                txeShortName.Focus();
            }
            else if (glueCustType.Text.Trim() == "")
            {
                FUNC.msgWarning("Please select customer type.");
                glueCustType.Focus();
            }
            //else if (glueCalendar.Text.Trim() == "")
            //{
            //    FUNC.msgWarning("Please select calendar no.");
            //    glueCalendar.Focus();
            //}
            else
            {
                if (FUNC.msgQuiz("Confirm save data ?") == true)
                {
                    StringBuilder sbSQL = new StringBuilder();

                    string strCREATE = UserLogin.OIDUser.ToString() != "" ? UserLogin.OIDUser.ToString() : "0";
                    string strUPDATE = UserLogin.OIDUser.ToString() != "" ? UserLogin.OIDUser.ToString() : "0";

                    string strCalendar = glueCalendar.Text.Trim() == "" ? "NULL" : "'" + glueCalendar.EditValue.ToString() + "'";

                    if (lblStatus.Text == "* Add Customer")
                    {
                        sbSQL.Append("  INSERT INTO Customer(Code, Name, ShortName, Contacts, Email, Address1, Address2, Address3, Country, PostCode, TelephoneNo, FaxNo, CustomerType, SalesSection, PaymentTerm, PaymentCurrency, CalendarNo, EvalutionPoint, OtherContact, OtherAddress1, OtherAddress2, OtherAddress3, CreatedBy, CreatedDate, UpdatedBy, UpdatedDate) ");
                        sbSQL.Append("  VALUES(N'" + glueCode.Text.Trim().Replace("'", "''") + "', N'" + txeName.Text.Trim().Replace("'", "''") + "', N'" + txeShortName.Text.Trim().Replace("'", "''") + "', N'" + txeContacts.Text.Trim().Replace("'", "''") + "', N'" + txeEmail.Text.Trim() + "', N'" + txeAddr1.Text.Trim() + "', N'" + txeAddr2.Text.Trim() + "', N'" + txeAddr3.Text.Trim() + "', N'" + txeCountry.Text.Trim() + "', N'" + txePostCode.Text.Trim() + "', N'" + txeTelNo.Text.Trim() + "', ");
                        sbSQL.Append("         N'" + txeFaxNo.Text.Trim() + "', '" + glueCustType.EditValue.ToString() + "', N'" + glueSection.Text.Trim() + "', N'" + glueTerm.Text.Trim() + "', N'" + glueCurrency.Text.Trim() + "', " + strCalendar + ", N'" + txeEval.Text.Trim() + "', N'" + txeOthContract.Text.Trim().Replace("'", "''") + "', N'" + txeOthAddr1.Text.Trim() + "', N'" + txeOthAddr2.Text.Trim() + "', N'" + txeOthAddr3.Text.Trim() + "', '" + strCREATE + "', GETDATE(), '" + strUPDATE + "', GETDATE()) ");
                    }
                    else if (lblStatus.Text == "* Edit Customer")
                    {
                        sbSQL.Append("  UPDATE Customer SET ");
                        sbSQL.Append("      Code = N'" + glueCode.Text.Trim().Replace("'", "''") + "', Name = N'" + txeName.Text.Trim().Replace("'", "''") + "', ShortName = N'" + txeShortName.Text.Trim().Replace("'", "''") + "', Contacts = N'" + txeContacts.Text.Trim().Replace("'", "''") + "', Email = N'" + txeEmail.Text.Trim() + "', Address1 = N'" + txeAddr1.Text.Trim() + "', Address2 = N'" + txeAddr2.Text.Trim() + "', Address3 = N'" + txeAddr3.Text.Trim() + "', ");
                        sbSQL.Append("      Country = N'" + txeCountry.Text.Trim() + "', PostCode = N'" + txePostCode.Text.Trim() + "', TelephoneNo = N'" + txeTelNo.Text.Trim() + "', FaxNo = N'" + txeFaxNo.Text.Trim() + "', CustomerType = '" + glueCustType.EditValue.ToString() + "', SalesSection = N'" + glueSection.Text.Trim() + "', PaymentTerm = N'" + glueTerm.Text.Trim() + "', ");
                        sbSQL.Append("      PaymentCurrency = N'" + glueCurrency.Text.Trim() + "', CalendarNo = " + strCalendar + ", EvalutionPoint = N'" + txeEval.Text.Trim() + "', OtherContact = N'" + txeOthContract.Text.Trim().Replace("'", "''") + "', OtherAddress1 = N'" + txeOthAddr1.Text.Trim() + "', OtherAddress2 = N'" + txeOthAddr2.Text.Trim() + "', OtherAddress3 = N'" + txeOthAddr3.Text.Trim() + "', ");
                        sbSQL.Append("      UpdatedBy = '" + strUPDATE + "', UpdatedDate = GETDATE() ");
                        sbSQL.Append("  WHERE(OIDCUST = '" + txeID.Text.Trim() + "') ");
                    }

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

        private void glueCode_EditValueChanged(object sender, EventArgs e)
        {
            //txeName.Focus();
            //LoadCode(glueCode.Text);
        }


        private void glueCode_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) 
            {
                txeName.Focus();
            }
        }

        private void glueCode_LostFocus(object sender, EventArgs e)
        {
            
            
        }

        private void LoadCode(string strCODE)
        {
            strCODE = strCODE.ToUpper().Trim();
            

            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT OIDCUST, Code, Name, ShortName, Contacts, Email, Address1, Address2, Address3, Country, PostCode, TelephoneNo, FaxNo, CustomerType, SalesSection, PaymentTerm, PaymentCurrency, CalendarNo, ");
            sbSQL.Append("       EvalutionPoint, OtherContact, OtherAddress1, OtherAddress2, OtherAddress3, CreatedBy, CreatedDate, UpdatedBy, UpdatedDate ");
            sbSQL.Append("FROM   Customer ");
            sbSQL.Append("WHERE (Code = N'" + strCODE.Replace("'", "''") + "') ");
            string[] arrCust = this.DBC.DBQuery(sbSQL).getMultipleValue();
            if (arrCust.Length > 0)
            {
                if (FUNC.msgQuiz("The system already has this name. Want to fix it ?") == true)
                {
                    txeID.Text = arrCust[0];
                    lblStatus.Text = "* Edit Customer";
                    lblStatus.ForeColor = Color.Red;
                    txeName.Text = arrCust[2];
                    txeShortName.Text = arrCust[3];
                    txeContacts.Text = arrCust[4];
                    txeEmail.Text = arrCust[5];
                    txeAddr1.Text = arrCust[6];
                    txeAddr2.Text = arrCust[7];
                    txeAddr3.Text = arrCust[8];
                    txeCountry.Text = arrCust[9];
                    txePostCode.Text = arrCust[10];
                    txeTelNo.Text = arrCust[11];
                    txeFaxNo.Text = arrCust[12];
                    glueCustType.EditValue = arrCust[13];
                    glueSection.EditValue = arrCust[14];
                    glueTerm.EditValue = arrCust[15];
                    glueCurrency.EditValue = arrCust[16];
                    glueCalendar.EditValue = arrCust[17];
                    txeEval.Text = arrCust[18]; ;
                    txeOthContract.Text = arrCust[19];
                    txeOthAddr1.Text = arrCust[20];
                    txeOthAddr2.Text = arrCust[21];
                    txeOthAddr3.Text = arrCust[22];
                    glueCREATE.EditValue = arrCust[23];
                    txeCDATE.Text = arrCust[24];
                    glueUPDATE.EditValue = arrCust[25];
                    txeUDATE.Text = arrCust[26];
                    txeName.Focus();
                }
                else
                {
                    txeID.Text = this.DBC.DBQuery("SELECT CASE WHEN ISNULL(MAX(OIDCUST), '') = '' THEN 1 ELSE MAX(OIDCUST) + 1 END AS NewNo FROM Customer").getString();
                    glueCode.Text = "";
                    glueCalendar.EditValue = "";
                    glueCode.Focus();
                    lblStatus.Text = "* Add Customer";
                    lblStatus.ForeColor = Color.Green;

                }
            }
            else
            {
                txeID.Text = this.DBC.DBQuery("SELECT CASE WHEN ISNULL(MAX(OIDCUST), '') = '' THEN 1 ELSE MAX(OIDCUST) + 1 END AS NewNo FROM Customer").getString();
                lblStatus.Text = "* Add Customer";
                lblStatus.ForeColor = Color.Green;
                glueCalendar.EditValue = "";

                bool chkNameDup = chkDuplicateName();
                if (chkNameDup == false)
                {
                    txeName.Text = "";
                }

                bool chkShortDup = chkDuplicateShortName();
                if (chkShortDup == false)
                {
                    txeShortName.Text = "";
                }
                txeName.Focus();
            }

            selCode = "";

            //Check new customer or edit customer
            //sbSQL.Clear();
            //sbSQL.Append("SELECT OIDCUST FROM Customer WHERE (OIDCUST = '" + txeID.Text.ToString() + "') ");
            //string strCHKID = this.DBC.DBQuery(sbSQL).getString();
            //if (strCHKID == "")
            //{
            //    lblStatus.Text = "* Add Customer";
            //    lblStatus.ForeColor = Color.Green;
            //}
            //else
            //{
            //    lblStatus.Text = "* Edit Customer";
            //    lblStatus.ForeColor = Color.Red;
            //}
            
        }

        private void gvCustomer_RowCellClick(object sender, RowCellClickEventArgs e)
        {
            
        }

        private void bbiExcel_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            string pathFile = new ObjSet.Folder(@"C:\MDS\Export\").GetPath() + "CustomerList_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
            gvCustomer.ExportToXlsx(pathFile);
            System.Diagnostics.Process.Start(pathFile);
        }

        private void txeName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeShortName.Focus();
            }
        }

        private void txeShortName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeContacts.Focus();
            }
        }

        private void txeContacts_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeEmail.Focus();
            }
        }

        private void txeEmail_KeyDown(object sender, KeyEventArgs e)
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
            if(e.KeyCode==Keys.Enter)
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
                glueCustType.Focus();
            }
        }

        private void glueCustType_EditValueChanged(object sender, EventArgs e)
        {
            glueSection.Focus();
        }

        private void glueSection_EditValueChanged(object sender, EventArgs e)
        {
            glueTerm.Focus();
        }

        private void glueTerm_EditValueChanged(object sender, EventArgs e)
        {
            glueCurrency.Focus();
        }

        private void glueCurrency_EditValueChanged(object sender, EventArgs e)
        {
            glueCalendar.Focus();
        }

        private void glueCalendar_EditValueChanged(object sender, EventArgs e)
        {
            txeEval.Focus();
        }

        private void txeEval_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeOthContract.Focus();
            }
        }

        private void txeOthContract_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeOthAddr1.Focus();
            }
        }

        private void txeOthAddr1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeOthAddr2.Focus();
            }
        }

        private void txeOthAddr2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeOthAddr3.Focus();
            }
        }

        private void gvCustomer_RowStyle(object sender, RowStyleEventArgs e)
        {
           
        }

        private void glueCode_ProcessNewValue(object sender, DevExpress.XtraEditors.Controls.ProcessNewValueEventArgs e)
        {
            GridLookUpEdit gridLookup = sender as GridLookUpEdit;
            if (e.DisplayValue == null) return;
            string newValue = e.DisplayValue.ToString();
            if (newValue == String.Empty) return;
        }

        private void glueCode_CloseUp(object sender, DevExpress.XtraEditors.Controls.CloseUpEventArgs e)
        {
            
        }

        private void glueCode_Closed(object sender, DevExpress.XtraEditors.Controls.ClosedEventArgs e)
        {
            glueCode.Focus();
            txeName.Focus();
        }

        private void gvCustomer_RowClick(object sender, RowClickEventArgs e)
        {
            if (gvCustomer.IsFilterRow(e.RowHandle)) return;
            //txeID.Text = gvCustomer.GetFocusedRowCellValue("No").ToString();
            //lblStatus.Text = "* Edit Customer";
            //lblStatus.ForeColor = Color.Red;
            //glueCode.EditValue = gvCustomer.GetFocusedRowCellValue("Customer").ToString();
            //txeName.Text = gvCustomer.GetFocusedRowCellValue("CustomerName").ToString();
            //txeShortName.Text = gvCustomer.GetFocusedRowCellValue("ShortName").ToString();
            //txeContacts.Text = gvCustomer.GetFocusedRowCellValue("ContactName").ToString();
            //txeEmail.Text = gvCustomer.GetFocusedRowCellValue("Email").ToString();
            //txeAddr1.Text = gvCustomer.GetFocusedRowCellValue("Address1").ToString();
            //txeAddr2.Text = gvCustomer.GetFocusedRowCellValue("Address2").ToString();
            //txeAddr3.Text = gvCustomer.GetFocusedRowCellValue("Address3").ToString();
            //txeCountry.Text = gvCustomer.GetFocusedRowCellValue("Country").ToString();
            //txePostCode.Text = gvCustomer.GetFocusedRowCellValue("PostCode").ToString();
            //txeTelNo.Text = gvCustomer.GetFocusedRowCellValue("TelephoneNo").ToString();
            //txeFaxNo.Text = gvCustomer.GetFocusedRowCellValue("FaxNo").ToString();
            //glueCustType.EditValue = gvCustomer.GetFocusedRowCellValue("CustomerType").ToString();
            //glueSection.EditValue = gvCustomer.GetFocusedRowCellValue("SalesSection").ToString();
            //glueTerm.EditValue = gvCustomer.GetFocusedRowCellValue("PaymentTerm").ToString();
            //glueCurrency.EditValue = gvCustomer.GetFocusedRowCellValue("PaymentCurrency").ToString();
            //glueCalendar.EditValue = gvCustomer.GetFocusedRowCellValue("CalendarNo").ToString();
            //txeEval.Text = gvCustomer.GetFocusedRowCellValue("CustomerEvalutionPoint").ToString();
            //txeOthContract.Text = gvCustomer.GetFocusedRowCellValue("OtherContactName").ToString();
            //txeOthAddr1.Text = gvCustomer.GetFocusedRowCellValue("OtherAddress1").ToString();
            //txeOthAddr2.Text = gvCustomer.GetFocusedRowCellValue("OtherAddress2").ToString();
            //txeOthAddr3.Text = gvCustomer.GetFocusedRowCellValue("OtherAddress3").ToString();

            //glueCREATE.EditValue = gvCustomer.GetFocusedRowCellValue("CreatedBy").ToString();
            //txeCDATE.Text = gvCustomer.GetFocusedRowCellValue("CreatedDate").ToString();
            //glueUPDATE.EditValue = gvCustomer.GetFocusedRowCellValue("UpdatedBy").ToString();
            //txeUDATE.Text = gvCustomer.GetFocusedRowCellValue("UpdatedDate").ToString();
        }

        private void gvCustomer_DoubleClick(object sender, EventArgs e)
        {
            DXMouseEventArgs ea = e as DXMouseEventArgs;
            GridView view = sender as GridView;
            GridHitInfo info = view.CalcHitInfo(ea.Location);
            if (info.InRow || info.InRowCell)
            {
                GridView gv = gvCustomer;
                txeID.Text = gv.GetFocusedRowCellValue("No").ToString();
                lblStatus.Text = "* Edit Customer";
                lblStatus.ForeColor = Color.Red;

                glueCode.EditValue = gv.GetFocusedRowCellValue("Customer").ToString();
                txeName.Text = gv.GetFocusedRowCellValue("CustomerName").ToString();
                txeShortName.Text = gv.GetFocusedRowCellValue("ShortName").ToString();
                txeContacts.Text = gv.GetFocusedRowCellValue("ContactName").ToString();
                txeEmail.Text = gv.GetFocusedRowCellValue("Email").ToString();
                txeAddr1.Text = gv.GetFocusedRowCellValue("Address1").ToString();
                txeAddr2.Text = gv.GetFocusedRowCellValue("Address2").ToString();
                txeAddr3.Text = gv.GetFocusedRowCellValue("Address3").ToString();
                txeCountry.Text = gv.GetFocusedRowCellValue("Country").ToString();
                txePostCode.Text = gv.GetFocusedRowCellValue("PostCode").ToString();
                txeTelNo.Text = gv.GetFocusedRowCellValue("TelephoneNo").ToString();
                txeFaxNo.Text = gv.GetFocusedRowCellValue("FaxNo").ToString();
                glueCustType.EditValue = gv.GetFocusedRowCellValue("CustomerType").ToString();
                glueSection.EditValue = gv.GetFocusedRowCellValue("SalesSection").ToString();
                glueTerm.EditValue = gv.GetFocusedRowCellValue("PaymentTerm").ToString();
                glueCurrency.EditValue = gv.GetFocusedRowCellValue("PaymentCurrency").ToString();
                glueCalendar.EditValue = gv.GetFocusedRowCellValue("CalendarNo").ToString();
                txeEval.Text = gv.GetFocusedRowCellValue("CustomerEvalutionPoint").ToString();
                txeOthContract.Text = gv.GetFocusedRowCellValue("OtherContactName").ToString();
                txeOthAddr1.Text = gv.GetFocusedRowCellValue("OtherAddress1").ToString();
                txeOthAddr2.Text = gv.GetFocusedRowCellValue("OtherAddress2").ToString();
                txeOthAddr3.Text = gv.GetFocusedRowCellValue("OtherAddress3").ToString();

                glueCREATE.EditValue = gv.GetFocusedRowCellValue("CreatedBy").ToString();
                txeCDATE.Text = gv.GetFocusedRowCellValue("CreatedDate").ToString();
                glueUPDATE.EditValue = gv.GetFocusedRowCellValue("UpdatedBy").ToString();
                txeUDATE.Text = gv.GetFocusedRowCellValue("UpdatedDate").ToString();
            }

        }

        private void bbiPrintPreview_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            gcCustomer.ShowPrintPreview();
        }

        private void bbiPrint_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            gcCustomer.Print();
        }

        private void glueCode_Leave(object sender, EventArgs e)
        {
            if (glueCode.Text.Trim() != "" && glueCode.Text.ToUpper().Trim() != selCode)
            {
                glueCode.Text = glueCode.Text.ToUpper().Trim();
                selCode = glueCode.Text;
                LoadCode(glueCode.Text);
                //MessageBox.Show(glueCode.Text);
            }
        }

        private void txeName_Leave(object sender, EventArgs e)
        {
            if (txeName.Text.Trim() != "")
            {
                txeName.Text = txeName.Text.Trim();
                bool chkDup = chkDuplicateName();
                if (chkDup == true)
                {
                    txeShortName.Focus();
                }
                else
                {
                    txeName.Text = "";
                    txeName.Focus();
                    FUNC.msgWarning("Duplicate customer name. !! Please Change.");

                }
            }
        }

        private void txeShortName_Leave(object sender, EventArgs e)
        {
            if (txeShortName.Text.Trim() != "")
            {
                txeShortName.Text = txeShortName.Text.ToUpper().Trim();
                bool chkDup = chkDuplicateShortName();
                if (chkDup == true)
                {
                    txeContacts.Focus();
                }
                else
                {
                    txeShortName.Text = "";
                    txeShortName.Focus();
                    FUNC.msgWarning("Duplicate customer short name. !! Please Change.");

                }
            }
        }

        private bool chkDuplicateName()
        {
            bool chkDup = true;
            if (txeName.Text != "")
            {
                txeName.Text = txeName.Text.Trim();

                if (txeName.Text.Trim() != "" && lblStatus.Text == "* Add Customer")
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT TOP(1) Name FROM Customer WHERE (Name = N'" + txeName.Text.Trim().Replace("'", "''") + "') ");
                    if (this.DBC.DBQuery(sbSQL).getString() != "")
                    {
                        chkDup = false;
                    }
                }
                else if (txeName.Text.Trim() != "" && lblStatus.Text == "* Edit Customer")
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT TOP(1) OIDCUST ");
                    sbSQL.Append("FROM Customer ");
                    sbSQL.Append("WHERE (Name = N'" + txeName.Text.Trim().Replace("'", "''") + "') ");
                    string strCHK = this.DBC.DBQuery(sbSQL).getString();
                    if (strCHK != "" && strCHK != txeID.Text.Trim())
                    {
                        chkDup = false;
                    }
                }
            }
            return chkDup;
        }

        private bool chkDuplicateShortName()
        {
            bool chkDup = true;
            if (txeShortName.Text != "")
            {
                txeShortName.Text = txeShortName.Text.ToUpper().Trim();

                if (txeShortName.Text.Trim() != "" && lblStatus.Text == "* Add Customer")
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT TOP(1) ShortName FROM Customer WHERE (ShortName = N'" + txeShortName.Text.Trim().Replace("'", "''") + "') ");
                    if (this.DBC.DBQuery(sbSQL).getString() != "")
                    {
                        chkDup = false;
                    }
                }
                else if (txeShortName.Text.Trim() != "" && lblStatus.Text == "* Edit Customer")
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT TOP(1) OIDCUST ");
                    sbSQL.Append("FROM Customer ");
                    sbSQL.Append("WHERE (ShortName = N'" + txeShortName.Text.Trim().Replace("'", "''") + "') ");
                    string strCHK = this.DBC.DBQuery(sbSQL).getString();
                    if (strCHK != "" && strCHK != txeID.Text.Trim())
                    {
                        chkDup = false;
                    }
                }
            }
            return chkDup;
        }

        private void txeTelNo_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = char.IsLetter(e.KeyChar);
        }

        private void txeFaxNo_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = char.IsLetter(e.KeyChar);
        }

        private void gvCustomer_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator) e.Info.DisplayText = (e.RowHandle + 1).ToString();
            gvCustomer.IndicatorWidth = 40;
        }

        private void ribbonControl_Click(object sender, EventArgs e)
        {

        }
    }
}