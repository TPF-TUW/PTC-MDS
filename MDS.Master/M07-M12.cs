using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data.SqlClient;
using DBConnect;

namespace MDS.Master
{
    public partial class M07_M12 : DevExpress.XtraEditors.XtraForm
    {
        private Functionality.Function FUNCT = new Functionality.Function();
        //Global Variable
        //classConn db = new classConn();
        //classTools ct = new classTools();
        //SqlConnection mainConn = new classConn().MDS();
        //SqlConnection conn;
        string sql = "";

        string _Type = "";

        int _UserID = 0;
        string CusCode = "";
        DatabaseConnect DB = new DatabaseConnect();
        public M07_M12(DatabaseConnect DBase, string Type, int UserID)
        {
            InitializeComponent();
            this.DB = DBase;
            this._Type = Type;
            this._UserID = UserID;
        }

        private void chkNull(string Alert, TextEdit txtName)
        {
            FUNCT.msgWarning("Please Key : "+Alert+"!"); txtName.Focus(); return;
        }

        private void btnAddCustomer_Click(object sender, EventArgs e)
        {
            CusCode = txeCode.Text.ToString().ToUpper().Trim().Replace("'","''");
            string CusName = txeName.Text.ToString().Trim().Replace("'", "''");
            string CusType = cbeType.EditValue.ToString();

            string Address1 = txeAddr1.Text.ToString().Trim().Replace("'", "''");
            string Address2 = txeAddr2.Text.ToString().Trim().Replace("'", "''");
            string Address3 = txeAddr3.Text.ToString().Trim().Replace("'", "''");
            string Country = txeCountry.Text.ToString().Trim().Replace("'", "''");
            string TelephoneNo = txeTel.Text.ToString().Trim().Replace("'", "''");
            string Email = txeEmail.Text.ToString().Trim().Replace("'", "''");

            string strCREATE = this._UserID.ToString() != "" ? this._UserID.ToString() : "0";

            if (CusCode == "") { chkNull("Vendor Code", txeCode); }
            else if (CusName == "") { chkNull("Vendor Name", txeName); }
            else if (cbeType.Text.Trim() == "") { chkNull("Vendor Type", cbeType); }
            else
            {
                //chkDup
                if (DB.DBQuery("SELECT TOP(1) OIDVEND FROM Vendor WHERE (Code = N'" + CusCode + "') ").getString() != "")
                {
                    FUNCT.msgWarning("Vendor Code is Duplicate!"); txeCode.Focus(); return;
                }
                else
                {
                    if (FUNCT.msgQuiz("SAVE Supplier (Vendor) ?") == true)
                    {
                        sql = "INSERT INTO Vendor(Code, Name, VendorType, Address1, Address2, Address3, Country, TelephoneNo, Email, CreatedBy, CreatedDate, UpdatedBy, UpdatedDate) VALUES(N'" + CusCode + "', N'" + CusName + "', '" + CusType + "', N'" + Address1 + "', N'" + Address2 + "', N'" + Address3 + "', N'" + Country + "', N'" + TelephoneNo + "', N'" + Email + "', '" + strCREATE + "', GETDATE(), '" + strCREATE + "', GETDATE()) ";
                        //Console.WriteLine(sql);
                        bool chkSave = DB.DBQuery(sql).runSQL();
                        if (chkSave == true)
                        {
                            FUNCT.msgInfo("Save Supplier (Vendor) is Successfull.");
                            this.Close();
                        }
                    }
                }
            }
        }

        private void M07_M12_Load(object sender, EventArgs e)
        {
            new ObjDE.setDatabase(this.DB);
            //MessageBox.Show(DB.getCONNECTION_STRING());
            //MessageBox.Show(ObjDE.GlobalVar.DBC.getCONNECTION_STRING());

            StringBuilder sbTYPE = new StringBuilder();
            sbTYPE.Append("SELECT Name AS VendorType, No AS ID FROM ENUMTYPE WHERE (Module = N'Vendor') ORDER BY No ");
            new ObjDE.setGridLookUpEdit(cbeType, sbTYPE, "VendorType", "ID").getData();
            cbeType.Properties.View.PopulateColumns(cbeType.Properties.DataSource);
            cbeType.Properties.View.Columns["ID"].Visible = false;
            cbeType.EditValue = 0;
          
            txeCode.Focus();
        }

        private void M07_M12_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (Application.OpenForms.OfType<M07>().Count() > 0)
            {
                var frmD01 = Application.OpenForms.OfType<M07>().FirstOrDefault();
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT VD.Code, VD.Name, ENT.Name AS Type, VD.OIDVEND AS ID ");
                sbSQL.Append("FROM   Vendor AS VD INNER JOIN ");
                sbSQL.Append("       ENUMTYPE AS ENT ON VD.VendorType = ENT.No AND ENT.Module = N'Vendor' ");
                sbSQL.Append("ORDER BY VD.Name ");
                new ObjDE.setSearchLookUpEdit(frmD01.slueDefaultVendor, sbSQL, "Name", "ID").getData();

                frmD01.slueFirstVendor.Properties.DataSource = frmD01.slueDefaultVendor.Properties.DataSource;
                frmD01.slueFirstVendor.Properties.DisplayMember = frmD01.slueDefaultVendor.Properties.DisplayMember;
                frmD01.slueFirstVendor.Properties.ValueMember = frmD01.slueDefaultVendor.Properties.ValueMember;
                //frmD01.slueFirstVendor.Properties.View.Columns["ID"].Visible = false;

                frmD01.slueVendorCode.Properties.DataSource = frmD01.slueDefaultVendor.Properties.DataSource;
                frmD01.slueVendorCode.Properties.DisplayMember = frmD01.slueDefaultVendor.Properties.DisplayMember;
                frmD01.slueVendorCode.Properties.ValueMember = frmD01.slueDefaultVendor.Properties.ValueMember;
                //frmD01.slueVendorCode.Properties.View.Columns["ID"].Visible = false;

                frmD01.slueSVendor.Properties.DataSource = frmD01.slueDefaultVendor.Properties.DataSource;
                frmD01.slueSVendor.Properties.DisplayMember = frmD01.slueDefaultVendor.Properties.DisplayMember;
                frmD01.slueSVendor.Properties.ValueMember = frmD01.slueDefaultVendor.Properties.ValueMember;
                //frmD01.slueSVendor.Properties.View.Columns["ID"].Visible = false;
            }
        }
    }
}