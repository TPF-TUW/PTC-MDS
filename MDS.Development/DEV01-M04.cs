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

namespace MDS.Development
{
    public partial class DEV01_M04 : DevExpress.XtraEditors.XtraForm
    {
        private Functionality.Function FUNCT = new Functionality.Function();
        //Global Variable
        //classConn db = new classConn();
        //classTools ct = new classTools();
        //SqlConnection mainConn = new classConn().MDS();
        //SqlConnection conn;
        string sql = "";

        int _UserID = 0;
        string CustomerCode = "";
        DatabaseConnect DB = new DatabaseConnect();
        public DEV01_M04(DatabaseConnect DBase, int UserID)
        {
            InitializeComponent();
            this.DB = DBase;
            this._UserID = UserID;
        }

        private void chkNull(string Alert, TextEdit txtName)
        {
            FUNCT.msgWarning("Please Key : " +Alert+"!"); txtName.Focus(); return;
        }

        private void btnAddCustomer_Click(object sender, EventArgs e)
        {
            string CustomerName = txtCustomerName.Text.ToString().Trim().Replace("'","''");
            string CustomerShortName = txtCustomerShortName.Text.ToString().Trim().Replace("'", "''");
            CustomerCode = txtCustomerCode.Text.ToString().Trim().Replace("'", "''");

            string strCREATE = this._UserID.ToString() != "" ? this._UserID.ToString() : "0";

            if (CustomerName == "") { chkNull("CustomerName", txtCustomerName); }
            else if (CustomerShortName == "") { chkNull("CustomerShortName", txtCustomerShortName); }
            else if (CustomerCode == "") { chkNull("CustomerCode", txtCustomerCode); }
            else
            {
                //chkDup
                if (DB.DBQuery("SELECT TOP (1) ShortName From Customer WHERE ShortName = '" + CustomerShortName + "' ").getString() != "")
                {
                    FUNCT.msgWarning("CustomerShortName is Duplicate!"); txtCustomerShortName.Focus(); return;
                }
                else if (DB.DBQuery("SELECT TOP(1) Code FROM Customer WHERE Code = '" + CustomerCode + "' ").getString() != "")
                {
                    FUNCT.msgWarning("CustomerCode is Duplicate!"); txtCustomerCode.Focus(); return;
                }
                else
                {
                    if (FUNCT.msgQuiz("SAVE Customer?") == true)
                    {
                        sql = "INSERT INTO Customer(Name, ShortName, Code, CustomerType, CreatedBy, CreatedDate, UpdatedBy, UpdatedDate) VALUES (N'" + CustomerName + "', N'" + CustomerShortName + "', N'" + CustomerCode + "', '9', '" + strCREATE + "', GETDATE(), '" + strCREATE + "', GETDATE())";
                        bool chkSave = DB.DBQuery(sql).runSQL();
                        if (chkSave == true)
                        {
                            FUNCT.msgInfo("Save Customer is Successfull.");
                            this.Close();
                        }
                    }
                }
            }
        }

        private void DEV01_M04_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (Application.OpenForms.OfType<DEV01>().Count() > 0)
            {
                var frmD01 = Application.OpenForms.OfType<DEV01>().FirstOrDefault();
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT ShortName, Name, OIDCUST AS ID FROM Customer ORDER BY ShortName");
                new ObjDE.setSearchLookUpEdit(frmD01.slCustomer_Main, sbSQL, "Name", "ID").getData();
                if(CustomerCode != "")
                    frmD01.slCustomer_Main.EditValue = DB.DBQuery("SELECT TOP(1) OIDCUST FROM Customer WHERE Code=N'" + CustomerCode + "'").getString();
                frmD01.slCustomer_Main.Properties.View.PopulateColumns(frmD01.slCustomer_Main.Properties.DataSource);
                frmD01.slCustomer_Main.Properties.View.Columns["ID"].Visible = false;
            }
        }

        private void DEV01_M04_Load(object sender, EventArgs e)
        {
            new ObjDE.setDatabase(this.DB);
            //MessageBox.Show(DB.getCONNECTION_STRING());
            //MessageBox.Show(ObjDE.GlobalVar.DBC.getCONNECTION_STRING());
        }
    }
}