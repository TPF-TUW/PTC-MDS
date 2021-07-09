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
    public partial class DEV01_M12 : DevExpress.XtraEditors.XtraForm
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
        public DEV01_M12(DatabaseConnect DBase, string Type, int UserID)
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
                        sql = "INSERT INTO Vendor(Code, Name, VendorType, CreatedBy, CreatedDate, UpdatedBy, UpdatedDate) VALUES(N'" + CusCode + "', N'" + CusName + "', '" + CusType + "', '" + strCREATE + "', GETDATE(), '" + strCREATE + "', GETDATE()) ";
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

        private void DEV01_M12_Load(object sender, EventArgs e)
        {
            new ObjDE.setDatabase(this.DB);
            //MessageBox.Show(DB.getCONNECTION_STRING());
            //MessageBox.Show(ObjDE.GlobalVar.DBC.getCONNECTION_STRING());

            StringBuilder sbTYPE = new StringBuilder();
            sbTYPE.Append("SELECT Name AS VendorType, No AS ID FROM ENUMTYPE WHERE (Module = N'Vendor') ORDER BY No ");
            new ObjDE.setGridLookUpEdit(cbeType, sbTYPE, "VendorType", "ID").getData();
            cbeType.Properties.View.PopulateColumns(cbeType.Properties.DataSource);
            cbeType.Properties.View.Columns["ID"].Visible = false;

            if (this._Type == "FG") //Finished Goods
            {
                cbeType.EditValue = 0;
            }
            else if (this._Type == "FB") //Fabric
            {
                cbeType.EditValue = 1;
            }
            else if (this._Type == "MT") //Material
            {
                cbeType.EditValue = 2;
            }
            txeCode.Focus();
        }

        private void DEV01_M12_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (Application.OpenForms.OfType<DEV01>().Count() > 0)
            {
                var frmD01 = Application.OpenForms.OfType<DEV01>().FirstOrDefault();
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT Code AS VendorCode, Name AS VendorName, OIDVEND AS ID FROM Vendor ORDER BY VendorType");
                if (cbeType.EditValue.ToString() == "1") //Fabric
                {
                    new ObjDE.setSearchLookUpEdit(frmD01.slVendor_FB, sbSQL, "VendorName", "ID").getData();
                    if(CusCode != "")
                        frmD01.slVendor_FB.EditValue = DB.DBQuery("SELECT TOP(1) OIDVEND FROM Vendor WHERE Code=N'" + CusCode + "'").getString();
                    frmD01.slVendor_FB.Properties.View.PopulateColumns(frmD01.slVendor_FB.Properties.DataSource);
                    frmD01.slVendor_FB.Properties.View.Columns["ID"].Visible = false;

                    if (frmD01.slVendor_Mat.Text.Trim() == "")
                    {
                        new ObjDE.setSearchLookUpEdit(frmD01.slVendor_Mat, sbSQL, "VendorName", "ID").getData();
                        if (CusCode != "")
                            frmD01.slVendor_Mat.EditValue = DB.DBQuery("SELECT TOP(1) OIDVEND FROM Vendor WHERE Code=N'" + CusCode + "'").getString();
                        frmD01.slVendor_Mat.Properties.View.PopulateColumns(frmD01.slVendor_Mat.Properties.DataSource);
                        frmD01.slVendor_Mat.Properties.View.Columns["ID"].Visible = false;
                    }
                }
                else if (cbeType.EditValue.ToString() == "2" || cbeType.EditValue.ToString() == "3") //Material
                {
                    new ObjDE.setSearchLookUpEdit(frmD01.slVendor_Mat, sbSQL, "VendorName", "ID").getData();
                    if (CusCode != "")
                        frmD01.slVendor_Mat.EditValue = DB.DBQuery("SELECT TOP(1) OIDVEND FROM Vendor WHERE Code=N'" + CusCode + "'").getString();
                    frmD01.slVendor_Mat.Properties.View.PopulateColumns(frmD01.slVendor_Mat.Properties.DataSource);
                    frmD01.slVendor_Mat.Properties.View.Columns["ID"].Visible = false;

                    if (frmD01.slVendor_FB.Text.Trim() == "")
                    {
                        new ObjDE.setSearchLookUpEdit(frmD01.slVendor_FB, sbSQL, "VendorName", "ID").getData();
                        if (CusCode != "")
                            frmD01.slVendor_FB.EditValue = DB.DBQuery("SELECT TOP(1) OIDVEND FROM Vendor WHERE Code=N'" + CusCode + "'").getString();
                        frmD01.slVendor_FB.Properties.View.PopulateColumns(frmD01.slVendor_FB.Properties.DataSource);
                        frmD01.slVendor_FB.Properties.View.Columns["ID"].Visible = false;
                    }
                }
            }
        }
    }
}