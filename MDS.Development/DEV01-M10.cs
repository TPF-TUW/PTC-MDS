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
    public partial class DEV01_M10 : DevExpress.XtraEditors.XtraForm
    {
        private Functionality.Function FUNCT = new Functionality.Function();
        //Global Variable
        //classConn db = new classConn();
        //classTools ct = new classTools();
        //SqlConnection mainConn = new classConn().MDS();
        //SqlConnection conn;
        string sql = "";

        int _UserID = 0;
        DatabaseConnect DB = new DatabaseConnect();
        public DEV01_M10(DatabaseConnect DBase, int UserID)
        {
            InitializeComponent();
            this._UserID = UserID;
            this.DB = DBase;
        }

        private void chkNull(string Alert, TextEdit txtName)
        {
            FUNCT.msgWarning("Please Key : "+Alert+"!"); txtName.Focus(); return;
        }

        private void btnAddCustomer_Click(object sender, EventArgs e)
        {
            string SizeNo = txeSizeNo.Text.ToString().ToUpper().Trim().Replace("'","''");
            string SizeName = txeSizeName.Text.ToString().Trim().Replace("'", "''");

            string strCREATE = this._UserID.ToString() != "" ? this._UserID.ToString() : "0";

            if (SizeNo == "") { chkNull("Size No.", txeSizeNo); }
            else if (SizeName == "") { chkNull("Size Name", txeSizeName); }
            else
            {
                //chkDup
                if (DB.DBQuery("SELECT TOP(1) SizeNo FROM ProductSize WHERE (SizeNo = N'" + SizeNo + "') ").getString() != "")
                {
                    FUNCT.msgWarning("Size No. is Duplicate!"); txeSizeNo.Focus(); return;
                }
                else
                {
                    if (FUNCT.msgQuiz("SAVE Size ?") == true)
                    {
                        sql = "INSERT INTO ProductSize(SizeNo, SizeName, CreatedBy, CreatedDate) VALUES(N'" + SizeNo + "', N'" + SizeName + "', '" + strCREATE + "', GETDATE()) ";
                        //Console.WriteLine(sql);
                        bool chkSave = DB.DBQuery(sql).runSQL();
                        if (chkSave == true)
                        {
                            FUNCT.msgInfo("Save Size is Successfull.");
                            this.Close();
                        }
                    }
                }
            }
        }

        private void DEV01_M10_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (Application.OpenForms.OfType<DEV01>().Count() > 0)
            {
                var frmD01 = Application.OpenForms.OfType<DEV01>().FirstOrDefault();
                frmD01.LoadSizeColor();
            }
        }

        private void DEV01_M10_Load(object sender, EventArgs e)
        {
            new ObjDE.setDatabase(this.DB);
            //MessageBox.Show(DB.getCONNECTION_STRING());
            //MessageBox.Show(ObjDE.GlobalVar.DBC.getCONNECTION_STRING());
        }
    }
}