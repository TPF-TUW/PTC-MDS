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
    public partial class DEV01_M09 : DevExpress.XtraEditors.XtraForm
    {
        private Functionality.Function FUNCT = new Functionality.Function();
        //classConn db = new classConn();
        //classTools ct = new classTools();
        string sql = "";
        //SqlConnection mainConn = new classConn().MDS();

        int _UserID = 0;
        string CategoryName = "";
        DatabaseConnect DB = new DatabaseConnect();
        public DEV01_M09(DatabaseConnect DBase, int UserID)
        {
            InitializeComponent();
            this.DB = DBase;
            this._UserID = UserID;
        }

        private void btnAddCategory_Click(object sender, EventArgs e)
        {
            CategoryName = txtCategoryName.Text.ToString().Trim().Replace("'","''");
            string strCREATE = this._UserID.ToString() != "" ? this._UserID.ToString() : "0";
            //chkNull or Empty
            if (CategoryName == "")
            {
                FUNCT.msgWarning("Please Key CategoryName!"); txtCategoryName.Focus(); return;
            }
            else
            {
                //chkDup
                if (DB.DBQuery("SELECT TOP(1) CategoryName FROM GarmentCategory WHERE CategoryName = N'" + CategoryName + "' ").getString() != "")
                {
                    FUNCT.msgWarning("CategroryName is Duplicate!"); txtCategoryName.Focus(); return;
                }
                else
                {
                    //Confirm Save
                    if (FUNCT.msgQuiz("Save CategoryName ?")==true)
                    {
                        sql = "INSERT INTO GarmentCategory (CategoryName, CreatedBy, CreatedDate) VALUES(N'" + CategoryName + "', '" + strCREATE + "', GETDATE())";
                        //Console.WriteLine(sql);
                        bool chkSave = DB.DBQuery(sql).runSQL();
                        if (chkSave == true)
                        {
                            FUNCT.msgInfo("Save CategoryName is Successufull.");
                            this.Close();
                        }
                    }
                }
            }
        }

        private void DEV01_M09_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (Application.OpenForms.OfType<DEV01>().Count() > 0)
            {
                var frmD01 = Application.OpenForms.OfType<DEV01>().FirstOrDefault();
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT CategoryName, OIDGCATEGORY AS ID FROM GarmentCategory ORDER BY CategoryName");
                new ObjDE.setSearchLookUpEdit(frmD01.glCategoryDivision_Main, sbSQL, "CategoryName", "ID").getData();
                if(CategoryName != "")
                    frmD01.glCategoryDivision_Main.EditValue = DB.DBQuery("SELECT TOP(1) OIDGCATEGORY FROM GarmentCategory WHERE CategoryName=N'" + CategoryName + "'").getString();
                frmD01.glCategoryDivision_Main.Properties.View.PopulateColumns(frmD01.glCategoryDivision_Main.Properties.DataSource);
                frmD01.glCategoryDivision_Main.Properties.View.Columns["ID"].Visible = false;
            }
        }

        private void DEV01_M09_Load(object sender, EventArgs e)
        {
            new ObjDE.setDatabase(this.DB);
            //MessageBox.Show(DB.getCONNECTION_STRING());
            //MessageBox.Show(ObjDE.GlobalVar.DBC.getCONNECTION_STRING());
        }
    }
}