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
    public partial class M07_03 : DevExpress.XtraEditors.XtraForm
    {
        private Functionality.Function FUNCT = new Functionality.Function();
        //Global Var
        //classConn db = new classConn();
        //classTools ct = new classTools();
        //SqlConnection mainConn = new classConn().MDS();
        string sql = "";

        string _Type = "";

        int _UserID = 0;
        string StyleName = "";
        DatabaseConnect DB = new DatabaseConnect();
        public M07_03(DatabaseConnect DBase, string Type, int UserID)
        {
            InitializeComponent();
            this.DB = DBase;
            this._Type = Type;
            this._UserID = UserID;
        }

        private void M07_03_Load(object sender, EventArgs e)
        {
            new ObjDE.setDatabase(this.DB);
            //MessageBox.Show(DB.getCONNECTION_STRING());
            //MessageBox.Show(ObjDE.GlobalVar.DBC.getCONNECTION_STRING());

            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT CategoryName, OIDGCATEGORY AS ID FROM GarmentCategory ORDER BY CategoryName");
            new ObjDE.setGridLookUpEdit(glCategoryName, sbSQL, "CategoryName", "ID").getData();
            glCategoryName.Properties.View.PopulateColumns(glCategoryName.Properties.DataSource);
            glCategoryName.Properties.View.Columns["ID"].Visible = false;

            glCategoryName.EditValue = this._Type;
            txtStyleName.Focus();
        }

        private void btnAddStyle_Click(object sender, EventArgs e)
        {
            StyleName = txtStyleName.Text.ToString().Trim().Replace("'", "''");
            string CategoryName = glCategoryName.Text.ToString();

            string strCREATE = this._UserID.ToString() != "" ? this._UserID.ToString() : "0";

            //chkNull or Empty
            if (StyleName == "")
            {
                FUNCT.msgWarning("Please Key StyleName!"); txtStyleName.Focus(); return;
            }
            else if (CategoryName == "")
            {
                FUNCT.msgWarning("Please Select CategoryName!"); glCategoryName.Focus(); return;
            }
            else
            {
                //chkDup
                if (DB.DBQuery("SELECT TOP(1) StyleName FROM ProductStyle WHERE StyleName = '" + StyleName + "' ").getString() != "")
                {
                    FUNCT.msgWarning("StyleName is Duplicate!"); txtStyleName.Focus(); return;
                }
                else
                {
                    //Confirm Save
                    if (FUNCT.msgQuiz("Save StyleName ? ") == true)
                    {
                        sql = "INSERT INTO ProductStyle (StyleName, OIDGCATEGORY, CreatedBy, CreatedDate) VALUES(N'" + StyleName + "', '"+ glCategoryName.EditValue.ToString() + "', '" + strCREATE + "', GETDATE())";
                        //Console.WriteLine(sql);
                        bool chkSave = DB.DBQuery(sql).runSQL();
                        if (chkSave == true)
                        {
                            FUNCT.msgInfo("Save StyleName is Successufull.");
                            this.Close();
                        }
                    }
                }
            }
        }

        private void M07_03_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (Application.OpenForms.OfType<M07>().Count() > 0)
            {
                var frmD01 = Application.OpenForms.OfType<M07>().FirstOrDefault();
                StringBuilder sbSQL = new StringBuilder();

                sbSQL.Append("SELECT StyleName, OIDSTYLE AS ID ");
                sbSQL.Append("FROM  ProductStyle ");
                if (frmD01.slueCategory.Text.Trim() != "")
                    sbSQL.Append("WHERE (OIDGCATEGORY = '" + frmD01.slueCategory.EditValue.ToString() + "') ");
                sbSQL.Append("ORDER BY StyleName ");
                new ObjDE.setSearchLookUpEdit(frmD01.slueStyle, sbSQL, "StyleName", "ID").getData();
                if (StyleName != "")
                    frmD01.slueStyle.EditValue = DB.DBQuery("SELECT TOP(1) OIDSTYLE FROM ProductStyle WHERE StyleName=N'" + StyleName + "'").getString();
                frmD01.slueStyle.Properties.View.PopulateColumns(frmD01.slueStyle.Properties.DataSource);
                frmD01.slueStyle.Properties.View.Columns["ID"].Visible = false;
            }
        }
    }
}