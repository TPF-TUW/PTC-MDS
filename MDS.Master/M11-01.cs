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
    public partial class M11_01 : DevExpress.XtraEditors.XtraForm
    {
        private Functionality.Function FUNC = new Functionality.Function();
        DatabaseConnect DB = new DatabaseConnect();
        int _UserID = 0;
        string CategoryName = "";
        public M11_01(DatabaseConnect DBase, int UserID)
        {
            InitializeComponent();
            this.DB = DBase;
            this._UserID = UserID;
        }

        private void M11_01_Load(object sender, EventArgs e)
        {
            new ObjDE.setDatabase(this.DB);
        }

        private bool chkDuplicate()
        {
            bool chkDup = true;
            txeCategoryName.Text = txeCategoryName.Text.Trim();
            if (txeCategoryName.Text != "")
            {
                string StyleName = txeCategoryName.Text.ToString().Trim().Replace("'", "''");
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT OIDGCATEGORY FROM GarmentCategory WHERE (CategoryName = N'" + StyleName + "') ");
                if (this.DB.DBQuery(sbSQL).getString() != "")
                {
                    chkDup = false;
                }
            }
            return chkDup;
        }

        private void btnAddStyle_Click(object sender, EventArgs e)
        {
            string StyleName = txeCategoryName.Text.ToString().Trim().Replace("'", "''");
            CategoryName = StyleName;
            //chkNull or Empty
            if (StyleName == "")
            {
                FUNC.msgWarning("Please input category name.");
                txeCategoryName.Focus();
            }
            else
            {

                bool chkDup = chkDuplicate();
                if (chkDup == true)
                {
                    if (FUNC.msgQuiz("Confirm save data ?") == true)
                    {
                        string strCREATE = this._UserID.ToString() != "" ? this._UserID.ToString() : "0";

                        StringBuilder sbSQL = new StringBuilder();
                        sbSQL.Append(" INSERT INTO GarmentCategory(CategoryName, CreatedBy, CreatedDate) ");
                        sbSQL.Append("  VALUES(N'" + StyleName + "', '" + strCREATE + "', GETDATE()) ");
                        try
                        {
                            bool chkSAVE = this.DB.DBQuery(sbSQL).runSQL();
                            if (chkSAVE == true)
                            {
                                if (Application.OpenForms.OfType<M11>().Count() > 0)
                                {
                                    M11 frmStyle = Application.OpenForms.OfType<M11>().First();
                                    sbSQL.Clear();
                                    sbSQL.Append("SELECT CategoryName, OIDGCATEGORY AS ID ");
                                    sbSQL.Append("FROM GarmentCategory ");
                                    sbSQL.Append("ORDER BY CategoryName ");
                                    new ObjDE.setGridLookUpEdit(frmStyle.glueCategory, sbSQL, "CategoryName", "ID").getData(true);
                                }
                                FUNC.msgInfo("Save complete.");
                                this.Close();
                            }
                        }
                        catch (Exception)
                        { }
                    }
                }
                else
                {
                    txeCategoryName.Text = "";
                    txeCategoryName.Focus();
                    FUNC.msgWarning("Duplicate category name. !! Please Change.");
                }
            }
        }

        private void M11_01_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (Application.OpenForms.OfType<M11>().Count() > 0)
            {
                var frmD01 = Application.OpenForms.OfType<M11>().FirstOrDefault();
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT CategoryName, OIDGCATEGORY AS ID ");
                sbSQL.Append("FROM GarmentCategory ");
                sbSQL.Append("ORDER BY CategoryName ");
                new ObjDE.setGridLookUpEdit(frmD01.glueCategory, sbSQL, "CategoryName", "ID").getData(true);
                if (CategoryName != "")
                    frmD01.glueCategory.EditValue = DB.DBQuery("SELECT TOP(1) OIDGCATEGORY FROM GarmentCategory WHERE CategoryName=N'" + CategoryName + "'").getString();
                frmD01.glueCategory.Properties.View.PopulateColumns(frmD01.glueCategory.Properties.DataSource);
                frmD01.glueCategory.Properties.View.Columns["ID"].Visible = false;
            }
        }
    }
}