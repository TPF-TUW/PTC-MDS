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
    public partial class DEV01_UF : DevExpress.XtraEditors.XtraForm
    {
        private Functionality.Function FUNCT = new Functionality.Function();
        int _UserID = 0;
        string UseFor = "";
        DatabaseConnect DB = new DatabaseConnect();
        public DEV01_UF(DatabaseConnect DBase, int UserID)
        {
            InitializeComponent();
            this.DB = DBase;
            this._UserID = UserID;
        }

        private void btnAddCategory_Click(object sender, EventArgs e)
        {
            UseFor = txtUseFor.Text.ToString().Trim().Replace("'","''");
            string strCREATE = this._UserID.ToString() != "" ? this._UserID.ToString() : "0";
            //chkNull or Empty
            if (UseFor == "")
            {
                FUNCT.msgWarning("Please Key UseFor!"); txtUseFor.Focus(); return;
            }
            else
            {
                //chkDup
                if (DB.DBQuery("SELECT TOP(1) OIDUF FROM SMPLUseFor WHERE UseFor = N'" + UseFor + "' ").getString() != "")
                {
                    FUNCT.msgWarning("UseFor is Duplicate!"); txtUseFor.Focus(); return;
                }
                else
                {
                    //Confirm Save
                    if (FUNCT.msgQuiz("Save UseFor ? ") ==true)
                    {
                        StringBuilder sbSQL = new StringBuilder();
                        sbSQL.Append("INSERT INTO SMPLUseFor (OIDUF, UseFor, CreatedBy) ");
                        sbSQL.Append(" SELECT MAX(OIDUF) + 1 AS OIDUF, N'" + UseFor + "' AS UseFor, '" + strCREATE + "' AS CreatedBy ");
                        sbSQL.Append(" FROM SMPLUseFor ");
                        //Console.WriteLine(sql);
                        bool chkSave = DB.DBQuery(sbSQL).runSQL();
                        if (chkSave == true)
                        {
                            FUNCT.msgInfo("Save UseFor is Successufull.");
                            this.Close();
                        }
                    }
                }
            }
        }

        private void DEV01_UF_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (Application.OpenForms.OfType<DEV01>().Count() > 0)
            {
                var frmD01 = Application.OpenForms.OfType<DEV01>().FirstOrDefault();
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT UseFor, OIDUF AS ID FROM SMPLUseFor ORDER BY OIDUF");
                new ObjDE.setGridLookUpEdit(frmD01.glUseFor, sbSQL, "UseFor", "ID").getData();
                if (UseFor != "")
                    frmD01.glUseFor.EditValue = DB.DBQuery("SELECT TOP(1) OIDUF FROM SMPLUseFor WHERE UseFor=N'" + UseFor + "'").getString();
                frmD01.glUseFor.Properties.View.PopulateColumns(frmD01.glUseFor.Properties.DataSource);
                frmD01.glUseFor.Properties.View.Columns["ID"].Visible = false;
            }
        }

        private void DEV01_UF_Load(object sender, EventArgs e)
        {
            new ObjDE.setDatabase(this.DB);
            //MessageBox.Show(DB.getCONNECTION_STRING());
            //MessageBox.Show(ObjDE.GlobalVar.DBC.getCONNECTION_STRING());
        }
    }
}