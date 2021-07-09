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
    public partial class DEV01_M13 : DevExpress.XtraEditors.XtraForm
    {
        private Functionality.Function FUNCT = new Functionality.Function();
        //classConn db = new classConn();
        //classTools ct = new classTools();
        string sql = "";
        //SqlConnection mainConn = new classConn().MDS();

        DatabaseConnect DB = new DatabaseConnect();
        int _UserID = 0;
        string UnitName = "";
        public DEV01_M13(DatabaseConnect DBase, int UserID)
        {
            InitializeComponent();
            this.DB = DBase;
            this._UserID = UserID;
        }

        private void btnAddCategory_Click(object sender, EventArgs e)
        {
            UnitName = txtUnit.Text.ToString().Trim().Replace("'","''");
            string strCREATE = this._UserID.ToString() != "" ? this._UserID.ToString() : "0";
            //chkNull or Empty
            if (UnitName == "")
            {
                FUNCT.msgWarning("Please Key Unit!"); txtUnit.Focus(); return;
            }
            else
            {
                //chkDup
                if (DB.DBQuery("SELECT OIDUNIT FROM Unit WHERE UnitName = N'" + UnitName + "' ").getString() != "")
                {
                    FUNCT.msgWarning("Unit is Duplicate!"); txtUnit.Focus(); return;
                }
                else
                {
                    //Confirm Save
                    if (FUNCT.msgQuiz("Save Unit ? ")==true)
                    {
                        sql = "INSERT INTO Unit (UnitName, CreatedBy, CreatedDate, UpdatedBy, UpdatedDate) VALUES(N'" + UnitName + "', '" + strCREATE + "', GETDATE(), '" + strCREATE + "', GETDATE())";
                        //Console.WriteLine(sql);
                        bool chkSave = DB.DBQuery(sql).runSQL();
                        if (chkSave == true)
                        {
                            FUNCT.msgInfo("Save Unit is Successufull.");
                            this.Close();
                        }
                    }
                }
            }
        }

        private void DEV01_M13_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (Application.OpenForms.OfType<DEV01>().Count() > 0)
            {
                var frmD01 = Application.OpenForms.OfType<DEV01>().FirstOrDefault();
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT UnitName, OIDUNIT AS ID FROM Unit ORDER BY UnitName");
                new ObjDE.setSearchLookUpEdit(frmD01.slueUnit, sbSQL, "UnitName", "ID").getData();
                if (UnitName != "")
                    frmD01.slueUnit.EditValue = DB.DBQuery("SELECT TOP(1) OIDUNIT FROM Unit WHERE UnitName=N'" + UnitName + "'").getString();
                frmD01.slueUnit.Properties.View.PopulateColumns(frmD01.slueUnit.Properties.DataSource);
                frmD01.slueUnit.Properties.View.Columns["ID"].Visible = false;

                frmD01.slueConsumpUnit.Properties.DataSource = frmD01.slueUnit.Properties.DataSource;
                frmD01.slueConsumpUnit.Properties.DisplayMember = "UnitName";
                frmD01.slueConsumpUnit.Properties.ValueMember = "ID";
                frmD01.slueConsumpUnit.Properties.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
                frmD01.slueConsumpUnit.Properties.View.PopulateColumns(frmD01.slueConsumpUnit.Properties.DataSource);
                frmD01.slueConsumpUnit.Properties.View.Columns["ID"].Visible = false;

            }
        }

        private void DEV01_M13_Load(object sender, EventArgs e)
        {
            new ObjDE.setDatabase(this.DB);
            //MessageBox.Show(DB.getCONNECTION_STRING());
            //MessageBox.Show(ObjDE.GlobalVar.DBC.getCONNECTION_STRING());
        }
    }
}