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
    public partial class DEV01_M03 : DevExpress.XtraEditors.XtraForm
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
        string ColorNo = "";

        DatabaseConnect DB = new DatabaseConnect();

        public DEV01_M03(DatabaseConnect DBase, string Type, int UserID)
        {
            InitializeComponent();
            this.DB = DBase;
            this._Type = Type;
            this._UserID = UserID;
        }

        private void chkNull(string Alert, TextEdit txtName)
        {
            FUNCT.msgWarning("Please Key : " + Alert + "!"); txtName.Focus(); return;
        }

        private void btnAddCustomer_Click(object sender, EventArgs e)
        {
            ColorNo = txeColorNo.Text.ToString().ToUpper().Trim().Replace("'","''");
            string ColorName = txeColorName.Text.ToString().Trim().Replace("'", "''");
            string ColorType = cbeColorType.EditValue.ToString();

            string strCREATE = this._UserID.ToString() != "" ? this._UserID.ToString() : "0";

            if (ColorNo == "") { chkNull("Color No.", txeColorNo); }
            else if (ColorName == "") { chkNull("Color Name", txeColorName); }
            else if (cbeColorType.Text.Trim() == "") { chkNull("Color Type", cbeColorType); }
            else
            {
                //chkDup
                if (DB.DBQuery("SELECT TOP(1) ColorNo FROM ProductColor WHERE (ColorType = '" + ColorType + "') AND (ColorNo = N'" + ColorNo + "') ").getString() != "")
                {
                    FUNCT.msgWarning("Color No. is Duplicate!"); txeColorNo.Focus(); return;  
                }
                else
                {
                    if (FUNCT.msgQuiz("SAVE Color ?") == true)
                    {
                        sql = "INSERT INTO ProductColor(ColorNo, ColorName, ColorType, CreatedBy, CreatedDate) VALUES(N'" + ColorNo + "', N'" + ColorName + "', '" + ColorType + "', '" + strCREATE + "', GETDATE()) ";
                        //Console.WriteLine(sql);
                        bool chkSave = DB.DBQuery(sql).runSQL();
                        if (chkSave == true)
                        {
                            FUNCT.msgInfo("Save Color is Successfull.");
                            this.Close();
                        }
                    }
                }
            }
        }

        private void DEV01_M03_Load(object sender, EventArgs e)
        {
            new ObjDE.setDatabase(this.DB);
            //MessageBox.Show(DB.getCONNECTION_STRING());
            //MessageBox.Show(ObjDE.GlobalVar.DBC.getCONNECTION_STRING());

            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT '0' AS ID, 'Finished Goods' AS ColorType ");
            sbSQL.Append("UNION ALL ");
            sbSQL.Append("SELECT '1' AS ID, 'Fabric' AS ColorType ");
            sbSQL.Append("UNION ALL ");
            sbSQL.Append("SELECT '2' AS ID, 'Accessory' AS ColorType ");
            sbSQL.Append("UNION ALL ");
            sbSQL.Append("SELECT '3' AS ID, 'Packaging' AS ColorType ");
            new ObjDE.setGridLookUpEdit(cbeColorType, sbSQL, "ColorType", "ID").getData();
            cbeColorType.Properties.View.PopulateColumns(cbeColorType.Properties.DataSource);
            cbeColorType.Properties.View.Columns["ID"].Visible = false;

            if (this._Type == "FG") //Finished Goods
            {
                cbeColorType.EditValue = 0;
            }
            else if (this._Type == "FB") //Fabric
            {
                cbeColorType.EditValue = 1;
            }
            else if (this._Type == "MT") //Material
            {
                cbeColorType.EditValue = 2;
            }
            txeColorNo.Focus();
        }

        private void DEV01_M03_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (Application.OpenForms.OfType<DEV01>().Count() > 0)
            {
                var frmD01 = Application.OpenForms.OfType<DEV01>().FirstOrDefault();
                if (cbeColorType.EditValue.ToString() == "0")
                    frmD01.LoadSizeColor();
                else if (cbeColorType.EditValue.ToString() == "1") //Fabric
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT ColorName, OIDCOLOR AS ID FROM ProductColor WHERE (ColorType = 1) ORDER BY ColorName");
                    new ObjDE.setSearchLookUpEdit(frmD01.slFBColor_FB, sbSQL, "ColorName", "ID").getData();
                    if(ColorNo != "")
                        frmD01.slFBColor_FB.EditValue = DB.DBQuery("SELECT OIDCOLOR FROM ProductColor WHERE ColorType=1 AND ColorNo=N'" + ColorNo + "'").getString();
                    frmD01.slFBColor_FB.Properties.View.PopulateColumns(frmD01.slFBColor_FB.Properties.DataSource);
                    frmD01.slFBColor_FB.Properties.View.Columns["ID"].Visible = false;
                }
                else if (cbeColorType.EditValue.ToString() == "2" || cbeColorType.EditValue.ToString() == "3") //Material
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT ColorName, OIDCOLOR AS ID FROM ProductColor WHERE (ColorType IN (2, 3)) ORDER BY ColorName");
                    new ObjDE.setSearchLookUpEdit(frmD01.slMatColor_Mat, sbSQL, "ColorName", "ID").getData();
                    if(ColorNo != "")
                        frmD01.slMatColor_Mat.EditValue = DB.DBQuery("SELECT TOP(1) OIDCOLOR FROM ProductColor WHERE ColorType IN (2, 3) AND ColorNo=N'" + ColorNo + "'").getString();
                    frmD01.slMatColor_Mat.Properties.View.PopulateColumns(frmD01.slMatColor_Mat.Properties.DataSource);
                    frmD01.slMatColor_Mat.Properties.View.Columns["ID"].Visible = false;
                }
            }
        }
    }
}