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
    public partial class DEV01_M07 : DevExpress.XtraEditors.XtraForm
    {
        private Functionality.Function FUNCT = new Functionality.Function();
        //classConn db = new classConn();
        //classTools ct = new classTools();
        string sql = "";
        //SqlConnection mainConn = new classConn().MDS();
        string ItemCode = "";

        const int TYPE_FG = 0;
        const int TYPE_FABRIC = 1;
        const int TYPE_ACCESSORY = 2;
        const int TYPE_PACKAGING = 3;
        const int TYPE_SAMPLE = 4;
        const int TYPE_OTHER = 9;
        const int TYPE_TEMPORARY = 8;

        string _tmpType = "";

        int _UserID = 0;
        DatabaseConnect DB = new DatabaseConnect();
        public DEV01_M07(DatabaseConnect DBase, string tmpType, int UserID)
        {
            InitializeComponent();
            this.DB = DBase;
            this._tmpType = tmpType;
            this._UserID = UserID;
        }

        private void DEV01_M07_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (Application.OpenForms.OfType<DEV01>().Count() > 0)
            {
                var frmD01 = Application.OpenForms.OfType<DEV01>().FirstOrDefault();
                StringBuilder sbSQL = new StringBuilder();
                if (this._tmpType == "FB")//Fabric
                {
                    sbSQL.Append("SELECT Code, Description, Type, ID "); //fabric & temporary
                    sbSQL.Append("FROM (");
                    sbSQL.Append("  SELECT Code, Description, 'Fabric' AS Type, MaterialType, OIDITEM AS ID FROM Items WHERE (MaterialType = '" + TYPE_FABRIC + "') ");
                    sbSQL.Append("  UNION ALL ");
                    sbSQL.Append("  SELECT Code, Description, 'Temporary' AS Type, MaterialType, OIDITEM AS ID FROM Items WHERE (MaterialType = '" + TYPE_TEMPORARY + "') AND (Code LIKE 'TMPFB%') ");
                    sbSQL.Append(") AS FBCode ");
                    sbSQL.Append("ORDER BY MaterialType, Code ");
                    new ObjDE.setSearchLookUpEdit(frmD01.slFBCode_FB, sbSQL, "Code", "ID").getData();
                    frmD01.slFBCode_FB.Properties.View.PopulateColumns(frmD01.slFBCode_FB.Properties.DataSource);
                    frmD01.slFBCode_FB.Properties.View.Columns["ID"].Visible = false;

                    frmD01.rep_slueFBCode.DataSource = frmD01.slFBCode_FB.Properties.DataSource;
                    frmD01.rep_slueFBCode.DisplayMember = "Code";
                    frmD01.rep_slueFBCode.ValueMember = "ID";
                    frmD01.rep_slueFBCode.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
                    frmD01.rep_slueFBCode.View.PopulateColumns(frmD01.rep_slueFBCode.DataSource);
                    frmD01.rep_slueFBCode.View.Columns["ID"].Visible = false;

                    frmD01.lblDescription.Text = "-";
                    frmD01.lblDescription.AppearanceItemCaption.BackColor = Color.Empty;
                }
                else if (this._tmpType == "MT") //Material
                {
                    sbSQL.Append("SELECT Code, Description, Type, ID "); //fabric & temporary
                    sbSQL.Append("FROM (");
                    sbSQL.Append("  SELECT Code, Description, 'Meterial' AS Type, MaterialType, OIDITEM AS ID FROM Items WHERE (MaterialType IN ('" + TYPE_ACCESSORY + "', '" + TYPE_PACKAGING + "')) ");
                    sbSQL.Append("  UNION ALL ");
                    sbSQL.Append("  SELECT Code, Description, 'Temporary' AS Type, MaterialType, OIDITEM AS ID FROM Items WHERE (MaterialType = '" + TYPE_TEMPORARY + "') AND (Code LIKE 'TMPMT%') ");
                    sbSQL.Append(") AS FBCode ");
                    sbSQL.Append("ORDER BY MaterialType, Code ");
                    new ObjDE.setSearchLookUpEdit(frmD01.slMatCode_Mat, sbSQL, "Code", "ID").getData();
                    frmD01.slMatCode_Mat.Properties.View.PopulateColumns(frmD01.slMatCode_Mat.Properties.DataSource);
                    frmD01.slMatCode_Mat.Properties.View.Columns["ID"].Visible = false;

                    frmD01.rep_MtrItem.DataSource = frmD01.slMatCode_Mat.Properties.DataSource;
                    frmD01.rep_MtrItem.DisplayMember = "Code";
                    frmD01.rep_MtrItem.ValueMember = "ID";
                    frmD01.rep_MtrItem.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
                    frmD01.rep_MtrItem.View.PopulateColumns(frmD01.rep_MtrItem.DataSource);
                    frmD01.rep_MtrItem.View.Columns["ID"].Visible = false;

                    frmD01.txeMatDescription.Text = "";
                    frmD01.txeMatDescription.BackColor = Color.Empty;
                }

                
                if (ItemCode != "")
                {
                    string[] arrFB = DB.DBQuery("SELECT TOP(1) Description, OIDITEM AS ID FROM Items WHERE (Code = N'" + ItemCode + "') ").getMultipleValue();
                    if (arrFB.Length > 0)
                    {
                        if (this._tmpType == "FB")//Fabric
                        {
                            frmD01.slFBCode_FB.EditValue = arrFB[1];
                            frmD01.lblDescription.Text = arrFB[0] == "" ? "-" : arrFB[0];
                            frmD01.lblDescription.AppearanceItemCaption.BackColor = Color.FromArgb(255, 255, 192);
                        }
                        else if (this._tmpType == "MT") //Material
                        {
                            frmD01.slMatCode_Mat.EditValue = arrFB[1];
                            frmD01.txeMatDescription.Text = arrFB[0] == "" ? "-" : arrFB[0];
                            frmD01.txeMatDescription.BackColor = Color.FromArgb(255, 255, 192);
                        }
                    }

                }
               
            }
        }

        private void btnAddItem_Click(object sender, EventArgs e)
        {
            string Description = txeDescription.Text.ToString().Trim().Replace("'", "''");

            string strCREATE = this._UserID.ToString() != "" ? this._UserID.ToString() : "0";

            //chkNull or Empty
            if (Description == "")
            {
                FUNCT.msgWarning("Please Key Description!"); txeDescription.Focus(); return;
            }
            else
            {
                //chkDup
                if (DB.DBQuery("SELECT TOP(1) OIDITEM FROM Items WHERE (Description = N'" + Description + "') ").getString() != "")
                {
                    FUNCT.msgWarning("Description is Duplicate!"); txeDescription.Focus(); return;
                }
                else
                {
                    //Confirm Save
                    if (FUNCT.msgQuiz("Save Item (Temporary) ? ") == true)
                    {
                        ItemCode = genNewItem();
                        //MessageBox.Show(ItemCode);
                        sql = "Insert Into Items (MaterialType, Code, Description, CreatedBy, CreatedDate, UpdatedBy, UpdatedDate) Values('" + TYPE_TEMPORARY + "', N'" + ItemCode + "', N'" + Description + "', '" + strCREATE + "', GETDATE(), '" + strCREATE + "', GETDATE())";
                        //Console.WriteLine(sql);
                        bool chkSave = DB.DBQuery(sql).runSQL();
                        if (chkSave == true)
                        {
                            FUNCT.msgInfo("Save Item (Temporary) is Successufull.");
                            this.Close();
                        }
                    }
                }
            }
        }

        private string genNewItem()
        {
            string newItem = "";
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT Code FROM Items WHERE (OIDITEM = (SELECT MAX(OIDITEM) AS OIDITEM FROM Items AS ITM WHERE (MaterialType = '" + TYPE_TEMPORARY + "') AND (Code LIKE 'TMP" + this._tmpType + "%'))) ");
            string xItem = DB.DBQuery(sbSQL.ToString()).getString();
            if (xItem == "")
                newItem = "TMP" + this._tmpType + "-00001";
            else
            {
                string subItem = (Convert.ToDouble(xItem.Substring(xItem.Length - 5, 5)) + 1).ToString("00000");
                newItem = "TMP" + this._tmpType + "-" + subItem;
            }
            return newItem;
        }

        private void DEV01_M07_Load(object sender, EventArgs e)
        {
            new ObjDE.setDatabase(this.DB);
            //MessageBox.Show(DB.getCONNECTION_STRING());
            //MessageBox.Show(ObjDE.GlobalVar.DBC.getCONNECTION_STRING());
        }
    }
}