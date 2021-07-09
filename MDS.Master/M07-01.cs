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
    public partial class M07_01 : DevExpress.XtraEditors.XtraForm
    {
        private Functionality.Function FUNC = new Functionality.Function();
        DatabaseConnect DB = new DatabaseConnect();
        int _UserID = 0;
        string _TypeName = "";
        string _TypeID = "";
        string _CodeID = "";
        public M07_01(DatabaseConnect DBase, int UserID, string TypeID, string TypeName, string CodeID) 
        {
            InitializeComponent();
            this.DB = DBase;
            this._UserID = UserID;
            this._TypeID = TypeID;
            this._TypeName = TypeName;
            this._CodeID = CodeID;
        }

        private void M07_01_Load(object sender, EventArgs e)
        {
            new ObjDE.setDatabase(this.DB);
            txeMaterialTypeID.Text = this._TypeID;
            txeMaterialType.Text = this._TypeName;
            txeItemCode.Focus();
        }

        private bool chkDuplicate()
        {
            bool chkDup = true;
            txeItemCode.Text = txeItemCode.Text.Trim();
            if (txeItemCode.Text != "")
            {
                string Code = txeItemCode.Text.ToString().ToUpper().Trim().Replace("'", "''").Replace(" ", "");
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT  TOP (1) OIDITEM FROM Items WHERE (Code = N'" + Code + "') ");
                if (this.DB.DBQuery(sbSQL).getString() != "")
                {
                    chkDup = false;
                }
            }
            return chkDup;
        }

        

        private void M07_01_FormClosed(object sender, FormClosedEventArgs e)
        {

        }

        private void M07_01_Shown(object sender, EventArgs e)
        {
            txeItemCode.Focus();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            string LineName = txeItemCode.Text.ToString().ToUpper().Trim().Replace("'", "''");
            if (LineName == "")
            {
                FUNC.msgWarning("Please input line name.");
                txeItemCode.Focus();
            }
            else
            {
                bool chkPass = true;
                if (txeItemCode.Text.ToUpper().Trim() != "")
                {
                    if (txeItemCode.Text.ToUpper().Trim().Length >= 5)
                        if (txeItemCode.Text.ToUpper().Trim().Substring(0, 5) == "TMPFB" || txeItemCode.Text.ToUpper().Trim().Substring(0, 5) == "TMPMT")
                        {
                            FUNC.msgWarning("Cannot set code starting with 'TMPFB' or 'TMPMT'. Please change code.");
                            txeItemCode.Focus();
                            chkPass = false;
                        }
                }

                if (chkPass == true)
                {
                    bool chkDup = chkDuplicate();
                    if (chkDup == true)
                    {
                        if (Application.OpenForms.OfType<M07>().Count() > 0)
                        {
                            var frmM07 = Application.OpenForms.OfType<M07>().FirstOrDefault();
                            if (this._CodeID == "")
                            {
                                StringBuilder sbSQL = new StringBuilder();
                                sbSQL.Append("SELECT ID, Code, Description ");
                                sbSQL.Append("FROM (");
                                sbSQL.Append("  SELECT OIDITEM AS ID, Code, Description ");
                                sbSQL.Append("  FROM  Items ");
                                sbSQL.Append("  WHERE (MaterialType = '" + this._TypeID + "') ");
                                sbSQL.Append("  UNION ALL ");
                                sbSQL.Append("  SELECT  99999 AS ID, '" + txeItemCode.Text.ToUpper().Trim() + "' AS Code, '' AS Description ");
                                sbSQL.Append(") AS LNN ");
                                sbSQL.Append("ORDER BY ID ");
                                new ObjDE.setGridLookUpEdit(frmM07.glueCode, sbSQL, "Code", "ID").getData(true);
                                frmM07.glueCode.Properties.View.PopulateColumns(frmM07.glueCode.Properties.DataSource);
                                frmM07.glueCode.Properties.View.Columns["ID"].Visible = false;
                                frmM07.glueCode.EditValue = 99999;
                                frmM07.txeDescription.Focus();
                            }
                            else
                            {
                                DataTable dtLN = (DataTable)frmM07.glueCode.Properties.DataSource;
                                if (dtLN.Rows.Count > 0)
                                {
                                    int runLoop = 0;
                                    foreach (DataRow drLN in dtLN.Rows)
                                    {
                                        string ID = drLN["ID"].ToString();
                                        if (ID == this._CodeID)
                                        {
                                            dtLN.Rows[runLoop].SetField("Code", txeItemCode.Text.ToUpper().Trim());
                                            break;
                                        }
                                        runLoop++;
                                    }

                                    frmM07.glueCode.Properties.DataSource = dtLN;
                                    frmM07.glueCode.Properties.DisplayMember = "Code";
                                    frmM07.glueCode.Properties.ValueMember = "ID";
                                    frmM07.glueCode.Properties.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
                                }
                                frmM07.glueCode.Properties.View.PopulateColumns(frmM07.glueCode.Properties.DataSource);
                                frmM07.glueCode.Properties.View.Columns["ID"].Visible = false;
                                frmM07.glueCode.EditValue = this._CodeID;
                            }
                        }
                        this.Close();
                    }
                    else
                    {
                        txeItemCode.Text = "";
                        txeItemCode.Focus();
                        FUNC.msgWarning("Duplicate Code !! Please change.\nรหัสซ้ำ กรุณาเปลี่ยน");
                    }
                }
            }
        }
    }
}