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
    public partial class M08_01 : DevExpress.XtraEditors.XtraForm
    {
        private Functionality.Function FUNC = new Functionality.Function();
        DatabaseConnect DB = new DatabaseConnect();
        int _UserID = 0;
        string _BranchName = "";
        string _BranchID = "";
        string _LineID = "";
        public M08_01(DatabaseConnect DBase, int UserID, string BranchID, string BranchName, string LineID) 
        {
            InitializeComponent();
            this.DB = DBase;
            this._UserID = UserID;
            this._BranchID = BranchID;
            this._BranchName = BranchName;
            this._LineID = LineID;
        }

        private void M08_01_Load(object sender, EventArgs e)
        {
            new ObjDE.setDatabase(this.DB);
            txeBranchID.Text = this._BranchID;
            txeBranch.Text = this._BranchName;
            txeLineName.Focus();
        }

        private bool chkDuplicate()
        {
            bool chkDup = true;
            txeLineName.Text = txeLineName.Text.Trim();
            if (txeLineName.Text != "")
            {
                string LineName = txeLineName.Text.ToString().ToUpper().Trim().Replace("'", "''").Replace(" ", "");
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT OIDLINE FROM LineNumber WHERE (Branch = '" + txeBranchID.Text.Trim() + "') AND (REPLACE(LINENAME, ' ', '') = N'" + LineName + "') ");
                if (this.DB.DBQuery(sbSQL).getString() != "")
                {
                    chkDup = false;
                }
            }
            return chkDup;
        }

        

        private void M08_01_FormClosed(object sender, FormClosedEventArgs e)
        {
        }

        private void M08_01_Shown(object sender, EventArgs e)
        {
            txeLineName.Focus();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            string LineName = txeLineName.Text.ToString().ToUpper().Trim().Replace("'", "''");
            if (LineName == "")
            {
                FUNC.msgWarning("Please input line name.");
                txeLineName.Focus();
            }
            else
            {
                bool chkDup = chkDuplicate();
                if (chkDup == true)
                {
                    if (Application.OpenForms.OfType<M08>().Count() > 0)
                    {
                        var frmM08 = Application.OpenForms.OfType<M08>().FirstOrDefault();
                        if (this._LineID == "")
                        {
                            StringBuilder sbSQL = new StringBuilder();
                            sbSQL.Append("SELECT ID, LINENAME, Branch, BranchID ");
                            sbSQL.Append("FROM (");
                            sbSQL.Append("  SELECT  LN.OIDLINE AS ID, LN.LINENAME, B.Name AS Branch, LN.Branch AS BranchID ");
                            sbSQL.Append("  FROM    LineNumber AS LN INNER JOIN ");
                            sbSQL.Append("          Branchs AS B ON LN.Branch = B.OIDBranch ");
                            sbSQL.Append("  WHERE (B.OIDBranch = '" + txeBranchID.Text.Trim() + "') ");
                            sbSQL.Append("  UNION ALL ");
                            sbSQL.Append("  SELECT  99999 AS ID, '" + txeLineName.Text.ToUpper().Trim() + "' AS LINENAME, '" + txeBranch.Text.Trim() + "' AS Branch, '" + txeBranchID.Text.Trim() + "' AS BranchID ");
                            sbSQL.Append(") AS LNN ");
                            sbSQL.Append("ORDER BY ID ");
                            new ObjDE.setSearchLookUpEdit(frmM08.glueLineName, sbSQL, "LINENAME", "ID").getData(true);
                            frmM08.glueLineName.EditValue = 99999;
                        }
                        else
                        {
                            DataTable dtLN = (DataTable)frmM08.glueLineName.Properties.DataSource;
                            if (dtLN.Rows.Count > 0)
                            {
                                int runLoop = 0;
                                foreach (DataRow drLN in dtLN.Rows)
                                {
                                    string ID = drLN["ID"].ToString();
                                    if (ID == this._LineID)
                                    {
                                        dtLN.Rows[runLoop].SetField("LINENAME", txeLineName.Text.ToUpper().Trim());
                                        break;
                                    }
                                    runLoop++;
                                }

                                frmM08.glueLineName.Properties.DataSource = dtLN;
                                frmM08.glueLineName.Properties.DisplayMember = "LINENAME";
                                frmM08.glueLineName.Properties.ValueMember = "ID";
                                frmM08.glueLineName.Properties.BestFitMode = DevExpress.XtraEditors.Controls.BestFitMode.BestFitResizePopup;
                            }

                            frmM08.glueLineName.EditValue = this._LineID;
                        }
                    }
                    this.Close();
                }
                else
                {
                    txeLineName.Text = "";
                    txeLineName.Focus();
                    FUNC.msgWarning("Duplicate line name. !! Please Change.");
                }
            }
        }
    }
}