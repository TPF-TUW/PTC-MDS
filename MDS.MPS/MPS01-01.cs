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
using DBConnection;

namespace MDS.MPS
{
    public partial class MPS01_01 : DevExpress.XtraEditors.XtraForm
    {
        ////Global Variable
        //classConn db = new classConn();
        //classTools ct = new classTools();
        //SqlConnection mainConn = new classConn().MDS();
        SqlConnection conn;
        string sql = string.Empty;
        private Functionality.Function FUNC = new Functionality.Function();
        public MPS01_01()
        {
            InitializeComponent();
        }

        //private void chkNull(string Alert, TextEdit txtName)
        //{
        //    ct.showWarningMessage("Please Key : "+Alert+"!"); txtName.Focus(); return;
        //}

        private void btnAddCustomer_Click(object sender, EventArgs e)
        {
            if (slueCustomer.Text.Trim() == "")
            {
                FUNC.msgWarning("Please select customer.");
                slueCustomer.Focus();
            }
            else if (txeItemCode.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input item code.");
                txeItemCode.Focus();
            }
            else if (txeItemName.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input item name.");
                txeItemName.Focus();
            }
            else if (txeStyleNo.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input style no.");
                txeStyleNo.Focus();
            }
            else if (glueSeason.Text.Trim() == "")
            {
                FUNC.msgWarning("Please select season.");
                glueSeason.Focus();
            }
            else
            {
                txeItemCode.Text = txeItemCode.Text.ToUpper().Trim();
                bool chkGMP = chkDuplicate();

                if (chkGMP == true)
                {
                    if (FUNC.msgQuiz("Confirm save data ?") == true)
                    {
                        string OIDSTYLE = slueStyle.Text.Trim() != "" ? "'" + slueStyle.EditValue.ToString() + "'" : "NULL";
                        string Season = glueSeason.Text.Trim() != "" ? speSeason.Value.ToString() + glueSeason.EditValue.ToString() : "";

                        StringBuilder sbSQL = new StringBuilder();
                        sbSQL.Append("  INSERT INTO ItemCustomer(OIDCUST, ItemCode, ItemName, OIDSTYLE, Season, FabricWidth, FBComposition, StyleNo) ");
                        sbSQL.Append("  VALUES(N'" + slueCustomer.EditValue.ToString() + "', N'" + txeItemCode.Text.Trim().Replace("'", "''") + "', N'" + txeItemName.Text.Trim().Replace("'", "''") + "', " + OIDSTYLE + ", N'" + Season + "', N'" + txeFabricWidth.Text.Trim().Replace("'", "''") + "', N'" + txeFBComposition.Text.Trim().Replace("'", "''") + "', N'" + txeStyleNo.Text.Trim() + txeStyleCode.Text.Trim() + "') ");

                        if (sbSQL.Length > 0)
                        {
                            try
                            {
                                bool chkSAVE = new DBQuery(sbSQL).runSQL();
                                if (chkSAVE == true)
                                {
                                    FUNC.msgInfo("Save complete.");
                                    NewData();
                                }
                            }
                            catch (Exception)
                            { }
                        }
                    }
                }
            }
        }

        private void NewData()
        {
            slueCustomer.EditValue = "";
            txeItemCode.Text = "";
            txeItemName.Text = "";
            txeStyleNo.Text = "";
            txeStyleCode.Text = "";
            txeFBComposition.Text = "";
            txeFabricWidth.Text = "";
            speSeason.Value = Convert.ToInt32(DateTime.Now.ToString("yyyy"));
            glueSeason.EditValue = "";
            slueStyle.EditValue = "";
            slueCustomer.Focus();
        }

        private void MPS01_01_FormClosed(object sender, FormClosedEventArgs e)
        {
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT ITC.ItemCode, ITC.ItemName, CUS.Name AS Customer, ITC.StyleNo AS [StyleNo.], ITC.Season, ITC.OIDCSITEM AS ID ");
            sbSQL.Append("FROM   ItemCustomer AS ITC LEFT OUTER JOIN ");
            sbSQL.Append("       Customer AS CUS ON ITC.OIDCUST = CUS.OIDCUST ");
            sbSQL.Append("ORDER BY ITC.ItemCode ");
            MPS01 frmMPS = Application.OpenForms.OfType<MPS01>().First();
            new ObjDevEx.setSearchLookUpEdit(frmMPS.slueItemCode, sbSQL, "ItemCode", "ID").getData();
        }

        private void MPS01_01_Load(object sender, EventArgs e)
        {
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT Code, Name AS Customer, OIDCUST AS ID ");
            sbSQL.Append("FROM Customer ");
            sbSQL.Append("ORDER BY Code ");
            new ObjDevEx.setSearchLookUpEdit(slueCustomer, sbSQL, "Customer", "ID").getData();

            sbSQL.Clear();
            sbSQL.Append("SELECT SeasonNo AS [Season No.], SeasonName AS [Season Name] ");
            sbSQL.Append("FROM Season ");
            sbSQL.Append("ORDER BY OIDSEASON");
            new ObjDevEx.setGridLookUpEdit(glueSeason, sbSQL, "Season No.", "Season No.").getData();

            sbSQL.Clear();
            sbSQL.Append("SELECT PS.StyleName, GC.CategoryName, PS.OIDSTYLE AS ID ");
            sbSQL.Append("FROM   ProductStyle AS PS INNER JOIN ");
            sbSQL.Append("       GarmentCategory AS GC ON PS.OIDGCATEGORY = GC.OIDGCATEGORY ");
            sbSQL.Append("ORDER BY PS.StyleName ");
            new ObjDevEx.setSearchLookUpEdit(slueStyle, sbSQL, "StyleName", "ID").getData();

            NewData();
        }

        private void slueCustomer_EditValueChanged(object sender, EventArgs e)
        {
            bool chkDup = chkDuplicate();
            if (chkDup == true)
                slueStyle.Focus();
        }

        private void txeItemCode_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeItemName.Focus();
            }
        }

        private void txeItemCode_Leave(object sender, EventArgs e)
        {
            if (txeItemCode.Text.Trim() != "")
            {
                bool chkDup = chkDuplicate();
                if (chkDup == false)
                {
                    txeItemCode.Text = "";
                    txeItemCode.Focus();
                    FUNC.msgWarning("Duplicate item code. !! Please Change.");
                }
            }
        }

        private bool chkDuplicate()
        {
            bool chkDup = true;
            if (txeItemCode.Text != "")
            {
                txeItemCode.Text = txeItemCode.Text.ToUpper().Trim();
                string CUST = slueCustomer.Text.Trim() != "" ? slueCustomer.EditValue.ToString() : "";
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT TOP(1) ItemCode FROM ItemCustomer WHERE (OIDCUST = '" + CUST + "') AND (ItemCode = N'" + txeItemCode.Text.Trim().Replace("'", "''") + "') ");
                if (new DBQuery(sbSQL).getString() != "")
                {
                    chkDup = false;
                }
            }
            return chkDup;
        }

        private void txeItemName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeStyleNo.Focus();
            }
        }

        private void txeStyleNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                speSeason.Focus();
            }
        }

        private void glueSeason_EditValueChanged(object sender, EventArgs e)
        {
            txeFabricWidth.Focus();
        }

        private void txeFabricWidth_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                txeFBComposition.Focus();
        }

        private void slueStyle_EditValueChanged(object sender, EventArgs e)
        {
            txeStyleCode.Text = slueStyle.Text;
            txeItemCode.Focus();
        }
    }
}