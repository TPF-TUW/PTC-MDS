using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using Microsoft.Win32;
using DBConnect;
using TheepClass;
using DevExpress.LookAndFeel;

namespace MDS00
{
    public partial class frmLogin : DevExpress.XtraEditors.XtraForm
    {
        cDatabase db = new cDatabase("Server=domain-ii;uid=sa;pwd=ZAQ113m4tuw;database=MDS");
        //private LogIn User_Login;

        public string ConnectionString { get; set; }
        string CONNECT_STRING = "";
        DatabaseConnect DBC;

        public frmLogin()
        {
            InitializeComponent();
            var skinName = cUtility.LoadRegistry(@"Software\MDS", "SkinName");
            var skinPalette = cUtility.LoadRegistry(@"Software\MDS", "SkinPalette");
            UserLookAndFeel.Default.SetSkinStyle(skinName == null ? "Basic" : skinName.ToString(), skinPalette == null ? "Default" : skinPalette.ToString());
            var userName = cUtility.LoadRegistry(@"Software\MDS", "UserName");
            txtUsername.Text = userName == null ? "" : userName.ToString();
        }
        
        private LogIn VerifyLogin(string userName, string passWord, string company)
        {
            db.ConnectionOpen();
            if (db.ExecuteFirstValue("SELECT COUNT(USERNAME) FROM Users WHERE OIDCompany='" + company + "' AND UserName='" + userName + "'")=="0")
            {
                MessageBox.Show("User Name is not valid.","Warning",MessageBoxButtons.OK,MessageBoxIcon.Warning);
                db.ConnectionClose();
                return null;
            }
            if (db.ExecuteFirstValue("SELECT COUNT(PASSWORD) FROM Users WHERE OIDCompany='" + company + "' AND UserName='" + userName+"' AND Password='"+passWord+"'")=="0")
            {
                MessageBox.Show("Password is not correct.","Warning",MessageBoxButtons.OK,MessageBoxIcon.Warning);
                return null;
            }
            db.ConnectionClose();
            string strSQL = "SELECT B.OIDUSER,B.USERNAME,B.FULLNAME,A.FUNCTIONNO,A.READWRITESTATUS,A.ALLOWDENYSTATUS "+ 
                "FROM FunctionAccess A INNER JOIN Users B ON A.OIDUser = B.OIDUSER "+
                "WHERE B.OIDCompany='" + company + "' AND B.UserName = '" + userName+"'";
            DataTable dtLogin = db.GetDataTable(strSQL);
            var User_Login = new LogIn();
            var functions = new List<LogIn_Function>();
            foreach (DataRow dr in dtLogin.Rows)
            {
                if(dtLogin.Rows.IndexOf(dr)==0)
                {
                    User_Login.OIDUser = Convert.ToInt32(dr["OIDUSER"]);
                    User_Login.UserName = dr["USERNAME"].ToString();
                    User_Login.FullName = dr["FULLNAME"].ToString();
                }
                functions.Add(new LogIn_Function {
                    FunctionNo = dr["FUNCTIONNO"].ToString(),
                    ReadWriteStatus = Convert.ToInt32(dr["READWRITESTATUS"]),
                    AllowDenyStatus = Convert.ToInt32(dr["ALLOWDENYSTATUS"])
                }); ;
            }
            User_Login.Functions = functions;
            return User_Login;

        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Dispose();
            Application.Exit();
        }
        private void btnOK_Click(object sender, EventArgs e)
        {
            if (glueCompany.Text == "")
            {
                MessageBox.Show("Please select company.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                glueCompany.Focus();
            }
            else
            {
                try
                {
                    var userLogin = VerifyLogin(txtUsername.Text, txtPassword.Text, glueCompany.EditValue.ToString());
                    if (userLogin == null)
                        return;
                    else
                        cUtility.SaveRegistry(@"Software\MDS", "UserName", userLogin.UserName);
                    this.Hide();
                    XtraForm3 frmMain = new XtraForm3();
                    frmMain.Company = Convert.ToInt32(glueCompany.EditValue.ToString());
                    frmMain.UserLogin = userLogin;
                    frmMain.WindowState = FormWindowState.Maximized;
                    frmMain.Show();
                }
                catch (SystemException ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch (ApplicationException ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        private void txtPassword_KeyPress(object sender, KeyPressEventArgs e)
        {
            if(e.KeyChar==(char)13){ btnOK_Click(sender,e);}
        }
        private void frmLogin_Shown(object sender, EventArgs e)
        {
            if (txtUsername.Text.Length == 0)
                txtUsername.Focus();
            else
                txtPassword.Focus();
        }

        private void frmLogin_Load(object sender, EventArgs e)
        {
            //***** SET CONNECT DB ********
            if (this.ConnectionString != null)
            {
                if (this.ConnectionString != "")
                {
                    CONNECT_STRING = this.ConnectionString;
                }
            }

            this.DBC = new DatabaseConnect(CONNECT_STRING);

            if (this.DBC.chkCONNECTION_STING() == false)
            {
                this.DBC.setCONNECTION_STRING_INIFILE();
                if (this.DBC.chkCONNECTION_STING() == false)
                {
                    return;
                }
            }
            new ObjDE.setDatabase(this.DBC);
            //*****************************

            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT Code AS [Company Code], EngName AS[Name (En)], THName AS[Name (Th)], OIDCOMPANY AS ID ");
            sbSQL.Append("FROM Company ");
            sbSQL.Append("ORDER BY OIDCOMPANY ");
            new ObjDE.setGridLookUpEdit(glueCompany, sbSQL, "Company Code", "ID").getData();
            glueCompany.Properties.View.PopulateColumns(glueCompany.Properties.DataSource);
            glueCompany.Properties.View.Columns["ID"].Visible = false;
        }
    }

}