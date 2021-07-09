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
    public partial class DEV01_M06 : DevExpress.XtraEditors.XtraForm
    {
        private Functionality.Function FUNCT = new Functionality.Function();
        //classConn db = new classConn();
        //classTools ct = new classTools();
        string sql = "";
        //SqlConnection mainConn = new classConn().MDS();

        int _UserID = 0;
        string FabricPart = "";

        string strFBID = "";
        DatabaseConnect DB = new DatabaseConnect();
        public DEV01_M06(DatabaseConnect DBase, int UserID)
        {
            InitializeComponent();
            this.DB = DBase;
            this._UserID = UserID;
        }

        private void btnAddCategory_Click(object sender, EventArgs e)
        {
            FabricPart = txtFabricParts.Text.ToString().Trim().Replace("'","''");
            string strCREATE = this._UserID.ToString() != "" ? this._UserID.ToString() : "0";
            //chkNull or Empty
            if (FabricPart == "")
            {
                FUNCT.msgWarning("Please Key Fabric Parts !"); txtFabricParts.Focus(); return;
            }
            else
            {
                //chkDup
                if (DB.DBQuery("SELECT TOP(1) OIDGParts FROM GarmentParts WHERE GarmentParts = N'" + FabricPart + "' ").getString() != "")
                {
                    FUNCT.msgWarning("Fabric Parts is Duplicate!"); txtFabricParts.Focus(); return;
                }
                else
                {
                    //Confirm Save
                    if (FUNCT.msgQuiz("Save Fabric Parts ? ") ==true)
                    {
                        sql = "INSERT INTO GarmentParts (GarmentParts, CreatedBy, CreatedDate) VALUES(N'" + FabricPart + "', '" + strCREATE + "', GETDATE())";
                        //Console.WriteLine(sql);
                        bool chkSave = DB.DBQuery(sql).runSQL();
                        if (chkSave == true)
                        {
                            FUNCT.msgInfo("Save Fabric Parts is Successufull.");
                            this.Close();
                        }
                    }
                }
            }
        }

        private void DEV01_M06_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (Application.OpenForms.OfType<DEV01>().Count() > 0)
            {
                var frmD01 = Application.OpenForms.OfType<DEV01>().FirstOrDefault();
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("SELECT OIDGParts AS ID, GarmentParts FROM GarmentParts ORDER BY GarmentParts");
                new ObjDE.setGridControl(frmD01.gcPart_Fabric, frmD01.gridView11, sbSQL).getData(false, false, true);

                string NewParts = DB.DBQuery("SELECT TOP(1) OIDGParts FROM GarmentParts WHERE GarmentParts=N'" + FabricPart + "'").getString();
                if (NewParts != "")
                {
                    if (strFBID != "")
                        strFBID += "," + NewParts;
                    else
                        strFBID = NewParts;
                }

                frmD01.gridView11.ClearSelection();
                strFBID = strFBID.Trim().Replace(" ", "");
                if (strFBID != "")
                {
                    DataTable dtGPart = (DataTable)frmD01.gcPart_Fabric.DataSource;
                    if (strFBID.IndexOf(',') != -1)
                    {
                        string[] ID = strFBID.Split(',');
                        if (ID.Length > 0)
                        {
                            foreach (string idPart in ID)
                            {
                                int iRow = 0;
                                foreach (DataRow drPart in dtGPart.Rows)
                                {
                                    string Part = drPart["ID"].ToString();
                                    if (idPart == Part)
                                    {
                                        frmD01.gridView11.SelectRow(iRow);
                                        break;
                                    }
                                    iRow++;
                                }
                            }
                        }

                    }
                    else
                    {
                        int iRow = 0;
                        foreach (DataRow drPart in dtGPart.Rows)
                        {
                            string Part = drPart["ID"].ToString();
                            if (strFBID == Part)
                            {
                                frmD01.gridView11.SelectRow(iRow);
                                break;
                            }
                            iRow++;
                        }
                    }

                }

            }
        }

        private void DEV01_M06_Load(object sender, EventArgs e)
        {
            new ObjDE.setDatabase(this.DB);
            //MessageBox.Show(DB.getCONNECTION_STRING());
            //MessageBox.Show(ObjDE.GlobalVar.DBC.getCONNECTION_STRING());
            if (Application.OpenForms.OfType<DEV01>().Count() > 0)
            {
                var frmD01 = Application.OpenForms.OfType<DEV01>().FirstOrDefault();
                int[] selectedRowHandles = frmD01.gridView11.GetSelectedRows();
                if (selectedRowHandles.Length > 0)
                {
                    int xLoop = 0;
                    frmD01.gridView11.FocusedRowHandle = selectedRowHandles[0];
                    for (int i = 0; i < selectedRowHandles.Length; i++)
                    {
                        string PartsID = frmD01.gridView11.GetRowCellDisplayText(selectedRowHandles[i], "ID");

                        if (xLoop > 0)
                        {
                            strFBID += ",";
                        }

                        strFBID += PartsID;

                        xLoop++;
                    }
                }
            }
        }
    }
}