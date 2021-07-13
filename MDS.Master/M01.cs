using System;
using System.Text;
using System.Windows.Forms;
using DevExpress.LookAndFeel;
using DevExpress.Utils.Extensions;
using DBConnect;
using System.Drawing;
using System.Data;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraEditors;
using TheepClass;

namespace MDS.Master
{
    public partial class M01 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private Functionality.Function FUNC = new Functionality.Function();
        public LogIn UserLogin { get; set; }
        public int Company { get; set; }
        public string ConnectionString { get; set; }
        string CONNECT_STRING = "";
        DatabaseConnect DBC;

        public M01()
        {
            InitializeComponent();
            UserLookAndFeel.Default.StyleChanged += MyStyleChanged;
        }

        private void MyStyleChanged(object sender, EventArgs e)
        {
            UserLookAndFeel userLookAndFeel = (UserLookAndFeel)sender;
            cUtility.SaveRegistry(@"Software\MDS", "SkinName", userLookAndFeel.SkinName);
            cUtility.SaveRegistry(@"Software\MDS", "SkinPalette", userLookAndFeel.ActiveSvgPaletteName);
        }
        private void XtraForm1_Load(object sender, EventArgs e)
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
            //****************************

            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT FunctionNo, FunctionName, Version ");
            sbSQL.Append("FROM FunctionList ");
            sbSQL.Append("WHERE (Status = 1) ");
            sbSQL.Append("ORDER BY FunctionNo ");
            new ObjDE.setGridControl(gcAbout, gvAbout, sbSQL).getData(false, false, true, true);
        }

        private void gvAbout_RowClick(object sender, RowClickEventArgs e)
        {
            if (gvAbout.IsFilterRow(e.RowHandle)) return;
        }
    }
}