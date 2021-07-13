using System;
using System.Text;
using DBConnect;
using System.Windows.Forms;
using DevExpress.LookAndFeel;
using DevExpress.Utils.Extensions;
using System.Drawing;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraEditors;
using DevExpress.XtraLayout.Utils;
using TheepClass;
using System.Globalization;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Data;
using DevExpress.Utils;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;

namespace MDS.Master
{
    public partial class M17 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private Functionality.Function FUNC = new Functionality.Function();
        public LogIn UserLogin { get; set; }
        public int Company { get; set; }
        public string ConnectionString { get; set; }
        string CONNECT_STRING = "";
        DatabaseConnect DBC;

        public M17()
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
            //*****************************

            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT TOP (1) ReadWriteStatus FROM FunctionAccess WHERE (OIDUser = '" + UserLogin.OIDUser + "') AND(FunctionNo = 'M17') ");
            int chkReadWrite = this.DBC.DBQuery(sbSQL).getInt();
            if (chkReadWrite == 0)
                ribbonPageGroup1.Visible = false;

            sbSQL.Clear();
            sbSQL.Append("SELECT FullName, OIDUSER FROM Users ORDER BY OIDUSER ");
            new ObjDE.setGridLookUpEdit(glueCREATE, sbSQL, "FullName", "OIDUSER").getData();

            glueCREATE.EditValue = UserLogin.OIDUser;

            LoadData();
            NewData();
        }

        private void LoadData()
        {
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT OIDPORT, PortCode, PortName, City, Country, CreatedBy, CreatedDate ");
            sbSQL.Append("FROM PortAndCity ");
            sbSQL.Append("ORDER BY Country, City, PortCode ");
            new ObjDE.setGridControl(gcPort, gvPort, sbSQL).getData(false, false, false, true);
            gvPort.Columns["OIDPORT"].Visible = false;
            gvPort.Columns["CreatedBy"].Visible = false;
            gvPort.Columns["CreatedDate"].Visible = false;

            DataTable dtCountries = new DataTable();
            dtCountries.Columns.Add("Country", typeof(System.String));
            int count = 0;
            foreach (string element in GetCountries())
            {
                count++;
                dtCountries.Rows.Add(element.Replace(" SAR", "").ToUpper().Trim());
            }

            slueCountry.Properties.DataSource = dtCountries;
            slueCountry.Properties.DisplayMember = "Country";
            slueCountry.Properties.ValueMember = "Country";
            
        }

        private void NewData()
        {
            lblStatus.Text = "* Add Port";
            lblStatus.ForeColor = Color.Green;

            txeID.Text = this.DBC.DBQuery("SELECT CASE WHEN ISNULL(MAX(OIDPORT), '') = '' THEN 1 ELSE MAX(OIDPORT) + 1 END AS NewNo FROM PortAndCity").getString();
            txeCode.Text = "";
            txeName.Text = "";
            txeCity.Text = "";
            slueCountry.EditValue = "";
            glueCREATE.EditValue = UserLogin.OIDUser;
            txeDATE.Text = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");

            txeCode.Focus();
        }

        private void bbiNew_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            LoadData();
            NewData();
        }

        private void bbiSave_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (txeCode.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input code.");
                txeCode.Focus();
            }
            else if (txeName.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input name.");
                txeName.Focus();
            }
            else if (txeCity.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input city.");
                txeCity.Focus();
            }
            else if (slueCountry.Text.Trim() == "")
            {
                FUNC.msgWarning("Please select country.");
                txeCity.Focus();
            }
            else
            {
                bool chkGMP = chkDuplicateNo();
                if (chkGMP == true)
                {
                    chkGMP = chkDuplicateName();
                    if (chkGMP == true)
                    {
                        if (FUNC.msgQuiz("Confirm save data ?") == true)
                        {
                            StringBuilder sbSQL = new StringBuilder();
                            string strCREATE = UserLogin.OIDUser.ToString() != "" ? UserLogin.OIDUser.ToString() : "0";

                            if (lblStatus.Text == "* Add Port")
                            {
                                sbSQL.Append("  INSERT INTO PortAndCity(PortCode, PortName, City, Country, CreatedBy, CreatedDate) ");
                                sbSQL.Append("  VALUES(N'" + txeCode.Text.Trim().Replace("'", "''") + "', N'" + txeName.Text.Trim().Replace("'", "''") + "', N'" + txeCity.Text.Trim().Replace("'", "''") + "', N'" + slueCountry.EditValue.ToString().Trim().Replace("'", "''") + "', '" + strCREATE + "', GETDATE()) ");
                            }
                            else if (lblStatus.Text == "* Edit Port")
                            {
                                sbSQL.Append("  UPDATE PortAndCity SET ");
                                sbSQL.Append("      PortCode=N'" + txeCode.Text.Trim().Replace("'", "''") + "', ");
                                sbSQL.Append("      PortName=N'" + txeName.Text.Trim().Replace("'", "''") + "', ");
                                sbSQL.Append("      City=N'" + txeCity.Text.Trim().Replace("'", "''") + "', ");
                                sbSQL.Append("      Country=N'" + slueCountry.EditValue.ToString().Trim().Replace("'", "''") + "' ");
                                sbSQL.Append("  WHERE(OIDPORT = '" + txeID.Text.Trim() + "') ");
                            }

                            //MessageBox.Show(sbSQL.ToString());
                            if (sbSQL.Length > 0)
                            {
                                try
                                {
                                    bool chkSAVE = this.DBC.DBQuery(sbSQL).runSQL();
                                    if (chkSAVE == true)
                                    {
                                        FUNC.msgInfo("Save complete.");
                                        bbiNew.PerformClick();
                                    }
                                }
                                catch (Exception)
                                { }
                            }
                        }
                    }
                }
            }
        }

        private void bbiExcel_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            string pathFile = new ObjSet.Folder(@"C:\MDS\Export\").GetPath() + "PortAndCityList_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
            gvPort.ExportToXlsx(pathFile);
            System.Diagnostics.Process.Start(pathFile);
        }


        private void bbiPrintPreview_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            gcPort.ShowPrintPreview();
        }

        private void bbiPrint_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            gcPort.Print();
        }

        private void txeCode_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeName.Focus();
            }
        }

        private void txeCode_Leave(object sender, EventArgs e)
        {
            if (txeCode.Text.Trim() != "")
            {
                txeCode.Text = txeCode.Text.ToUpper().Trim();
                bool chkDup = chkDuplicateNo();
                if (chkDup == true)
                {
                    txeName.Focus();
                }
                else
                {
                    txeCode.Text = "";
                    txeCode.Focus();
                    //FUNC.msgWarning("Duplicate code. !! Please Change.");

                }
            }
        }

        private void txeName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txeCity.Focus();
            }
        }

        private void txeName_Leave(object sender, EventArgs e)
        {
            if (txeName.Text.Trim() != "")
            {
                txeName.Text = txeName.Text.ToUpper().Trim();
                bool chkDup = chkDuplicateName();
                if (chkDup == true)
                {
                    txeCity.Focus();
                }
                else
                {
                    txeName.Text = "";
                    txeName.Focus();
                    //FUNC.msgWarning("Duplicate name. !! Please Change.");

                }
            }
        }

        private void txeCity_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                slueCountry.Focus();
            }
        }

        private void txeCity_Leave(object sender, EventArgs e)
        {

        }

        private bool chkDuplicateNo()
        {
            bool chkDup = true;
            if (txeCode.Text != "")
            {
                txeCode.Text = txeCode.Text.ToUpper().Trim();
                if (txeCode.Text.Trim() != "" && lblStatus.Text == "* Add Port")
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT TOP(1) PortCode FROM PortAndCity WHERE (PortCode = N'" + txeCode.Text.Trim().Trim().Replace("'", "''") + "') ");
                    if (this.DBC.DBQuery(sbSQL).getString() != "")
                    {
                        txeCode.Text = "";
                        txeCode.Focus();
                        chkDup = false;
                        FUNC.msgWarning("Duplicate code. !! Please Change.");
                    }
                }
                else if (txeCode.Text.Trim() != "" && lblStatus.Text == "* Edit Port")
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT TOP(1) OIDPORT ");
                    sbSQL.Append("FROM PortAndCity ");
                    sbSQL.Append("WHERE (PortCode = N'" + txeCode.Text.Trim().Trim().Replace("'", "''") + "') ");
                    string strCHK = this.DBC.DBQuery(sbSQL).getString();
                    if (strCHK != "" && strCHK != txeID.Text.Trim())
                    {
                        txeCode.Text = "";
                        txeCode.Focus();
                        chkDup = false;
                        FUNC.msgWarning("Duplicate code. !! Please Change.");
                    }
                }
            }
            return chkDup;
        }

        private bool chkDuplicateName()
        {
            bool chkDup = true;
            if (txeName.Text != "")
            {
                txeName.Text = txeName.Text.ToUpper().Trim();
                if (txeName.Text.Trim() != "" && lblStatus.Text == "* Add Port")
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT TOP(1) PortName FROM PortAndCity WHERE (PortName = N'" + txeName.Text.Trim().Replace("'", "''") + "') ");
                    if (this.DBC.DBQuery(sbSQL).getString() != "")
                    {
                        txeName.Text = "";
                        txeName.Focus();
                        chkDup = false;
                        FUNC.msgWarning("Duplicate name. !! Please Change.");
                    }
                }
                else if (txeName.Text.Trim() != "" && lblStatus.Text == "* Edit Port")
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("SELECT TOP(1) OIDPORT ");
                    sbSQL.Append("FROM PortAndCity ");
                    sbSQL.Append("WHERE (PortName = N'" + txeName.Text.Trim().Replace("'", "''") + "') ");
                    string strCHK = this.DBC.DBQuery(sbSQL).getString();
                    if (strCHK != "" && strCHK != txeID.Text.Trim())
                    {
                        txeName.Text = "";
                        txeName.Focus();
                        chkDup = false;
                        FUNC.msgWarning("Duplicate name. !! Please Change.");
                    }
                }
            }
            return chkDup;
        }

        //*********** REGION ************
        public class LocalesRetrievalException : Exception
        {
            public LocalesRetrievalException(string message)
                : base(message)
            {
            }
        }

        #region Windows API

        private delegate bool EnumLocalesProcExDelegate(
           [MarshalAs(UnmanagedType.LPWStr)] String lpLocaleString,
           LocaleType dwFlags, int lParam);

        [DllImport(@"kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern bool EnumSystemLocalesEx(EnumLocalesProcExDelegate pEnumProcEx,
           LocaleType dwFlags, int lParam, IntPtr lpReserved);

        private enum LocaleType : uint
        {
            LocaleAll = 0x00000000,             // Enumerate all named based locales
            LocaleWindows = 0x00000001,         // Shipped locales and/or replacements for them
            LocaleSupplemental = 0x00000002,    // Supplemental locales only
            LocaleAlternateSorts = 0x00000004,  // Alternate sort locales
            LocaleNeutralData = 0x00000010,     // Locales that are "neutral" (language only, region data is default)
            LocaleSpecificData = 0x00000020,    // Locales that contain language and region data
        }

        #endregion

        public enum CultureTypes : uint
        {
            SpecificCultures = LocaleType.LocaleSpecificData,
            NeutralCultures = LocaleType.LocaleNeutralData,
            AllCultures = LocaleType.LocaleWindows
        }

        public static List<CultureInfo> GetCultures(CultureTypes cultureTypes)
        {
            List<CultureInfo> cultures = new List<CultureInfo>();
            EnumLocalesProcExDelegate enumCallback = (locale, flags, lParam) =>
            {
                try
                {
                    cultures.Add(new CultureInfo(locale));
                }
                catch (CultureNotFoundException)
                {
                    // This culture is not supported by .NET (not happened so far)
                    // Must be ignored.
                }
                return true;
            };

            if (EnumSystemLocalesEx(enumCallback, (LocaleType)cultureTypes, 0, (IntPtr)0) == false)
            {
                int errorCode = Marshal.GetLastWin32Error();
                throw new LocalesRetrievalException("Win32 error " + errorCode + " while trying to get the Windows locales");
            }

            // Add the two neutral cultures that Windows misses 
            // (CultureInfo.GetCultures adds them also):
            if (cultureTypes == CultureTypes.NeutralCultures || cultureTypes == CultureTypes.AllCultures)
            {
                cultures.Add(new CultureInfo("en-US"));
                //cultures.Add(new CultureInfo("zh-CHS"));
                //cultures.Add(new CultureInfo("zh-CHT"));
            }

            return cultures;
        }

        public static List<string> GetCountries()
        {
            List<CultureInfo> cultures = GetCultures(CultureTypes.SpecificCultures);
            List<string> countries = new List<string>();
           
            foreach (CultureInfo culture in cultures)
            {
                RegionInfo region = new RegionInfo(culture.Name);

                if (!(countries.Contains(region.EnglishName)))
                {
                    countries.Add(region.EnglishName);
                }
            }
            countries.Sort();
            return countries;
        }


        //*********** END-REGION ********

        private void slueCountry_Popup(object sender, EventArgs e)
        {
            //(sender as SearchLookUpEdit).Properties.View.Columns["Country"].SortOrder = DevExpress.Data.ColumnSortOrder.Ascending;
        }

        private void gvPort_DoubleClick(object sender, EventArgs e)
        {
            DXMouseEventArgs ea = e as DXMouseEventArgs;
            GridView view = sender as GridView;
            GridHitInfo info = view.CalcHitInfo(ea.Location);
            if (info.InRow || info.InRowCell)
            {
                GridView gv = gvPort;
                lblStatus.Text = "* Edit Port";
                lblStatus.ForeColor = Color.Red;
                
                txeID.Text = gv.GetFocusedRowCellValue("OIDPORT").ToString();
                txeCode.Text = gv.GetFocusedRowCellValue("PortCode").ToString();
                txeName.Text = gv.GetFocusedRowCellValue("PortName").ToString();
                txeCity.Text = gv.GetFocusedRowCellValue("City").ToString();
                slueCountry.EditValue = gv.GetFocusedRowCellValue("Country").ToString();

                string CreatedBy = gv.GetFocusedRowCellValue("CreatedBy").ToString() == null ? "" : gv.GetFocusedRowCellValue("CreatedBy").ToString();
                glueCREATE.EditValue = CreatedBy;
                txeDATE.Text =  gv.GetFocusedRowCellValue("CreatedDate").ToString();
            }
        }

        private void gvPort_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator) e.Info.DisplayText = (e.RowHandle + 1).ToString();
            gvPort.IndicatorWidth = 40;
        }

        private void gvPort_RowClick(object sender, RowClickEventArgs e)
        {
            if (gvPort.IsFilterRow(e.RowHandle)) return;
        }
    }
}