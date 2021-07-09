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
using System.IO;
using DevExpress.Spreadsheet;
using System.Text.RegularExpressions;

namespace MDS.Master
{
    public partial class M16 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private Functionality.Function FUNC = new Functionality.Function();
        private const string xlsxPathFile = @"\\172.16.0.190\MDS_Project\MDS\ImportFile\Vessel\";
        string LongGestDays = "";
        private const int PORT_THAI = 8602;
        public LogIn UserLogin { get; set; }

        public string ConnectionString { get; set; }
        string CONNECT_STRING = "";
        DatabaseConnect DBC;

        public M16()
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
            sbSQL.Append("SELECT TOP (1) ReadWriteStatus FROM FunctionAccess WHERE (OIDUser = '" + UserLogin.OIDUser + "') AND(FunctionNo = 'M16') ");
            int chkReadWrite = this.DBC.DBQuery(sbSQL).getInt();
            if (chkReadWrite == 0)
                ribbonPageGroup1.Visible = false;

            sbSQL.Clear();
            sbSQL.Append("SELECT FullName, OIDUSER FROM Users ORDER BY OIDUSER ");
            new ObjDE.setGridLookUpEdit(glueCREATE, sbSQL, "FullName", "OIDUSER").getData();

            glueCREATE.EditValue = UserLogin.OIDUser;

            tabbedControlGroup1.SelectedTabPage = layoutControlGroup1;
            LoadData();
            NewData();
        }

        private void LoadData()
        {
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT Code AS VendorCode, Name AS Vendor, ShotName AS ShortName, OIDVEND AS ID ");
            sbSQL.Append("FROM Vendor ");
            sbSQL.Append("WHERE (VendorType = 5) ");
            sbSQL.Append("ORDER BY VendorCode ");
            new ObjDE.setSearchLookUpEdit(slueCarrier, sbSQL, "Vendor", "ID").getData();

            sbSQL.Clear();
            sbSQL.Append("SELECT PortCode, PortName, City, Country, PortCode AS ID ");
            sbSQL.Append("FROM PortAndCity ");
            sbSQL.Append("ORDER BY PortCode ");
            new ObjDE.setSearchLookUpEdit(slueFrom, sbSQL, "City", "ID").getData();
            new ObjDE.setSearchLookUpEdit(slueTo, sbSQL, "City", "ID").getData();

            //sbSQL.Clear();
            //sbSQL.Append("SELECT TOP(1) OIDPORT FROM PortAndCity WHERE (City = N'THAILAND') ");
            slueFrom.EditValue = PORT_THAI; //THAILAND DEFAULT
        }

        private void NewData()
        {
            bbiSave.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
            layoutBrowse.Visibility = LayoutVisibility.Always;
            rpgPrint.Visible = false;
            rpgExport.Visible = false;

            slueCarrier.EditValue = "";
            speYear.Value = Convert.ToInt32(DateTime.Now.ToString("yyyy"));

            slueFrom.EditValue = PORT_THAI;
            dteFileDate.EditValue = DateTime.Now;

            speTime.Value = 1;

            slueTo.EditValue = "";

            txeLimit.Text = "";
            txeStdDay.Text = "";
            txeCyCut.Text = "";
            txeEtdEta.Text = "";
            txeEtaWh.Text = "";
            rgStatus.EditValue = 1;

            txeFilePath.Text = "";

            glueCREATE.EditValue = UserLogin.OIDUser;
            txeDATE.Text = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");

            spsVessel.CloseCellEditor(DevExpress.XtraSpreadsheet.CellEditorEnterValueMode.Default);
            spsVessel.CreateNewDocument();

            gcVessel.DataSource = null;

            slueCarrier.Focus();
        }

        private void bbiNew_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            LoadData();
            NewData();
        }

        private void bbiSave_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (slueCarrier.Text.Trim() == "")
            {
                FUNC.msgWarning("Please select carrier.");
                slueCarrier.Focus();
            }
            else if (slueFrom.Text.Trim() == "")
            {
                FUNC.msgWarning("Please select departure.");
                slueFrom.Focus();
            }
            else if (slueTo.Text.Trim() == "")
            {
                FUNC.msgWarning("Please select destination.");
                slueTo.Focus();
            }
            else if (dteFileDate.Text.Trim() == "")
            {
                FUNC.msgWarning("Please select date.");
                dteFileDate.Focus();
            }
            else if (txeStdDay.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input Standard days(Longest).");
                txeStdDay.Focus();
            }
            else if (txeLimit.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input Limit of LCL>>FCL.");
                txeLimit.Focus();
            }
            else if (txeCyCut.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input CY Cut>>ETD(day).");
                txeCyCut.Focus();
            }
            else if (txeEtdEta.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input ETD>>ETA(day).");
                txeEtdEta.Focus();
            }
            else if (txeEtaWh.Text.Trim() == "")
            {
                FUNC.msgWarning("Please input ETA>>WH(day).");
                txeEtaWh.Focus();
            }
            else if (txeFilePath.Text.Trim() == "")
            {
                FUNC.msgWarning("Please select excel file.");
                txeFilePath.Focus();
            }
            else
            {
                if (FUNC.msgQuiz("The system will save data from the excel files only the first 2 sheets.\nConfirm save data ?") == true)
                {
                    StringBuilder sbSQL = new StringBuilder();
                    string strCREATE = UserLogin.OIDUser.ToString() != "" ? UserLogin.OIDUser.ToString() : "0";

                    string Status = "NULL";
                    if (rgStatus.SelectedIndex != -1)
                    {
                        Status = rgStatus.Properties.Items[rgStatus.SelectedIndex].Value.ToString();
                    }

                    string newPathFileName = "";
                    //**** SAVE EXCEL FILE ******
                    if (txeFilePath.Text.Trim() != "")
                    {
                        IWorkbook workbook = spsVessel.Document;
                        // Save the modified document to a stream.
                        System.IO.FileInfo fi = new System.IO.FileInfo(txeFilePath.Text);
                        string extn = fi.Extension;
                        string newFileName = slueCarrier.Text.Trim().Replace(" ", "_") + "-" + slueFrom.EditValue.ToString() + "-" + slueTo.EditValue.ToString() + "-" + speYear.Value.ToString() + "-" + speTime.Value.ToString() + extn;
                        newPathFileName = xlsxPathFile + newFileName;
                        using (FileStream stream = new FileStream(newPathFileName, FileMode.Create, FileAccess.ReadWrite))
                        {
                            workbook.SaveDocument(stream, DocumentFormat.Xlsx);
                        }
                    }

                    //*** Vessel ****
                    sbSQL.Append("IF NOT EXISTS(SELECT OIDVessel FROM Vessel WHERE OIDVend = '" + slueCarrier.EditValue.ToString() + "' AND FileYear = '" + speYear.Value.ToString() + "' AND OIDDeparturePort = '" + slueFrom.EditValue.ToString() + "' AND OIDDestinationPort = '" + slueTo.EditValue.ToString() + "' AND TimeOfDocument = '" + speTime.Value.ToString() + "') ");
                    sbSQL.Append(" BEGIN ");
                    sbSQL.Append("  INSERT INTO Vessel(OIDVend, TimeOfDocument, OIDDeparturePort, OIDDestinationPort, FileDate, FileYear, StdLongestDay, Status, LCLLimitOfCBM, DayOfCYCutToETD, DayOfETDtoETA, DayOfETAtoWH, PathFile, UpdatedBy, UpdatedDate) ");
                    sbSQL.Append("   VALUES('" + slueCarrier.EditValue.ToString() + "', '" + speTime.Value.ToString() + "', '" + slueFrom.EditValue.ToString() + "', '" + slueTo.EditValue.ToString() + "', '" + Convert.ToDateTime(dteFileDate.EditValue.ToString()).ToString("yyyy-MM-dd") + "', '" + speYear.Value.ToString() + "', '" + txeStdDay.Text.Trim() + "', " + Status + ", '" + txeLimit.Text.Trim() + "', '" + txeCyCut.Text.Trim() + "', '" + txeEtdEta.Text.Trim() + "', '" + txeEtaWh.Text.Trim() + "', N'" + newPathFileName + "', '" + strCREATE + "', GETDATE())  ");
                    sbSQL.Append(" END  ");
                    sbSQL.Append("ELSE ");
                    sbSQL.Append(" BEGIN ");
                    sbSQL.Append("  UPDATE Vessel SET ");
                    sbSQL.Append("     OIDDeparturePort='" + slueFrom.EditValue.ToString() + "', ");
                    sbSQL.Append("     OIDDestinationPort='" + slueTo.EditValue.ToString() + "', ");
                    sbSQL.Append("     FileDate='" + Convert.ToDateTime(dteFileDate.EditValue.ToString()).ToString("yyyy-MM-dd") + "', ");
                    sbSQL.Append("     StdLongestDay='" + txeStdDay.Text.Trim() + "', ");
                    sbSQL.Append("     Status=" + Status + ", ");
                    sbSQL.Append("     LCLLimitOfCBM='" + txeLimit.Text.Trim() + "', ");
                    sbSQL.Append("     DayOfCYCutToETD='" + txeCyCut.Text.Trim() + "', ");
                    sbSQL.Append("     DayOfETDtoETA='" + txeEtdEta.Text.Trim() + "', ");
                    sbSQL.Append("     DayOfETAtoWH='" + txeEtaWh.Text.Trim() + "', ");
                    sbSQL.Append("     PathFile=N'" + newPathFileName + "' ");
                    sbSQL.Append("  WHERE (OIDVend = '" + slueCarrier.EditValue.ToString() + "') AND (FileYear = '" + speYear.Value.ToString() + "') AND (OIDDeparturePort = '" + slueFrom.EditValue.ToString() + "') AND (OIDDestinationPort = '" + slueTo.EditValue.ToString() + "') AND (TimeOfDocument = '" + speTime.Value.ToString() + "') ");
                    sbSQL.Append(" END ");

                    sbSQL.Append("UPDATE Vessel SET ");
                    sbSQL.Append("  Status = 0 ");
                    sbSQL.Append("WHERE (OIDVend = '" + slueCarrier.EditValue.ToString() + "') ");
                    sbSQL.Append("AND (FileYear = '" + speYear.Value.ToString() + "') ");
                    sbSQL.Append("AND (OIDDeparturePort = '" + slueFrom.EditValue.ToString() + "') ");
                    sbSQL.Append("AND (OIDDestinationPort = '" + slueTo.EditValue.ToString() + "') ");
                    sbSQL.Append("AND (TimeOfDocument < " + speTime.Value.ToString() + ") ");
                   // MessageBox.Show(sbSQL.ToString());
                    if (sbSQL.Length > 0)
                    {
                        try
                        {
                            bool chkSAVE = this.DBC.DBQuery(sbSQL).runSQL();
                            if (chkSAVE == true)
                            {
                                sbSQL.Clear();
                                sbSQL.Append("SELECT MAX(OIDVessel) AS ID FROM Vessel WHERE (OIDVend = '" + slueCarrier.EditValue.ToString() + "') ");
                                string OIDVessel = this.DBC.DBQuery(sbSQL).getString();

                                sbSQL.Clear();
                                sbSQL.Append("DELETE FROM VesselDetail WHERE (OIDVessel = '" + OIDVessel + "')  ");
                                if (OIDVessel != "")
                                {
                                    //*** VesselDetail ****
                                    IWorkbook workbook = spsVessel.Document;
                                    int TTSHEET = workbook.Worksheets.Count;
                                    if (TTSHEET > 2)
                                    {
                                        TTSHEET = 2;
                                    }
                                    for (int Sheet = 0; Sheet < TTSHEET; Sheet++)
                                    {
                                        Worksheet WSHEET = workbook.Worksheets[Sheet];
                                        for (int i = 4; i < WSHEET.GetDataRange().RowCount; i++)
                                        {
                                            string Vessel = WSHEET.Rows[i][0].DisplayText.ToString();
                                            string Voy = WSHEET.Rows[i][1].DisplayText.ToString();
                                            string TSorDirect = WSHEET.Rows[i][2].DisplayText.ToString();
                                            string TSPort = WSHEET.Rows[i][3].DisplayText.ToString();
                                            string VesselType = WSHEET.Rows[i][4].DisplayText.ToString();
                                            string Carrier = WSHEET.Rows[i][5].DisplayText.ToString();
                                            string CFSCutDate = WSHEET.Rows[i][6].DisplayText.ToString();
                                            CFSCutDate = CFSCutDate != "" ? "'" + Convert.ToDateTime(CFSCutDate).ToString("yyyy-MM-dd HH:mm:ss") + "'" : "NULL";
                                            string CFSCutDay = WSHEET.Rows[i][7].DisplayText.ToString();
                                            string CFSCutTime = WSHEET.Rows[i][8].DisplayText.ToString();
                                            CFSCutTime = CFSCutTime != "" ? "'" + Convert.ToDateTime(CFSCutTime).ToString("HH:mm:ss") + "'" : "NULL";
                                            string ETDDate = WSHEET.Rows[i][9].DisplayText.ToString();
                                            ETDDate = ETDDate != "" ? "'" + Convert.ToDateTime(ETDDate).ToString("yyyy-MM-dd HH:mm:ss") + "'" : "NULL";
                                            string ETDDay = WSHEET.Rows[i][10].DisplayText.ToString();
                                            string ETADate = WSHEET.Rows[i][11].DisplayText.ToString();
                                            ETADate = ETADate != "" ? "'" + Convert.ToDateTime(ETADate).ToString("yyyy-MM-dd HH:mm:ss") + "'" : "NULL";
                                            string ETADay = WSHEET.Rows[i][12].DisplayText.ToString();
                                            string ETAWHDate = WSHEET.Rows[i][13].DisplayText.ToString();
                                            ETAWHDate = ETAWHDate != "" ? "'" + Convert.ToDateTime(ETAWHDate).ToString("yyyy-MM-dd HH:mm:ss") + "'" : "NULL";
                                            string ETAWHDay = WSHEET.Rows[i][14].DisplayText.ToString();
                                            string ETDtoWHDays = WSHEET.Rows[i][15].DisplayText.ToString();
                                            string ETAtoWHDays = WSHEET.Rows[i][16].DisplayText.ToString();
                                            string Priority = WSHEET.Rows[i][17].DisplayText.ToString();
                                            string Remarks = WSHEET.Rows[i][18].DisplayText.ToString();

                                            if (Vessel != "")
                                            {
                                                sbSQL.Append("INSERT INTO VesselDetail(OIDVessel, Vessel, Voy, TSorDirect, TSPort, VesselType, Carrier, CFSCutDate, CFSCutDay, CFSCutTime, ETDDate, ETDDay, ETADate, ETADay, ETAWHDate, ETAWHDay, ETDtoWHDays, ETAtoWHDays, Priority, Remarks) ");
                                                sbSQL.Append(" VALUES('" + OIDVessel + "', N'" + Vessel + "', N'" + Voy + "', N'" + TSorDirect + "', N'" + TSPort + "', N'" + VesselType + "', N'" + Carrier + "', " + CFSCutDate + ", N'" + CFSCutDay + "', " + CFSCutTime + ", " + ETDDate + ", N'" + ETDDay + "', " + ETADate + ", N'" + ETADay + "', " + ETAWHDate + ", N'" + ETAWHDay + "', '" + ETDtoWHDays + "', '" + ETAtoWHDays + "', N'" + Priority + "', N'" + Remarks + "')  ");
                                            }
                                            //MessageBox.Show(sbSQL.ToString());
                                        }
                                    }

                                    if (sbSQL.Length > 0)
                                    {
                                        try
                                        {
                                            chkSAVE = this.DBC.DBQuery(sbSQL).runSQL();
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
                        catch (Exception)
                        { }
                    }
                    
                }
            }
          
        }

        private void bbiExcel_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (tabbedControlGroup1.SelectedTabPage == layoutControlGroup1) //Excel
            {
                IWorkbook workbook = spsVessel.Document;
                string pathFile = new ObjSet.Folder(@"C:\MDS\Export\").GetPath() + slueCarrier.Text.Trim().Replace(" ", "_") + "-" + speYear.Value.ToString() + "-" + speTime.Value.ToString() + ".xlsx";
                using (FileStream stream = new FileStream(pathFile, FileMode.Create, FileAccess.ReadWrite))
                {
                    workbook.SaveDocument(stream, DocumentFormat.Xlsx);
                    System.Diagnostics.Process.Start(pathFile);
                }
                
            }
            else //GridControl
            {
                string pathFile = new ObjSet.Folder(@"C:\MDS\Export\").GetPath() + "VesselList_" + slueCarrier.Text.Trim().Replace(" ", "_") + "-" + speYear.Value.ToString() + "-" + speTime.Value.ToString() + "-" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
                gvVessel.ExportToXlsx(pathFile);
                System.Diagnostics.Process.Start(pathFile);
            }

        }


        private void bbiPrintPreview_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (tabbedControlGroup1.SelectedTabPage == layoutControlGroup1) //Excel
            {
                spsVessel.ShowPrintPreview();
            }
            else //GridControl
            {
                gcVessel.ShowPrintPreview();
            }
        }

        private void bbiPrint_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (tabbedControlGroup1.SelectedTabPage == layoutControlGroup1) //Excel
            {
                spsVessel.Print();
            }
            else //GridControl
            {
                gcVessel.Print();
            }
        }


        private void sbOpenFile_Click(object sender, EventArgs e)
        {
            if (slueCarrier.Text.Trim() == "")
            {
                FUNC.msgWarning("Please select carrier before select file.");
                slueCarrier.Focus();
            }
            else if (slueFrom.Text.Trim() == "")
            {
                FUNC.msgWarning("Please select departure before select file.");
                slueFrom.Focus();
            }
            else if (slueTo.Text.Trim() == "")
            {
                FUNC.msgWarning("Please select destination before select file.");
                slueTo.Focus();
            }
            else
            {
                xtraOpenFileDialog1.Filter = "Excel Files|*.xlsx";
                xtraOpenFileDialog1.FileName = "";
                xtraOpenFileDialog1.Title = "Select Excel File";

                if (xtraOpenFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    txeFilePath.Text = xtraOpenFileDialog1.FileName;

                    IWorkbook workbook = spsVessel.Document;

                    try
                    {
                        //workbook.LoadDocument(txeFilePath.Text, DocumentFormat.OpenXml);
                        // Load a workbook from a stream.
                        using (FileStream stream = new FileStream(txeFilePath.Text, FileMode.Open))
                        {
                            // workbook.CalculateFull();
                            workbook.LoadDocument(stream, DocumentFormat.Xlsx);
                            workbook.Worksheets.ActiveWorksheet = workbook.Worksheets[0];

                            //*** Delete sheet > 2
                            //if (workbook.Worksheets.Count > 2)
                            //{
                            //    for (int i = workbook.Worksheets.Count - 1; i > 1; i--)
                            //    {
                            //        workbook.Worksheets.RemoveAt(i);
                            //    }
                            //}

                            LoadSheetHead(workbook.Worksheets[0]);
                        }
                    }
                    catch (Exception)
                    {
                        FUNC.msgWarning("Please close excel file before import.");
                        txeFilePath.Text = "";
                    }

                    //// Access a collection of worksheets.
                    //WorksheetCollection worksheets = workbook.Worksheets;

                    // Access a worksheet by its index.
                    //Worksheet worksheet2 = workbook.Worksheets[1];

                    //// Access a worksheet by its name.
                    //Worksheet worksheet2 = workbook.Worksheets["Sheet2"];

                    // txeLimit.Text = worksheet2.Rows[0]["B"].DisplayText;

                }
            }
        }

        private void spsVessel_SelectionChanged(object sender, EventArgs e)
        {
            Worksheet worksheet = spsVessel.ActiveWorksheet;
            callSheetActive(worksheet);
        }

        private void spsVessel_ActiveSheetChanged(object sender, ActiveSheetChangedEventArgs e)
        {
            //Worksheet worksheet = spsVessel.ActiveWorksheet;
            //callSheetActive(worksheet);
        }

        private void callSheetActive(Worksheet wsActive)
        {
            //if (this.nowSheet != wsActive.Name)
            //{
            //    LoadSheetHead(wsActive);


            //}
            //this.nowSheet = wsActive.Name;
        }


        private void LoadSheetHead(Worksheet wsActive)
        {
            //Set null
            txeLimit.Text = "";
            txeStdDay.Text = "";
            txeCyCut.Text = "";
            txeEtdEta.Text = "";
            txeEtaWh.Text = "";

            //Set Value
            //**** Limit of LCL >> FCL *****
            string LimitOfCBM = "";
            if (wsActive.Rows[1]["E"].DisplayText != "")
            {
                LimitOfCBM = wsActive.Rows[1]["E"].DisplayText;
            }
            else if (wsActive.Rows[1]["F"].DisplayText != "")
            {
                LimitOfCBM = wsActive.Rows[1]["F"].DisplayText;
            }
            else if (wsActive.Rows[1]["G"].DisplayText != "")
            {
                LimitOfCBM = wsActive.Rows[1]["G"].DisplayText;
            }
            LimitOfCBM = Regex.Match(LimitOfCBM, @"\d+([,\.]\d+)?").Value;
            txeLimit.Text = LimitOfCBM;

            //**** Standard days (Longest) ****
            string StdDay = wsActive.Rows[1]["P"].DisplayText.Trim();
            txeStdDay.Text = StdDay;

            //**** Longest Days (Check) ****
            string LongGestDay = wsActive.Rows[2]["P"].DisplayText;
            this.LongGestDays = LongGestDay;

            //**** CY Cut >> ETD ****
            string CyCut = wsActive.Rows[2]["G"].DisplayText;
            CyCut = Regex.Match(CyCut, @"\d+([,\.]\d+)?").Value;
            txeCyCut.Text = CyCut;

            //**** ETD >> ETA ****
            string EtdEta = wsActive.Rows[2]["J"].DisplayText;
            EtdEta = Regex.Match(EtdEta, @"\d+([,\.]\d+)?").Value;
            txeEtdEta.Text = EtdEta;

            //**** ETD >> ETA ****
            string EtaWh = wsActive.Rows[2]["Q"].DisplayText;
            //EtaWh = Regex.Match(EtaWh, @"\d+([,\.]\d+)?").Value;
            EtaWh = "0";
            txeEtaWh.Text = EtaWh;
        }

        private void M16_Shown(object sender, EventArgs e)
        {
            //StringBuilder sbSQL = new StringBuilder();
            //sbSQL.Append("SELECT TOP(1) OIDPORT FROM PortAndCity WHERE (City = N'Bangkok') ");
            slueFrom.EditValue = PORT_THAI;
        }

        private void slueCarrier_EditValueChanged(object sender, EventArgs e)
        {
            speTime.Value = 1;
            if (slueCarrier.Text.Trim() != "")
            {
                findTime();
            }
        }

        private void findTime()
        {
            if (slueCarrier.Text.Trim() != "" && speYear.Value.ToString() != "" && slueFrom.Text.Trim() != "" && slueTo.Text.Trim() != "")
            {
                int strTime = this.DBC.DBQuery("SELECT MAX(TimeOfDocument) + 1 AS NewNo FROM Vessel WHERE (OIDVend = '" + slueCarrier.EditValue.ToString() + "') AND (FileYear = '" + speYear.Value.ToString() + "') AND (OIDDeparturePort = '" + slueFrom.EditValue.ToString() + "') AND (OIDDestinationPort = '" + slueTo.EditValue.ToString() + "') ").getInt();
                if (strTime == 0)
                {
                    strTime = 1;
                }
                speTime.Value = strTime;
            }
            else
            {
                speTime.Value = 1;
            }

            LoadVessel();
        }

        private void spsVessel_CellEndEdit(object sender, DevExpress.XtraSpreadsheet.SpreadsheetCellValidatingEventArgs e)
        {
            
        }

        private void spsVessel_CellValueChanged(object sender, DevExpress.XtraSpreadsheet.SpreadsheetCellEventArgs e)
        {

        }

        private void speYear_EditValueChanged(object sender, EventArgs e)
        {
            speTime.Value = 1;
            if (slueCarrier.Text.Trim() != "" && slueFrom.Text.Trim() != "" && slueTo.Text.Trim() != "")
            {
                findTime();
            }
        }

        private void LoadVessel()
        {
            bbiSave.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
            layoutBrowse.Visibility = LayoutVisibility.Always;
            rpgPrint.Visible = false;
            rpgExport.Visible = false;

            //SET DEFALUT
            StringBuilder sbSQL = new StringBuilder();
            //sbSQL.Append("SELECT TOP(1) OIDPORT FROM PortAndCity WHERE (City = N'Bangkok') ");
            //slueFrom.EditValue = this.DBC.DBQuery(sbSQL).getInt();

            dteFileDate.EditValue = DateTime.Now;
            //slueTo.EditValue = "";

            txeLimit.Text = "";
            txeStdDay.Text = "";
            txeCyCut.Text = "";
            txeEtdEta.Text = "";
            txeEtaWh.Text = "";
            rgStatus.EditValue = 1;

            txeFilePath.Text = "";

            txeDATE.Text = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");

            spsVessel.CloseCellEditor(DevExpress.XtraSpreadsheet.CellEditorEnterValueMode.Default);
            spsVessel.CreateNewDocument();

            gcVessel.DataSource = null;

            int Defalut_Time = Convert.ToInt32(speTime.Value);
            //LOAD DATA
            string CarrierID = slueCarrier.Text.Trim() != "" ? slueCarrier.EditValue.ToString() : "";
            string Departure = slueFrom.Text.Trim() != "" ? slueFrom.EditValue.ToString() : "";
            string Destination = slueTo.Text.Trim() != "" ? slueTo.EditValue.ToString() : "";

            if (CarrierID != "" && Departure != "" && Destination != "")
            {
                sbSQL.Clear();
                sbSQL.Append("SELECT OIDVessel, OIDVend, TimeOfDocument, OIDDeparturePort, OIDDestinationPort, FileDate, FileYear, StdLongestDay, Status, LCLLimitOfCBM, DayOfCYCutToETD, DayOfETDtoETA, DayOfETAtoWH, PathFile, UpdatedBy, UpdatedDate ");
                sbSQL.Append("FROM Vessel ");
                sbSQL.Append("WHERE(OIDVend = '" + CarrierID + "') AND (FileYear = '" + speYear.Value.ToString() + "') AND (OIDDeparturePort = '" + Departure + "') AND (OIDDestinationPort = '" + Destination + "') AND (TimeOfDocument = '" + speTime.Value.ToString() + "') ");
                string[] arrVessel = this.DBC.DBQuery(sbSQL).getMultipleValue();
                if (arrVessel.Length > 0)
                {
                    bbiSave.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
                    rpgPrint.Visible = true;
                    rpgExport.Visible = true;
                    layoutBrowse.Visibility = LayoutVisibility.Never;

                    string OIDVessel = arrVessel[0];
                    slueFrom.EditValue = Convert.ToInt32(arrVessel[3]);
                    slueTo.EditValue = Convert.ToInt32(arrVessel[4]);
                    dteFileDate.EditValue = Convert.ToDateTime(arrVessel[5]);
                    txeStdDay.Text = arrVessel[7];
                    rgStatus.EditValue = Convert.ToInt32(arrVessel[8]);
                    txeLimit.Text = arrVessel[9];
                    txeCyCut.Text = arrVessel[10];
                    txeEtdEta.Text = arrVessel[11];
                    txeEtaWh.Text = arrVessel[12];

                    txeFilePath.Text = arrVessel[13];
                    glueCREATE.EditValue = arrVessel[14];
                    txeDATE.Text = Convert.ToDateTime(arrVessel[15]).ToString("dd/MM/yyyy HH:mm:ss");

                    if (txeFilePath.Text.Trim() != "")
                    {
                        //OPEN EXCEL
                        try
                        {
                            IWorkbook workbook = spsVessel.Document;
                            using (FileStream stream = new FileStream(txeFilePath.Text.Trim(), FileMode.Open))
                            {
                                workbook.LoadDocument(stream, DocumentFormat.Xlsx);
                                workbook.Worksheets.ActiveWorksheet = workbook.Worksheets[0];
                                LoadSheetHead(workbook.Worksheets[0]);
                            }
                        }
                        catch (Exception)
                        { }
                    }

                    //Load to GridControl
                    sbSQL.Clear();
                    sbSQL.Append("SELECT VSD.OIDVesselDT AS ID, VSD.OIDVessel AS [Vessel ID], VD.Code AS [Vendor Code], VD.Name AS [Vendor Name], VD.ShotName AS [Vendor Short Name], VSD.Vessel, VSD.Voy, VSD.TSorDirect AS [T/S or Direct], ");
                    sbSQL.Append("       VSD.TSPort AS[T / S Port], VSD.VesselType AS[FCL / LCL], VSD.Carrier, VSD.CFSCutDate AS[CY - Cut Date], VSD.CFSCutDay AS[CY - Cut Day], VSD.CFSCutTime AS[CY - Cut Time], VSD.ETDDate AS[ETD Date], ");
                    sbSQL.Append("       VSD.ETDDay AS[ETD Day], VSD.ETADate AS[ETA Date], VSD.ETADay AS[ETA Day], VSD.ETAWHDate AS[ETA - WH Date], VSD.ETAWHDay AS[ETA - WH Day], VSD.ETDtoWHDays AS[Days ETD - WH], ");
                    sbSQL.Append("       VSD.ETAtoWHDays AS[Days ETA - WH], VSD.Priority, VSD.Remarks ");
                    sbSQL.Append("FROM   VesselDetail AS VSD INNER JOIN ");
                    sbSQL.Append("       Vessel AS VS ON VSD.OIDVessel = VS.OIDVessel LEFT OUTER JOIN ");
                    sbSQL.Append("       Vendor AS VD ON VS.OIDVend = VD.OIDVEND ");
                    sbSQL.Append("WHERE (VSD.OIDVessel = '" + OIDVessel + "') ");
                    sbSQL.Append("ORDER BY[FCL / LCL], [CY - Cut Date] ");
                    new ObjDE.setGridControl(gcVessel, gvVessel, sbSQL).getDataShowOrder(false, false, false, true);
                    gvVessel.Columns["NO"].Visible = false;
                    gvVessel.Columns["ID"].Visible = false;
                    gvVessel.Columns["Vessel ID"].Visible = false;
                }
            }
        }

        private void speTime_EditValueChanged(object sender, EventArgs e)
        {
            int strTime = this.DBC.DBQuery("SELECT MAX(TimeOfDocument) + 1 AS NewNo FROM Vessel WHERE (OIDVend = '" + slueCarrier.EditValue.ToString() + "') AND (FileYear = '" + speYear.Value.ToString() + "') AND (OIDDeparturePort = '" + slueFrom.EditValue.ToString() + "') AND (OIDDestinationPort = '" + slueTo.EditValue.ToString() + "') ").getInt();
            if (strTime == 0)
            {
                strTime = 1;
            }

            if (speTime.Value > strTime)
            {
                speTime.Value = strTime;
            }

            LoadVessel();
        }

        private void ribbonControl_Click(object sender, EventArgs e)
        {

        }

        private void gvVessel_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator) e.Info.DisplayText = (e.RowHandle + 1).ToString();
            gvVessel.IndicatorWidth = 40;
        }

        private void slueFrom_EditValueChanged(object sender, EventArgs e)
        {
            if (slueCarrier.Text.Trim() != "" && slueFrom.Text.Trim() != "" && slueTo.Text.Trim() != "")
            {
                findTime();
            }
            slueTo.Focus();
        }

        private void slueTo_EditValueChanged(object sender, EventArgs e)
        {
            if (slueCarrier.Text.Trim() != "" && slueFrom.Text.Trim() != "" && slueTo.Text.Trim() != "")
            {
                findTime();
            }
            txeFilePath.Focus();
        }
    }
}