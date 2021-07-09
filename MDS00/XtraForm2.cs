using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Data;
using System.Diagnostics;
using DevExpress.LookAndFeel;
using DevExpress.XtraBars.Navigation;
using TheepClass;

namespace MDS00
{
    public partial class XtraForm2 : DevExpress.XtraEditors.XtraForm
    {
        cDatabase db;

        public XtraForm2()
        {
            InitializeComponent();
            localConfig = new IniFile("Config.ini");
            UserLookAndFeel.Default.SetSkinStyle(localConfig.Read("SkinName", "DevExpress"), localConfig.Read("SkinPalette", "DevExpress"));
            shareConfig = new IniFile(@"\\172.16.0.190\mds_project\MDS\FileConfig\Configue.ini");
            db = new cDatabase("Server=" + shareConfig.Read("Server","Connectionstring")+
                ";uid="+ shareConfig.Read("Uid","ConnectionString")+
                ";pwd="+ shareConfig.Read("Pwd","ConnectionString")+
                ";database="+ shareConfig.Read("Database","ConnectionString"));
            InitAccordionControl();
        }

        private IniFile shareConfig;
        private IniFile localConfig;

        private const int SW_MAXIMIZE = 3;
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        private void RunProcess(string pathFile)
        {
            var processName = pathFile.Reverse().ToString().Split('\\')[0].Reverse().ToString();
            var aryProcess = Process.GetProcesses();
            if (aryProcess.Any(x=>x.ProcessName==processName))
            {
                MessageBox.Show(processName + " is running.");
                return;
            }
            else
            {
                Process p = Process.Start(pathFile);
                SetForegroundWindow(p.MainWindowHandle);
                ShowWindow(p.MainWindowHandle,SW_MAXIMIZE);
            }
        }
        private void CreateSplashScreen(string processName)
        {
            DevExpress.XtraSplashScreen.SplashScreenManager.ShowSkinSplashScreen(
                logoImage: null,
                title: "MDS",
                subtitle: "Merchandise and Development System",
                footer: "Copyright © 2020-2021 IT Integration Team",
                loading: "Starting..."+processName,
                parentForm: this,
                useFadeIn: true,
                useFadeOut: true,
                throwExceptionIfAlreadyOpened: true,
                startPos: DevExpress.XtraSplashScreen.SplashFormStartPosition.Default,
                location: default
                );
        }
        private void CloseSplashScreen()
        {
            DevExpress.XtraSplashScreen.SplashScreenManager.CloseForm();
        }
        private void InitAccordionControl()
        {
            string strSQL = "SELECT FunctionNo,FunctionName,Version,"+
                "CASE WHEN GroupFunction=0 THEN 'Master'"+
                "     WHEN GroupFunction = 1 THEN 'Development'"+
                "     WHEN GroupFunction = 2 THEN 'MPS: Master Production Schedule'"+
                "     WHEN GroupFunction = 3 THEN 'MRP: Material Resource Planning'"+
                "     WHEN GroupFunction = 4 THEN 'Shipment'"+
                "     WHEN GroupFunction = 5 THEN 'EXIMs'"+
                "     WHEN GroupFunction = 6 THEN 'Administrator'"+
                " END AS GroupFunction,PathFile "+
                "FROM FunctionList WHERE Status = 1 ORDER BY GroupFunction,FunctionNo";
            DataTable dt = db.GetDataTable(strSQL);
            if (dt == null || dt.Rows.Count == 0) return;
            DataView view = new DataView(dt);
            DataTable distinctGroupFunction = view.ToTable(true,"GroupFunction");
            accordionControl1.BeginUpdate();
            foreach (DataRow dr in distinctGroupFunction.Rows)
            {
                AccordionControlElement acRoot = new AccordionControlElement();
                acRoot.Text = dr["GroupFunction"].ToString();
                var itemInGroup = dt.Select("GroupFunction='"+dr["GroupFunction"].ToString()+"'");
                foreach (DataRow item in itemInGroup)
                {
                    AccordionControlElement acItem = new AccordionControlElement();
                    acItem.Text = item["FunctionName"].ToString();
                    acItem.Hint = item["FunctionNo"].ToString() + " Version " + item["Version"].ToString();
                    acItem.Tag =  item["PathFile"].ToString();
                    acItem.Style = ElementStyle.Item;
                    acItem.ImageOptions.Image =(Image)Resource1.ResourceManager.GetObject(item["FunctionNo"].ToString());
                    acRoot.Elements.Add(acItem);
                }
                accordionControl1.Elements.Add(acRoot);
            }
            accordionControl1.ElementClick += AccordionControl1_ElementClick;
            accordionControl1.EndUpdate();
            

        }

        private void AccordionControl1_ElementClick(object sender, ElementClickEventArgs e)
        {
            if (e.Element.Style == ElementStyle.Group) return;
            try
            {
                CreateSplashScreen(e.Element.Hint);
                RunProcess(e.Element.Tag.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            CloseSplashScreen();
        }

        //private void ace_Click(object sender, EventArgs e)
        //{
        //    var obj = sender as AccordionControlElement;
        //    try
        //    {
        //        RunProcess(obj.Tag.ToString());
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}
        //private void aceM01_Click(object sender, EventArgs e)
        //{
        //    var obj = sender as AccordionControlElement;
        //    try
        //    {
        //        RunProcess(obj.Tag.ToString());
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}
        //private void aceM02_Click(object sender, EventArgs e)
        //{
        //    var obj = sender as AccordionControlElement;
        //    try
        //    {
        //        RunProcess(obj.Tag.ToString());
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}
        //private void aceM03_Click(object sender, EventArgs e)
        //{
        //    var obj = sender as AccordionControlElement;
        //    try
        //    {
        //        RunProcess(obj.Tag.ToString());
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}
        //private void aceM04_Click(object sender, EventArgs e)
        //{
        //    var obj = sender as AccordionControlElement;
        //    try
        //    {
        //        RunProcess(obj.Tag.ToString());
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}
        //private void aceM05_Click(object sender, EventArgs e)
        //{
        //    var obj = sender as AccordionControlElement;
        //    try
        //    {
        //        RunProcess(obj.Tag.ToString());
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}
        //private void aceM06_Click(object sender, EventArgs e)
        //{
        //    var obj = sender as AccordionControlElement;
        //    try
        //    {
        //        RunProcess(obj.Tag.ToString());
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}
        //private void aceM07_Click(object sender, EventArgs e)
        //{
            
        //    var obj = sender as AccordionControlElement;
        //    try
        //    {
        //        RunProcess(obj.Tag.ToString());
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
            
        //}
        //private void aceM08_Click(object sender, EventArgs e)
        //{
        //    var obj = sender as AccordionControlElement;
        //    try
        //    {
        //        RunProcess(obj.Tag.ToString());
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}
        //private void aceM09_Click(object sender, EventArgs e)
        //{
        //    var obj = sender as AccordionControlElement;
        //    try
        //    {
        //        RunProcess(obj.Tag.ToString());
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}
        //private void aceM10_Click(object sender, EventArgs e)
        //{
        //    var obj = sender as AccordionControlElement;
        //    try
        //    {
        //        RunProcess(obj.Tag.ToString());
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}
        //private void aceM11_Click(object sender, EventArgs e)
        //{
        //    var obj = sender as AccordionControlElement;
        //    try
        //    {
        //        RunProcess(obj.Tag.ToString());
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}
        //private void aceM12_Click(object sender, EventArgs e)
        //{
        //    var obj = sender as AccordionControlElement;
        //    try
        //    {
        //        RunProcess(obj.Tag.ToString());
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}
        //private void aceM13_Click(object sender, EventArgs e)
        //{
        //    var obj = sender as AccordionControlElement;
        //    try
        //    {
        //        RunProcess(obj.Tag.ToString());
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}
        //private void aceM14_Click(object sender, EventArgs e)
        //{
        //    var obj = sender as AccordionControlElement;
        //    try
        //    {
        //        RunProcess(obj.Tag.ToString());
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}
        //private void aceM15_Click(object sender, EventArgs e)
        //{
        //    var obj = sender as AccordionControlElement;
        //    try
        //    {
        //        RunProcess(obj.Tag.ToString());
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}
        //private void aceDEV01_Click(object sender, EventArgs e)
        //{
        //    var obj = sender as AccordionControlElement;
        //    try
        //    {
        //        RunProcess(obj.Tag.ToString());
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}
        //private void aceDEV02_Click(object sender, EventArgs e)
        //{
        //    var obj = sender as AccordionControlElement;
        //    try
        //    {
        //        RunProcess(obj.Tag.ToString());
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}
        //private void aceDEV03_Click(object sender, EventArgs e)
        //{
        //    var obj = sender as AccordionControlElement;
        //    try
        //    {
        //        RunProcess(obj.Tag.ToString());
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}
        //private void aceF01_Click(object sender, EventArgs e)
        //{
        //    var obj = sender as AccordionControlElement;
        //    try
        //    {
        //        RunProcess(obj.Tag.ToString());
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}
        //private void aceF02_Click(object sender, EventArgs e)
        //{
        //    var obj = sender as AccordionControlElement;
        //    try
        //    {
        //        RunProcess(obj.Tag.ToString());
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}
        //private void aceF03_Click(object sender, EventArgs e)
        //{
        //    var obj = sender as AccordionControlElement;
        //    try
        //    {
        //        RunProcess(obj.Tag.ToString());
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}
        //private void aceF04_Click(object sender, EventArgs e)
        //{
        //    var obj = sender as AccordionControlElement;
        //    try
        //    {
        //        RunProcess(obj.Tag.ToString());
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}
        //private void aceF05_Click(object sender, EventArgs e)
        //{
        //    var obj = sender as AccordionControlElement;
        //    try
        //    {
        //        RunProcess(obj.Tag.ToString());
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}

    }
}