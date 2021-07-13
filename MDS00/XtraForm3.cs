using DevExpress.XtraEditors;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Reflection;
using System.IO;
using DevExpress.XtraBars.Navigation;
using DevExpress.LookAndFeel;
using TheepClass;
using MDS.Master;
using MDS.Development;
using MDS.Function;
using MDS.MPS;


namespace MDS00
{
    public partial class XtraForm3 : DevExpress.XtraEditors.XtraForm
    {
        internal LogIn UserLogin { get; set; }
        
        public XtraForm3()
        {
            InitializeComponent();
            var skinName = cUtility.LoadRegistry(@"Software\MDS", "SkinName");
            var skinPalette = cUtility.LoadRegistry(@"Software\MDS", "SkinPalette");
            UserLookAndFeel.Default.SetSkinStyle(skinName == null ? "Basic" : skinName.ToString(), skinPalette == null ? "Default" : skinPalette.ToString());
            SetMenuExpandedOrCollapse();
        }

        private const int SW_MAXIMIZE = 3;
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);
        
        private void CreateSplashScreen(string processName)
        {
            DevExpress.XtraSplashScreen.SplashScreenManager.ShowSkinSplashScreen(
                logoImage: null,
                title: "MDS",
                subtitle: "Merchandise and Development System",
                footer: "Copyright © 2020-2021 IT Integration Team",
                loading: "Starting..." + processName,
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
        private void RunProcess(string processName,string projectName)
        {
            foreach (DevExpress.XtraEditors.XtraForm frmActive in this.MdiChildren)
            {
                if (frmActive.Name == processName)
                {
                    frmActive.Activate();
                    return;
                }
            }
            var objectType = Type.GetType("MDS."+projectName+"."+processName+",MDS."+projectName);
            dynamic frm = Activator.CreateInstance(objectType);
            frm.UserLogin = UserLogin;
            try
            {
                frm.ConnectionString = "domain-ii,sa,ZAQ113m4tuw,MDS";
            }
            catch (Exception)
            { }
            frm.MdiParent = this;
            frm.WindowState = FormWindowState.Maximized;
            frm.Show();
            
            //if (processName == "F01")
            //{
            //    foreach (DevExpress.XtraEditors.XtraForm frmActive in this.MdiChildren)
            //    {
            //        if (frmActive.Name == processName)
            //        {
            //            frmActive.Activate();
            //            return;
            //        }
            //    }
            //    FileInfo f = new FileInfo(Application.StartupPath + "\\" + processName + ".dll");
            //    var a = Assembly.LoadFrom(f.FullName);
            //    var t = a.GetType(processName + "." + processName);
            //    DevExpress.XtraBars.ToolbarForm.ToolbarForm frm = (DevExpress.XtraBars.ToolbarForm.ToolbarForm)Activator.CreateInstance(t);
            //    frm.MdiParent = this;
            //    frm.WindowState = FormWindowState.Maximized;
            //    frm.Show();
            //}
            //else
            //{
            //    foreach (DevExpress.XtraEditors.XtraForm frmActive in this.MdiChildren)
            //    {
            //        if (frmActive.Name == processName)
            //        {
            //            frmActive.Activate();
            //            return;
            //        }
            //    }
            //    FileInfo f = new FileInfo(Application.StartupPath + "\\" + processName + ".dll");
            //    var a = Assembly.LoadFrom(f.FullName);
            //    var t = a.GetType(processName + "." + processName);
            //    DevExpress.XtraBars.Ribbon.RibbonForm frm = (DevExpress.XtraBars.Ribbon.RibbonForm)Activator.CreateInstance(t);
            //    frm.MdiParent = this;
            //    frm.WindowState = FormWindowState.Maximized;
            //    frm.Show();
            //}
        }
        private void SetMenuExpandedOrCollapse()
        {
            accordionControlElement1.Expanded = Convert.ToBoolean(cUtility.LoadRegistry(@"Software\MDS", "MenuAdministratorExpanded"));
            accordionControlElement2.Expanded = Convert.ToBoolean(cUtility.LoadRegistry(@"Software\MDS", "MenuMasterExpanded"));
            accordionControlElement23.Expanded = Convert.ToBoolean(cUtility.LoadRegistry(@"Software\MDS", "MenuDevelopmentExpanded"));
            accordionControlElement30.Expanded = Convert.ToBoolean(cUtility.LoadRegistry(@"Software\MDS", "MenuMPSExpanded"));
            accordionControlElement41.Expanded = Convert.ToBoolean(cUtility.LoadRegistry(@"Software\MDS", "MenuMRPExpanded"));
            accordionControlElement26.Expanded = Convert.ToBoolean(cUtility.LoadRegistry(@"Software\MDS", "MenuShipmentExpanded"));
            accordionControlElement27.Expanded = Convert.ToBoolean(cUtility.LoadRegistry(@"Software\MDS", "MenuEXIMsExpanded"));
        }
        private void SetMenuEnableOrDisable()
        {
            foreach(var itemGroup in accordionControl1.Elements)
            {
                foreach (var item in itemGroup.Elements)
                {
                    if (UserLogin.Functions.Exists(x => x.FunctionNo == item.Hint && x.AllowDenyStatus==1))
                        item.Visible = true;
                    else
                        item.Visible = false;
                }
            }
        }

        private void accordionControl1_ElementClick(object sender, ElementClickEventArgs e)
        {
            if (e.Element.Style == ElementStyle.Group)
            {
                switch (e.Element.Text)
                {
                    case "Administrator":
                        cUtility.SaveRegistry(@"Software\MDS", "MenuAdministratorExpanded", !e.Element.Expanded);
                        break;
                    case "Master":
                        cUtility.SaveRegistry(@"Software\MDS", "MenuMasterExpanded", !e.Element.Expanded);
                        break;
                    case "Development":
                        cUtility.SaveRegistry(@"Software\MDS", "MenuDevelopmentExpanded", !e.Element.Expanded);
                        break;
                    case "MPS":
                        cUtility.SaveRegistry(@"Software\MDS", "MenuMPSExpanded", !e.Element.Expanded);
                        break;
                    case "MRP":
                        cUtility.SaveRegistry(@"Software\MDS", "MenuMRPExpanded", !e.Element.Expanded);
                        break;
                    case "Shipment":
                        cUtility.SaveRegistry(@"Software\MDS", "MenuShipmentExpanded", !e.Element.Expanded);
                        break;
                    case "EXIMs":
                        cUtility.SaveRegistry(@"Software\MDS", "MenuEXIMsExpanded", !e.Element.Expanded);
                        break;
                }
            }
            else
            {
                CreateSplashScreen(e.Element.Hint);
                try
                {
                    RunProcess(e.Element.Hint,e.Element.Tag.ToString());
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                CloseSplashScreen();
            }
        }
        private void XtraForm3_Load(object sender, EventArgs e)
        {
            SetMenuEnableOrDisable();
        }

        private void XtraForm3_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
            Application.Exit();
        }
    }

    


}