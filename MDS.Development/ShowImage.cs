using DevExpress.XtraEditors.ViewInfo;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;


namespace MDS.Development
{
    public partial class ShowImage : DevExpress.XtraEditors.XtraForm
    {
        string _pathPicture = "";
        public ShowImage(string pathPicture)
        {
            InitializeComponent();
            this._pathPicture = pathPicture;
        }

        private void ShowImage_Load(object sender, EventArgs e)
        {
            if(this._pathPicture != "")
            {
                pictureEdit.Image = null;
                try
                {
                    pictureEdit.Image = Image.FromFile(this._pathPicture);
                }
                catch (Exception) { }
            }

        }
    }
}