using DevExpress.XtraEditors.ViewInfo;
using PictureEditZoomAndMove.MarkerRectangles;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace MDS.Master
{
    public partial class ShowImage : DevExpress.XtraEditors.XtraForm
    {
        private List<RectangleMarker> rectangleMarkers;

        public ShowImage(string pathPicture)
        {
            rectangleMarkers = new List<RectangleMarker>();
            InitializeComponent();
            InitPictureEdit(pathPicture);
            pictureEdit.MouseDoubleClick += EventPictureEditXrayOnDoubleClick;
        }

        private void InitPictureEdit(string pathPicture)
        {
            //string path = Path.GetFullPath(Path.Combine(Environment.CurrentDirectory, pathPicture)) + fileName;
            pictureEdit.Image = Image.FromFile(pathPicture);
            pictureEdit.Properties.SizeMode = DevExpress.XtraEditors.Controls.PictureSizeMode.Squeeze;
        }

        private void EventPictureEditXrayOnDoubleClick(object sender, MouseEventArgs e)
        {
            PictureEditViewInfo viewInfo = pictureEdit.GetViewInfo() as PictureEditViewInfo;
            if (!viewInfo.PictureScreenBounds.Contains(e.Location)) return;

            rectangleMarkers.Add(new RectangleMarker(pictureEdit, new Rectangle(e.X, e.Y, 100, 100)));

        }

        private void ShowImage_Load(object sender, EventArgs e)
        {

        }
    }
}