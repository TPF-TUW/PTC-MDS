using DevExpress.Utils.Drawing;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.ViewInfo;
using System;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace PictureEditZoomAndMove.MarkerRectangles
{
    public class RectangleMarker : IDisposable
    {
        private PictureEdit mainPictureEdit;
        private PictureEditViewInfo viewInfo;
        private Rectangle drawingRectangle;
        public Rectangle DrawingRectangle { get { return drawingRectangle; } }

        #region Mouse actions
        private bool isMouseClicked = false;
        private bool mouseMove = false;
        #endregion

        private int oldX;
        private int oldY;
        private int sizeNodeRect = 10;

        private RectangleResizePoints nodeSelected = RectangleResizePoints.None;

        private const int defaultRectangleHeight = 100;
        private const int defaultRectangleWidth = 100;
        private const float rectangleBorderTickness = 5;
        private Color rectangleBorderColor = Color.Red;

        private RectangleF defaultRectangle;
        private PointF imageCenter;
        private double defaultZoomPercent;

        public bool Hidden { get; private set; }
        public bool Editable { get; private set; }

        private Point insertPoint = new Point(0, 0);

        public RectangleMarker(PictureEdit pictureEditForDraw, Rectangle rectangle, bool editable = true)
        {
            mainPictureEdit = pictureEditForDraw;
            viewInfo = mainPictureEdit.GetViewInfo() as PictureEditViewInfo;
            //drawingRectangle = rectangle;
            GetImageCenter();

            //defaultRectangle = new RectangleF(
            //    (float)(imageCenter.X - drawingRectangle.X * pictureEditForDraw.Properties.ZoomPercent / 100),
            //    (float)(imageCenter.Y - drawingRectangle.Y * pictureEditForDraw.Properties.ZoomPercent / 100),
            //    (float)(drawingRectangle.Width * pictureEditForDraw.Properties.ZoomPercent / 100),
            //    (float)(drawingRectangle.Height * pictureEditForDraw.Properties.ZoomPercent / 100));

            defaultZoomPercent = pictureEditForDraw.Properties.ZoomPercent;


            drawingRectangle = new Rectangle(
                rectangle.X,
                rectangle.Y,
                (int)(rectangle.Width * pictureEditForDraw.Properties.ZoomPercent / 100),
                (int)(rectangle.Height * pictureEditForDraw.Properties.ZoomPercent / 100));

            drawingRectangle.Location = new Point(drawingRectangle.Location.X - drawingRectangle.Width / 2, drawingRectangle.Location.Y - drawingRectangle.Height / 2);

            insertPoint = mainPictureEdit.ViewportToImage(rectangle.Location);

            defaultRectangle = new RectangleF(
                (float)(insertPoint.X),
                (float)(insertPoint.Y),
                (float)(drawingRectangle.Width ),
                (float)(drawingRectangle.Height));

            isMouseClicked = false;
            Editable = editable;
            InitEvents();
        }

        private void GetImageCenter()
        {
            imageCenter = new PointF(
                mainPictureEdit.TopLevelControl.Location.X + viewInfo.PictureScreenBounds.X + viewInfo.PictureSourceBounds.X + viewInfo.PictureScreenBounds.Width / 2,
                mainPictureEdit.TopLevelControl.Location.Y + viewInfo.PictureScreenBounds.Y + viewInfo.PictureSourceBounds.Y + viewInfo.PictureScreenBounds.Height / 2);
        }

        private void InitEvents()
        {
            if (Editable)
            {
                mainPictureEdit.MouseDown += new MouseEventHandler(mPictureEdit_MouseDown);
                mainPictureEdit.MouseUp += new MouseEventHandler(mPictureEdit_MouseUp);
                mainPictureEdit.MouseMove += new MouseEventHandler(mPictureEdit_MouseMove);
            }
            // mainPictureEdit.Paint += new PaintEventHandler(mPictureEdit_Paint);
            // mainPictureEdit.ZoomPercentChanged += ZoomPercentChanged;
            mainPictureEdit.PaintEx += MainPictureEdit_PaintEx;
        }

        private void MainPictureEdit_PaintEx(object sender, DevExpress.XtraGrid.PaintExEventArgs e)
        {
            double actualZoomFactor = GetActualZoomFactor(sender as PictureEdit);
            if (actualZoomFactor == -1) return;

            double multiplier = actualZoomFactor * 100 / defaultZoomPercent;

            if (insertPoint.X == 0 && insertPoint.Y == 0) return;

            drawingRectangle.Width = Convert.ToInt32(defaultRectangle.Width * multiplier);
            drawingRectangle.Height = Convert.ToInt32(defaultRectangle.Height * multiplier);

            Point viewportPoint = mainPictureEdit.ImageToViewport(insertPoint);
            viewportPoint.X -= drawingRectangle.Width / 2;
            viewportPoint.Y -= drawingRectangle.Height / 2;
            drawingRectangle.Location = viewportPoint;


            try
            {
                if (!Hidden)
                    DrawEx(e.Cache);
            }
            catch (Exception exception)
            {
                XtraMessageBox.Show(exception.ToString());
            }
        }

        private void DrawEx(GraphicsCache graphic)
        {
            graphic.DrawRectangle(new Pen(rectangleBorderColor, rectangleBorderTickness), drawingRectangle);

            foreach (RectangleResizePoints point in Enum.GetValues(typeof(RectangleResizePoints)))
                graphic.DrawRectangle(new Pen(rectangleBorderColor), GetRectangle(point));
        }

        protected double GetActualZoomFactor(PictureEdit Owner)
        {
            PictureEditViewInfo viewInfo = Owner.GetViewInfo() as PictureEditViewInfo;
            if (Owner.Image != null && viewInfo.PictureSourceBounds.Width > 0)
            {
                return viewInfo.PictureScreenBounds.Width / viewInfo.PictureSourceBounds.Width;
            }
            return -1;
        }

        /* private void Draw(Graphics graphic)
         {
             graphic.DrawRectangle(new Pen(rectangleBorderColor, rectangleBorderTickness), drawingRectangle);

             foreach (RectangleResizePoints point in Enum.GetValues(typeof(RectangleResizePoints)))
                 graphic.DrawRectangle(new Pen(rectangleBorderColor), GetRectangle(point));
         }*/

        /*private void mPictureEdit_Paint(object sender, PaintEventArgs e)
        {
            try
            {
                if (!Hidden)
                    Draw(e.Graphics);
            }
            catch (Exception exception)
            {
                XtraMessageBox.Show(exception.ToString());
            }
        }*/
        /// <summary>
        /// This updates the default rectangle (we keep the one from start) as we can be zoomed so we want to handle that too
        /// </summary>
        private void UpdateDefaultRectangle()
        {
            double multiplier = defaultZoomPercent / mainPictureEdit.Properties.ZoomPercent;

            defaultRectangle.X = (float)(multiplier * (drawingRectangle.X - imageCenter.X) + imageCenter.X);
            defaultRectangle.Y = (float)(multiplier * (drawingRectangle.Y - imageCenter.Y) + imageCenter.Y);
            defaultRectangle.Height = (float)(drawingRectangle.Height * multiplier);
            defaultRectangle.Width = (float)(drawingRectangle.Width * multiplier);
        }

        //private void UpdateRectangleZoomLevelChanged(double newZoomPercent)
        //{
        //    double multiplier = newZoomPercent / defaultZoomPercent;

        //    drawingRectangle.X = Convert.ToInt32(imageCenter.X - multiplier * (imageCenter.X - defaultRectangle.X));
        //    drawingRectangle.Y = Convert.ToInt32(imageCenter.Y - multiplier * (imageCenter.Y - defaultRectangle.Y));
        //    drawingRectangle.Width = Convert.ToInt32(defaultRectangle.Width * multiplier);
        //    drawingRectangle.Height = Convert.ToInt32(defaultRectangle.Height * multiplier);

        //    PropertyInfo highlightedItemProperty = viewInfo.GetType().GetProperties(BindingFlags.NonPublic | BindingFlags.Instance).Single(pi => pi.Name == "ImageSize");
        //    Size imageSize = (Size)highlightedItemProperty.GetValue(viewInfo, null);


        //    // #TODO: adjust rectangle somehow on X
        //    if (viewInfo.ContentRect.Width < imageSize.Width)
        //    {
        //    }

        //    // #TODO: adjust rectangle somehow on Y
        //    if (viewInfo.ContentRect.Height < imageSize.Height)
        //    {
        //    }
        //}

        //private void ZoomPercentChanged(object sender, EventArgs e)
        //{
        //    UpdateRectangleZoomLevelChanged(mainPictureEdit.Properties.ZoomPercent);
        //}
        
        private void mPictureEdit_MouseDown(object sender, MouseEventArgs e)
        {
            isMouseClicked = true;

            nodeSelected = RectangleResizePoints.None;
            nodeSelected = GetNodeSelectable(e.Location);

            if (drawingRectangle.Contains(new Point(e.X, e.Y)))
            {
                mouseMove = true;
            }
            oldX = e.X;
            oldY = e.Y;
        }

        private void mPictureEdit_MouseUp(object sender, MouseEventArgs e)
        {
            isMouseClicked = false;
            mouseMove = false;
            mainPictureEdit.Refresh();
        }

        private void mPictureEdit_MouseMove(object sender, MouseEventArgs e)
        {
            ChangeCursor(e.Location);
            if (isMouseClicked == false)
                return;

            Rectangle backupRectangle = drawingRectangle;

            switch (nodeSelected)
            {
                case RectangleResizePoints.LeftUp:
                    drawingRectangle.X += e.X - oldX;
                    drawingRectangle.Width -= e.X - oldX;
                    drawingRectangle.Y += e.Y - oldY;
                    drawingRectangle.Height -= e.Y - oldY;
                    break;
                case RectangleResizePoints.LeftMiddle:
                    drawingRectangle.X += e.X - oldX;
                    drawingRectangle.Width -= e.X - oldX;
                    break;
                case RectangleResizePoints.LeftBottom:
                    drawingRectangle.Width -= e.X - oldX;
                    drawingRectangle.X += e.X - oldX;
                    drawingRectangle.Height += e.Y - oldY;
                    break;
                case RectangleResizePoints.BottomMiddle:
                    drawingRectangle.Height += e.Y - oldY;
                    break;
                case RectangleResizePoints.RightUp:
                    drawingRectangle.Width += e.X - oldX;
                    drawingRectangle.Y += e.Y - oldY;
                    drawingRectangle.Height -= e.Y - oldY;
                    break;
                case RectangleResizePoints.RightBottom:
                    drawingRectangle.Width += e.X - oldX;
                    drawingRectangle.Height += e.Y - oldY;
                    break;
                case RectangleResizePoints.RightMiddle:
                    drawingRectangle.Width += e.X - oldX;
                    break;

                case RectangleResizePoints.UpMiddle:
                    drawingRectangle.Y += e.Y - oldY;
                    drawingRectangle.Height -= e.Y - oldY;
                    break;

                default:
                    if (mouseMove)
                    {
                        drawingRectangle.X = drawingRectangle.X + e.X - oldX;
                        drawingRectangle.Y = drawingRectangle.Y + e.Y - oldY;
                    }
                    break;
            }
            oldX = e.X;
            oldY = e.Y;

            // no change
            if (drawingRectangle.Width < 5 || drawingRectangle.Height < 5)
            {
                drawingRectangle = backupRectangle;
            }
            else
            {
                // have to change the default one too, taking into consideration the zoom level
                UpdateDefaultRectangle();
            }

            TestIfRectangleInsidePictureEditArea();

            Hidden = DetectIsRectangleOutsideAllowedArea();

            mainPictureEdit.Invalidate();
        }

        /// <summary>
        /// This code ensures the rectangle does not get out of the main picture edit area
        /// </summary>
        private void TestIfRectangleInsidePictureEditArea()
        {

            if (drawingRectangle.X < 0)
                drawingRectangle.X = 0;
            if (drawingRectangle.Y < 0)
                drawingRectangle.Y = 0;
            if (drawingRectangle.Width <= 0)
                drawingRectangle.Width = 1;
            if (drawingRectangle.Height <= 0)
                drawingRectangle.Height = 1;

            if (drawingRectangle.X + drawingRectangle.Width > mainPictureEdit.Width)
            {
                drawingRectangle.Width = mainPictureEdit.Width - drawingRectangle.X - 1; // -1 to be still show 
                isMouseClicked = false;
            }
            if (drawingRectangle.Y + drawingRectangle.Height > mainPictureEdit.Height)
            {
                drawingRectangle.Height = mainPictureEdit.Height - drawingRectangle.Y - 1;// -1 to be still show 
                isMouseClicked = false;
            }
        }

        private bool DetectIsRectangleOutsideAllowedArea()
        {
            Rectangle imageArea = GetImageArea();

            return !(imageArea.Contains(drawingRectangle) || imageArea.IntersectsWith(drawingRectangle));
        }

        private Rectangle GetImageArea()
        {
            Rectangle imageArea = new Rectangle(
                Convert.ToInt32(mainPictureEdit.TopLevelControl.Location.X + viewInfo.PictureScreenBounds.X),
                Convert.ToInt32(mainPictureEdit.TopLevelControl.Location.Y + viewInfo.PictureScreenBounds.Y),
                Convert.ToInt32(viewInfo.PictureScreenBounds.Width),
                Convert.ToInt32(viewInfo.PictureScreenBounds.Height));

            return imageArea;
        }

        private Rectangle CreateRectSizableNode(int x, int y)
        {
            return new Rectangle(x - sizeNodeRect / 2, y - sizeNodeRect / 2, sizeNodeRect, sizeNodeRect);
        }

        private Rectangle GetRectangle(RectangleResizePoints resizePoint)
        {
            switch (resizePoint)
            {
                case RectangleResizePoints.LeftUp:
                    return CreateRectSizableNode(drawingRectangle.X, drawingRectangle.Y);

                case RectangleResizePoints.LeftMiddle:
                    return CreateRectSizableNode(drawingRectangle.X, drawingRectangle.Y + +drawingRectangle.Height / 2);

                case RectangleResizePoints.LeftBottom:
                    return CreateRectSizableNode(drawingRectangle.X, drawingRectangle.Y + drawingRectangle.Height);

                case RectangleResizePoints.BottomMiddle:
                    return CreateRectSizableNode(drawingRectangle.X + drawingRectangle.Width / 2, drawingRectangle.Y + drawingRectangle.Height);

                case RectangleResizePoints.RightUp:
                    return CreateRectSizableNode(drawingRectangle.X + drawingRectangle.Width, drawingRectangle.Y);

                case RectangleResizePoints.RightBottom:
                    return CreateRectSizableNode(drawingRectangle.X + drawingRectangle.Width, drawingRectangle.Y + drawingRectangle.Height);

                case RectangleResizePoints.RightMiddle:
                    return CreateRectSizableNode(drawingRectangle.X + drawingRectangle.Width, drawingRectangle.Y + drawingRectangle.Height / 2);

                case RectangleResizePoints.UpMiddle:
                    return CreateRectSizableNode(drawingRectangle.X + drawingRectangle.Width / 2, drawingRectangle.Y);
                default:
                    return new Rectangle();
            }
        }

        private RectangleResizePoints GetNodeSelectable(Point point)
        {
            foreach (RectangleResizePoints resizePoint in Enum.GetValues(typeof(RectangleResizePoints)))
                if (GetRectangle(resizePoint).Contains(point))
                    return resizePoint;

            return RectangleResizePoints.None;
        }

        private void ChangeCursor(Point point)
        {
            mainPictureEdit.Cursor = GetCursor(GetNodeSelectable(point));
        }

        /// <summary>
        /// Get cursor for the handle
        /// </summary>
        /// <param name="resizePoint"></param>
        /// <returns></returns>
        private Cursor GetCursor(RectangleResizePoints resizePoint)
        {
            switch (resizePoint)
            {
                case RectangleResizePoints.LeftUp:
                    return Cursors.SizeNWSE;

                case RectangleResizePoints.LeftMiddle:
                    return Cursors.SizeWE;

                case RectangleResizePoints.LeftBottom:
                    return Cursors.SizeNESW;

                case RectangleResizePoints.BottomMiddle:
                    return Cursors.SizeNS;

                case RectangleResizePoints.RightUp:
                    return Cursors.SizeNESW;

                case RectangleResizePoints.RightBottom:
                    return Cursors.SizeNWSE;

                case RectangleResizePoints.RightMiddle:
                    return Cursors.SizeWE;

                case RectangleResizePoints.UpMiddle:
                    return Cursors.SizeNS;
                default:
                    return Cursors.Default;
            }
        }

        public void Dispose()
        {
            if (Editable)
            {
                mainPictureEdit.MouseDown -= new MouseEventHandler(mPictureEdit_MouseDown);
                mainPictureEdit.MouseUp -= new MouseEventHandler(mPictureEdit_MouseUp);
                mainPictureEdit.MouseMove -= new MouseEventHandler(mPictureEdit_MouseMove);
            }
            // mainPictureEdit.Paint -= new PaintEventHandler(mPictureEdit_Paint);
            //mainPictureEdit.ZoomPercentChanged -= ZoomPercentChanged;

            mainPictureEdit.PaintEx -= MainPictureEdit_PaintEx;

            mainPictureEdit = null;
            viewInfo = null;
        }
    }
}
