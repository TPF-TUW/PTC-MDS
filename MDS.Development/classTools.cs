using DevExpress.Utils;
using DevExpress.XtraBars;
using DevExpress.XtraEditors;
using DevExpress.XtraGrid.Views.Grid;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MDS.Development
{
    public class classTools
    {
        public void validate_Numeric(object sender, DevExpress.XtraEditors.Controls.BaseContainerValidateEditorEventArgs e, string colName)
        {
            GridView view = sender as GridView;
            if (view.FocusedColumn.FieldName == colName)
            {
                double val = 0;
                if (!Double.TryParse(e.Value as String, out val))
                {
                    e.Valid = false;
                    e.ErrorText = "Only numeric values are accepted.";
                }
            }
        }

        public void bbi_Hide(BarItem bbi)
        {
            bbi.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
        }

        public void bbi_Show(BarItem bbi)
        {
            bbi.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
        }

        // show Information Message
        public void showInfoMessage(string StrText)
        {
            XtraMessageBox.Show(StrText, "Programs Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // show Error Message
        public void showErrorMessage(string StrText)
        {
            XtraMessageBox.Show(StrText, "Programs Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        // Show Warning Message
        public void showWarningMessage(string StrText)
        {
            XtraMessageBox.Show(StrText, "Programs Warning!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        // Show Question Message
        public void showQuestionMessage(string StrText)
        {
            XtraMessageBox.Show(StrText, "Programs Warning!", MessageBoxButtons.OK, MessageBoxIcon.Question);
        }

        // Show ConfirmDialog YesNo Option
        public bool doConfirm(string text)
        {
            bool b = false;
            DialogResult status = XtraMessageBox.Show(text, "Confirm Dialog", MessageBoxButtons.YesNo);
            if (status == DialogResult.Yes)
            {
                b = true;
            }
            return b;
        }

        // Replace String 
        public string rep(string s)
        {
            return s.Trim().Replace("'", "''");
        }

        //getRowCell
        public string getcell(GridView gvName, string cellName)
        {
            string val = string.Empty;
            val = gvName.GetFocusedRowCellValue(cellName).ToString();
            return val;
        }

        //getCellVal in Gridview
        public string getCellVal(object sender, string colName)
        {
            var s = (sender as GridView);
            return s.GetFocusedRowCellValue(colName).ToString();
        }

        //clone SMPLNo
        public string genSMPLNo_Clone(string oldSMPLNo, string oldRevise)
        {
            string NewSMPLNo = string.Empty;
            int LEN_SMPLNo = oldSMPLNo.Length;
            string split_SMPLNo = oldSMPLNo.Substring(0, LEN_SMPLNo - 1);
            int revise = Convert.ToInt32(oldRevise) + 1;
            NewSMPLNo = split_SMPLNo + revise;
            return NewSMPLNo;
        }

        //chkCell_isnull
        public bool chkCell_isnull(GridView gvName,string colName,int index,string msg)
        {
            bool b = false;
            if (gvName.GetRowCellValue(gvName.FocusedRowHandle, colName) == null)
            {
                showWarningMessage(msg);
                gvName.FocusedColumn = gvName.VisibleColumns[index];
                gvName.ShowEditor();
                b = true;
            }
            return b;
        }

        //get row in chkbox is checked
        public ArrayList getList_isChecked(GridView gvName)
        {
            ArrayList rows = new ArrayList();
            // Add the selected rows to the list.
            Int32[] selectedRowHandles = gvName.GetSelectedRows();  //getSelectedRow
            for (int i = 0; i < selectedRowHandles.Length; i++)     //Loop SelectedRow
            {
                int selectedRowHandle = selectedRowHandles[i];
                if (selectedRowHandle >= 0)                         //if getSelectedRow >= 0
                {
                    rows.Add(gvName.GetDataRow(selectedRowHandle)); //Add SelectedRow to ArrayList
                }
            }
            return rows;
        }

        /* Get Value Part ----------------------------------------*/
        //getVal_string
        public string getVal_string(string s)
        {
            string val = (s == "") ? "null" : "N'" + s.Trim().Replace("'", "''") + "'"; return val;
        }
        //getVal_text
        public string getVal_text(TextEdit txt)
        {
            string val = (txt.Text.ToString() == "") ? "null" : "N'"+txt.Text.ToString().Trim().Replace("'", "''")+"'"; return val;
        }
        //getVal_num
        public string getVal_num(TextEdit txt)
        {
            string val = (txt.Text.ToString() == "") ? "null" : txt.Text.ToString().Trim().Replace("'", "''"); return val;
        }
        //getVal_sl
        public string getVal_sl(SearchLookUpEdit sl)
        {
            string val = (sl.Text.ToString() == "") ? "null" : sl.EditValue.ToString(); return val;
        }
        //getVal_gl
        public string getVal_gl(GridLookUpEdit gl)
        {
            string val = (gl.Text.ToString() == "") ? "null" : gl.EditValue.ToString(); return val;
        }
        /* End - Get Value Part ----------------------------------------*/

        public enum DepartmentType
        {
            Admin           = 0,
            Packing         = 1,
            NeedleRoom      = 2,
            Warehouse       = 3,
            StoreFabric     = 4,
            StoreAccessory  = 5,
            Delivery        = 6,
            FOA             = 7,
            CMT             = 8,
            Sales           = 9,
            Purchasing      = 10,
            Production      = 11,
            Export          = 12,
            Other           = 99,
        }

        //Open File Image
        public void openFile_Image(XtraOpenFileDialog xopen,TextEdit txt,PictureEdit pic)
        {
            //string fileName = string.Empty;
            xopen.Filter = "Image files | *.jpg; *.jpeg; *.jpe; *.jfif; *.png";
            if (xopen.ShowDialog() == DialogResult.OK)
            {
                string filename = xopen.FileName;
                txt.Text = filename;
                pic.Image = Image.FromFile(filename);
                pic.Properties.SizeMode = DevExpress.XtraEditors.Controls.PictureSizeMode.Zoom;
            }
            //return fileName;
        }

        //Upload Image
        public string uploadImg(TextEdit txt,string newFilenames)
        {
            string imgName = txt.Text.ToString().Trim().Replace("'", "''");
            string newFileName = "null";
            if (imgName != "")
            {
                try
                {
                    string path = @"\\172.16.0.190\MDS_Project\MDS\Pictures\";
                    string filename = imgName;
                    string extension = Path.GetExtension(filename);
                    Random generator = new Random();
                    string r = generator.Next(0, 999999).ToString("D4");
                    newFileName = newFilenames+"-"+DateTime.Now.ToString("yyyyMMdd") + "-" + r + extension;
                    File.Copy(filename, path + Path.GetFileName(newFileName));
                    //MessageBox.Show("Upload Files is Successfull.", "Upload Status");
                }
                catch (Exception)
                {
                    showWarningMessage("Uplaod ไม่ได้ เนื่องจากมีไฟล์นี้ใน Directory ปัจจุบันแล้ว!");
                }
            }
            //if (newFileName != "null")
            //{
            //    newFileName = "N'"+ newFileName + "'";
            //}
            return newFileName;
        }

        public string uploadImg(string txt, string newFilenames)
        {
            string imgName = txt.Trim().Replace("'", "''");
            string newFileName = "null";
            if (imgName != "")
            {
                try
                {
                    string path = @"\\172.16.0.190\MDS_Project\MDS\Pictures\";
                    string filename = imgName;
                    string extension = Path.GetExtension(filename);
                    Random generator = new Random();
                    string r = generator.Next(0, 999999).ToString("D4");
                    newFileName = newFilenames + "-" + DateTime.Now.ToString("yyyyMMdd") + "-" + r + extension;
                    File.Copy(filename, path + Path.GetFileName(newFileName));
                    //MessageBox.Show("Upload Files is Successfull.", "Upload Status");
                }
                catch (Exception)
                {
                    showWarningMessage("Uplaod ไม่ได้ เนื่องจากมีไฟล์นี้ใน Directory ปัจจุบันแล้ว!");
                }
            }
            //if (newFileName != "null")
            //{
            //    newFileName = "N'" + newFileName + "'";
            //}
            return newFileName;
        }
    }
}
