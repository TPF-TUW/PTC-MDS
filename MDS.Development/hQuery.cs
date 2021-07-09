using DevExpress.XtraEditors;
using DevExpress.XtraGrid;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MDS.Development
{
    public class hardQuery
    {
        goClass.dbConn db = new goClass.dbConn();
        goClass.ctool ct = new goClass.ctool();
        SqlConnection mainConn = new goClass.dbConn().MDS();

        
        /* -------------------------------------------------------- First Page -------------------------------------------------------- */
        public void get_sl_smplNo(SearchLookUpEdit sl)
        {
            string sql = "Select OIDSMPL, SMPLNo From SMPLRequest";
            db.getSl(sql,mainConn,sl, "OIDSMPL", "SMPLNo");
        }

        public void get_gl_Season(GridLookUpEdit gl)
        {
            string sql = "Select distinct s.Season as Season From( Select SUBSTRING( cast(Year(GETDATE())-1 as nvarchar(4)) , 3 , 2)+SeasonNo as Season From Season union Select SUBSTRING( cast(Year(GETDATE()) as nvarchar(4)) , 3 , 2) +SeasonNo as Season From Season union Select SUBSTRING( cast(Year(GETDATE())+1 as nvarchar(4)) , 3 , 2)+SeasonNo as Season From Season) as s left join SMPLRequest as smpl on s.Season = smpl.Season";
            db.getGl(sql,mainConn,gl, "Season", "Season");
        }

        public void get_sl_Customer(SearchLookUpEdit sl)
        {
            string sql = "Select OIDCUST,Name From Customer";
            db.getSl(sql, mainConn, sl, "OIDCUST", "Name");
        }

        public void get_gcListof_Bom(GridControl gc)
        {
            /* รอตาราง Bom ตัวจริงให้ฟาสร้างก่อน */
            string sql = "Select 1 as No , '' as Status , '' as BomNo , '' as Revise , '' as SMPLItem, '' as SMPLNo, '' as Season , '' as Customer , '' as Item , '' as Category , '' as Style , '' as PatternNo , '' as Status";
            db.getDgv(sql,gc,mainConn);
        }
        /* -------------------------------------------------------- End First Page -------------------------------------------------------- */




        /* -------------------------------------------------------- Tab Entry ----------------------------------------------------------- */
        public void get_gl_Branch(GridLookUpEdit gl)
        {
            string sql = "Select OIDBranch,Name From Branchs";
            db.getGl(sql,mainConn,gl, "OIDBranch", "Name");
        }

        public void get_gcListof_SMPL(GridControl gc)
        {
            string sql = "Select smpl.OIDSMPL,SMPLItem,c.ColorName,s.SizeName From SMPLRequest smpl inner join SMPLQuantityRequired q on q.OIDSMPL = smpl.OIDSMPL inner join ProductColor c on c.OIDCOLOR = q.OIDCOLOR inner join ProductSize s on s.OIDSIZE = q.OIDSIZE Where Status = 2";
            db.getDgv(sql,gc,mainConn);
        }
        /* -------------------------------------------------------- End Tab Entry -------------------------------------------------------- */



        /* -------------------------------------------------------- Tab Header ----------------------------------------------------------- */
        public string get_running_BomNo()
        {
            /* รอแก้ Query ดึงจากตาราง Bom */
            string sql = "SELECT CASE WHEN ISNULL(MAX(OIDSMPLMT), '') = '' THEN 1 ELSE MAX(OIDSMPLMT)+1 END AS newOIDMat FROM SMPLRequestMaterial";
            return db.get_oneParameter(sql,mainConn, "newOIDMat");
        }
        public void get_sl_StyleNmae(SearchLookUpEdit sl)
        {
            string sql = "Select OIDSTYLE,StyleName From ProductStyle";
            db.getSl(sql,mainConn,sl, "OIDSTYLE", "StyleName");
        }

        public void get_gl_Category(GridLookUpEdit gl)
        {
            string sql = "Select OIDGCATEGORY,CategoryName FRom GarmentCategory";
            db.getGl(sql,mainConn,gl, "OIDGCATEGORY", "CategoryName");
        }

        public void get_sl_Color(SearchLookUpEdit sl)
        {
            string sql = "Select OIDCOLOR,ColorName From ProductColor";
            db.getSl(sql,mainConn,sl, "OIDCOLOR", "ColorName");
        }

        public void get_sl_Size(SearchLookUpEdit sl)
        {
            string sql = "Select OIDSIZE,SizeName From ProductSize";
            db.getSl(sql,mainConn,sl, "OIDSIZE", "SizeName");
        }

        public void get_gl_Unit(GridLookUpEdit gl)
        {
            string sql = "Select OIDUNIT,UnitName From Unit";
            db.getGl(sql,mainConn,gl, "OIDUNIT", "UnitName");
        }
        /* -------------------------------------------------------- End Tab Header -------------------------------------------------------- */
    }
}
