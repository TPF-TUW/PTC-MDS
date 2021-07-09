using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraGrid;
using System;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Text;

namespace MDS.Development
{
    public class classConn
    {
        IniFile ini = new IniFile(@"\\172.16.0.190\MDS_Project\MDS\FileConfig\Configue.ini");

        public string _server;
        public string _dbname;
        public string _user;
        public string _password;

        SqlConnection conn;
        SqlDataAdapter da;
        DataSet ds;
        SqlCommand cmd;
        SqlDataReader dr;
        DataTable dt;
        StringBuilder sql;

        public SqlConnection MDS()
        {
            _server = "172.16.0.30";
            _dbname = "MDS";
            _user = "sa";
            _password = "gik8nv@tpf";

            //string section = "ConnectionString";
            //_server = ini.Read("Server",section);
            //_dbname = ini.Read("Database", section);
            //_user = ini.Read("Uid", section);
            //_password = ini.Read("Pwd", section);

            return new SqlConnection("Data Source=" + _server + ";Initial Catalog=" + _dbname + ";Persist Security Info=True;User ID=" + _user + ";Password=" + _password + "");
        }

        public SqlConnection DellInspiron15()
        {
            _server = "S410717NB0201\\MSSQLSERVER2";
            _dbname = "GSS_Test";
            _user = "sa";
            _password = "gik8nv@tpf";
            return new SqlConnection("Data Source=" + _server + ";Initial Catalog=" + _dbname + ";Persist Security Info=True;User ID=" + _user + ";Password=" + _password + "");
        }

        // getData to GridControl
        public void getGc(StringBuilder sql, GridControl dgvName, SqlConnection conn)
        {
            cmd = new SqlCommand(sql.ToString(), conn);
            conn.Open();
            da = new SqlDataAdapter(cmd);
            dt = new DataTable();
            da.Fill(dt);
            conn.Close();
            dgvName.DataSource = dt;
        }

        // getData to GridView
        public void getDgv(string sql, GridControl dgvName, SqlConnection conn)
        {
            cmd = new SqlCommand(sql, conn);
            conn.Open();
            da = new SqlDataAdapter(cmd);
            dt = new DataTable();
            da.Fill(dt);
            conn.Close();
            dgvName.DataSource = dt;
        }

        // get Data to LookupEdit
        public void getGl(string sql, SqlConnection conn, GridLookUpEdit glName, string valName, string displayName)
        {
            cmd = new SqlCommand(sql, conn);
            conn.Open();
            da = new SqlDataAdapter(cmd);
            dt = new DataTable();
            da.Fill(dt);
            conn.Close();
            glName.Properties.DataSource = dt;
            glName.Properties.DisplayMember = displayName;
            glName.Properties.ValueMember = valName;
        }

        // get Data to Repo_LookupEdit
        public void get_repGl(string sql, SqlConnection conn, RepositoryItemGridLookUpEdit glName, string valName, string displayName)
        {
            cmd = new SqlCommand(sql, conn);
            conn.Open();
            da = new SqlDataAdapter(cmd);
            dt = new DataTable();
            da.Fill(dt);
            conn.Close();
            glName.DataSource = dt;
            glName.DisplayMember = displayName;
            glName.ValueMember = valName;
        }

        // get Data to SearchLookupEdit
        public void getSl(string sql, SqlConnection conn, SearchLookUpEdit slName, string valName, string displayName)
        {
            cmd = new SqlCommand(sql, conn);
            conn.Open();
            da = new SqlDataAdapter(cmd);
            dt = new DataTable();
            da.Fill(dt);
            conn.Close();
            slName.Properties.DataSource = dt;
            slName.Properties.DisplayMember = displayName;
            slName.Properties.ValueMember = valName;
        }

        // Load Data to ComboBox
        //public void getCbo(string sql, ComboBox cboName, SqlConnection conn)
        //{
        //    cboName.Items.Clear();
        //    //conn = new dbConn().GSSv2_Prod();
        //    cmd = new SqlCommand(sql, conn);
        //    cmd.CommandText = sql;
        //    conn.Open();
        //    dr = cmd.ExecuteReader();
        //    while (dr.Read())
        //    {
        //        cboName.Items.Add(dr[0].ToString());
        //    }
        //    conn.Close();
        //}

        // dbQuery Select : Check True/False
        public bool get(string sql, SqlConnection conn)
        {
            bool b = false;
            cmd = new SqlCommand(sql, conn);
            conn.Open();
            dr = cmd.ExecuteReader();
            if (dr.Read() == true)
            {
                b = true;
            }
            cmd.Dispose();
            conn.Close();
            return b;
        }

        // Select One Columns
        public string get_oneParameter(string sql, SqlConnection conn, string colName)
        {
            string rs = string.Empty;
            cmd = new SqlCommand(sql, conn);
            conn.Open();
            dr = cmd.ExecuteReader();
            if (dr.Read() == true)
            {
                rs = dr[colName].ToString();
            }
            dr.Close();
            cmd.Dispose();
            conn.Close();
            return rs;
        }

        // Select One Columns
        public string getsb_oneParameter(StringBuilder sql, SqlConnection conn, string colName)
        {
            string rs = string.Empty;
            cmd = new SqlCommand(sql.ToString(), conn);
            conn.Open();
            dr = cmd.ExecuteReader();
            if (dr.Read() == true)
            {
                rs = dr[colName].ToString();
            }
            dr.Close();
            cmd.Dispose();
            conn.Close();
            return rs;
        }

        // dbQuery Insert/Update/Delete
        public int Query(string sql, SqlConnection conn)
        {
            int i;
            cmd = new SqlCommand(sql, conn);
            conn.Open();
            cmd.CommandType = CommandType.Text;
            i = cmd.ExecuteNonQuery();
            conn.Close();
            return i;
        }

        /*Main Repo :: SMPLQuantiityRequired*/
        public class FGRequest
        {
            public FGRequest() { }
            public FGRequest(Int32 no, string color, string size, Int32 quantity, string unit)
            {
                No = no; Color = color; Size = size; Quantity = quantity; Unit = unit;
            }
            public Int32 No { get; set; }
            public string Color { get; set; }
            public string Size { get; set; }
            public Int32 Quantity { get; set; }
            public string Unit { get; set; }
        }
        //Create DataSourse
        public BindingList<FGRequest> FGRequestDS()
        {
            BindingList<FGRequest> ds = new BindingList<FGRequest>();
            //ds.Add(new FGRequest(1, "Black", "XL", 10));
            //ds.AllowNew = true;
            return ds;
        }

        /*Mat Repo :: G7*/
        public class MatRequest
        {
            public MatRequest() { }
            public MatRequest(Int32 no, string color, string size, string consumption, string unit, string smplID)
            {
                No = no; Color = color; Size = size; Consumption = consumption; Unit = unit; SmplID = smplID;
            }
            public Int32 No { get; set; }
            public string Color { get; set; }
            public string Size { get; set; }
            public string Consumption { get; set; }
            public string Unit { get; set; }
            public string SmplID { get; set; }
        }
        // ds Mat
        public BindingList<MatRequest> dsMat()
        {
            BindingList<MatRequest> ds = new BindingList<MatRequest>();
            return ds;
        }

        public string getDataFrom_SMPL(string fieldName, string OID)
        {
            string s = string.Empty;
            s = get_oneParameter("Select " + fieldName + " From SMPLRequest Where OIDSMPL = " + OID + " ", MDS(), fieldName);
            return s.Trim();
        }

        /* --------------------------------------------------------------------------------- Special Query This Project Only ------------------------------------------------------------------------- */
        public string get_newOIDMat()
        {
            string sql          = "SELECT CASE WHEN ISNULL(MAX(OIDSMPLMT), '') = '' THEN 1 ELSE MAX(OIDSMPLMT) + 1 END AS newOIDMat FROM SMPLRequestMaterial";
            string newOIDMat    = get_oneParameter(sql, MDS(), "newOIDMat");
            return newOIDMat;
        }

        public void getGrid_SMPL(GridControl glName, DevExpress.XtraGrid.Views.Grid.GridView gvName, int OIDUser = 0, int showDoc = 1, int showUser = 0)
        {
            sql = new StringBuilder();
            sql.Append("SELECT smpl.OIDSMPL AS ID, smpl.SMPLNo AS [SMPL No.], smpl.Status, (CASE smpl.Status WHEN 0 THEN 'New' WHEN 1 THEN 'Wait Approved' WHEN 2 THEN 'Customer Approved' END) AS [Status Name],  ");
            sql.Append("       (CASE WHEN smpl.SMPLRevise = 0 THEN '' ELSE CONVERT(VARCHAR, smpl.SMPLRevise) END) AS [SMPL Revise], smpl.OIDBranch, b.Name AS Branch, smpl.OIDDEPT, d.Name AS[Sales Section], (CASE WHEN ISNULL(smpl.RequestDate, '') = '' THEN '' ELSE CONVERT(VARCHAR(10), smpl.RequestDate, 103) END) AS RequestDate, smpl.SpecificationSize, (CASE smpl.SpecificationSize WHEN 0 THEN 'Necessary' WHEN 1 THEN 'Unnecessary' END) AS [Specification Size], smpl.Season, smpl.OIDCUST, c.Name AS Customer, ");
            sql.Append("       smpl.UseFor, ISNULL(UF.UseFor, '') AS [Use For], ");
            sql.Append("       smpl.OIDCATEGORY, g.CategoryName AS Category, smpl.OIDSTYLE, p.StyleName AS Style, smpl.SMPLItem AS[SMPL Item], smpl.SMPLPatternNo AS[SMPL Pattern No.], smpl.PatternSizeZone, (CASE smpl.PatternSizeZone WHEN 0 THEN 'Japan' WHEN 1 THEN 'Europe' WHEN 2 THEN 'US' END) AS [Pattern Size Zone], ");
            sql.Append("       smpl.CustApproved AS[Customer Approved], (CASE smpl.CustApproved WHEN 0 THEN '-' WHEN 1 THEN 'Yes' END) AS CustomerApprovedStatus, (CASE WHEN ISNULL(smpl.CustApprovedDate, '') = '' THEN '' ELSE CONVERT(VARCHAR(10), smpl.CustApprovedDate, 103) END) AS CustomerApprovedDate, smpl.ReferenceNo, smpl.ContactName, ");
            sql.Append("       (CASE WHEN ISNULL(smpl.DeliveryRequest, '') = '' THEN '' ELSE CONVERT(VARCHAR(10), smpl.DeliveryRequest, 103) END) AS DeliveryRequest, smpl.ModelName, smpl.Situation, smpl.StateArrangements, smpl.ACPurRecBy, (CASE smpl.ACPurRecBy WHEN 0 THEN '-' WHEN 1 THEN 'Yes' END) AS[Accessory Purchase Received], ");
            sql.Append("       (CASE WHEN ISNULL(smpl.ACPurRecDate, '') = '' THEN '' ELSE CONVERT(VARCHAR(10), smpl.ACPurRecDate, 103) END) AS [Accessory Purchase Received Date], smpl.FBPurRecBy, (CASE smpl.FBPurRecBy WHEN 0 THEN '-' WHEN 1 THEN 'Yes' END) AS[Fabric Purchase Received],  ");
            sql.Append("       (CASE WHEN ISNULL(smpl.FBPurRecDate, '') = '' THEN '' ELSE CONVERT(VARCHAR(10), smpl.FBPurRecDate, 103) END) AS [Fabric Purchase Received Date], ISNULL(smpl.PictureFile, '') AS PictureFile, smpl.CreatedBy AS ByCreated, smpl.CreatedDate AS DateCreated, smpl.UpdatedBy, smpl.UpdatedDate, ");
            sql.Append("       ISNULL((SELECT TOP(1) ITM.Code FROM SMPLRequestFabric AS SRFB INNER JOIN Items AS ITM ON SRFB.OIDITEM = ITM.OIDITEM WHERE(SRFB.OIDSMPLDT IN (SELECT xSQR.OIDSMPLDT FROM SMPLRequest AS xSRQ INNER JOIN SMPLQuantityRequired AS xSQR ON xSRQ.OIDSMPL = xSQR.OIDSMPL WHERE(xSRQ.OIDSMPL = smpl.OIDSMPL))) AND (ITM.MaterialType = '8') AND (ITM.Code LIKE 'TMPFB%')), '') AS ChkFBCode, ");
            sql.Append("       ISNULL((SELECT TOP(1) ITM.Code FROM SMPLRequestMaterial AS SRMT INNER JOIN Items AS ITM ON SRMT.OIDITEM = ITM.OIDITEM WHERE(SRMT.OIDSMPLDT IN (SELECT xSQR.OIDSMPLDT FROM SMPLRequest AS xSRQ INNER JOIN SMPLQuantityRequired AS xSQR ON xSRQ.OIDSMPL = xSQR.OIDSMPL WHERE(xSRQ.OIDSMPL = smpl.OIDSMPL))) AND (ITM.MaterialType = '8') AND (ITM.Code LIKE 'TMPMT%')), '') AS ChkMTCode, u.FullName AS CreatedBy, smpl.CreatedDate AS CreatedDate, smpl.SMPLStatus  ");
            sql.Append("FROM   SMPLRequest AS smpl LEFT OUTER JOIN ");
            sql.Append("       SMPLUseFor AS UF ON smpl.UseFor = UF.OIDUF LEFT OUTER JOIN ");
            sql.Append("       Branchs AS b ON b.OIDBranch = smpl.OIDBranch LEFT OUTER JOIN ");
            sql.Append("       Departments AS d ON d.OIDDEPT = smpl.OIDDEPT LEFT OUTER JOIN ");
            sql.Append("       Customer AS c ON c.OIDCUST = smpl.OIDCUST LEFT OUTER JOIN ");
            sql.Append("       GarmentCategory AS g ON g.OIDGCATEGORY = smpl.OIDCATEGORY LEFT OUTER JOIN ");
            sql.Append("       ProductStyle AS p ON p.OIDSTYLE = smpl.OIDSTYLE LEFT OUTER JOIN ");
            sql.Append("       Users AS u ON smpl.CreatedBy = u.OIDUSER ");
            sql.Append("WHERE  (smpl.SMPLNo <> N'') ");
            if (showDoc == 1)
                sql.Append("AND  (smpl.SMPLStatus = 1) ");
            if( showUser == 0)
                sql.Append("AND  (smpl.CreatedBy = '" + OIDUser + "') ");
            sql.Append("ORDER BY smpl.CreatedDate DESC ");
            new ObjDevEx.setGridControl(glName, gvName, sql).getData(false, false, false, true);
            //getGc(sql, glName,MDS());

            gvName.Columns[0].Visible = false; //OIDSMPL
            gvName.Columns[2].Visible = false; //Status
            gvName.Columns[5].Visible = false; //OIDBranch
            gvName.Columns[7].Visible = false; //OIDDEPT
            gvName.Columns[10].Visible = false; //SpecificationSize
            gvName.Columns[13].Visible = false; //OIDCUST
            gvName.Columns[15].Visible = false; //UseFor
            gvName.Columns[17].Visible = false; //OIDCATEGORY
            gvName.Columns[19].Visible = false; //OIDSTYLE
            gvName.Columns[23].Visible = false; //PatternSizeZone
            gvName.Columns[25].Visible = false; //CustApproved
            gvName.Columns[34].Visible = false; //ACPurRecBy
            gvName.Columns[37].Visible = false; //FBPurRecBy
            gvName.Columns[40].Visible = false; //PictureFile
            gvName.Columns[41].Visible = false; //ByCreate
            gvName.Columns[42].Visible = false; //CreateDate
            gvName.Columns[43].Visible = false; //UpdateBy
            gvName.Columns[44].Visible = false; //UpdateDate
            gvName.Columns[45].Visible = false; //ChkFBCode
            gvName.Columns[46].Visible = false; //UpdateDate
            gvName.Columns[49].Visible = false; //ChkMTCode

            gvName.Columns[28].VisibleIndex = 4;

            gvName.Columns["SMPL Revise"].Width = 60;
            gvName.Columns["Sales Section"].Width = 80;
            gvName.Columns["RequestDate"].Width = 100;
            gvName.Columns["Specification Size"].Width = 100;
            gvName.Columns["Pattern Size Zone"].Width = 100;
            gvName.Columns["CustomerApprovedStatus"].Width = 120;
            gvName.Columns["CustomerApprovedDate"].Width = 110;
            gvName.Columns["DeliveryRequest"].Width = 100;
            gvName.Columns["DeliveryRequest"].Width = 100;
            gvName.Columns["Accessory Purchase Received"].Width = 130;
            gvName.Columns["Accessory Purchase Received Date"].Width = 130;
            gvName.Columns["Fabric Purchase Received"].Width = 110;
            gvName.Columns["Fabric Purchase Received Date"].Width = 120;

            gvName.Columns["Status Name"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gvName.Columns["SMPL Revise"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gvName.Columns["Sales Section"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gvName.Columns["RequestDate"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gvName.Columns["Season"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gvName.Columns["Pattern Size Zone"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gvName.Columns["CustomerApprovedStatus"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gvName.Columns["CustomerApprovedDate"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gvName.Columns["DeliveryRequest"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gvName.Columns["Accessory Purchase Received"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gvName.Columns["Accessory Purchase Received Date"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gvName.Columns["Fabric Purchase Received"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gvName.Columns["Fabric Purchase Received Date"].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;

            gvName.Columns[0].Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
            gvName.Columns[1].Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
            gvName.Columns[2].Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
            gvName.Columns[3].Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
            gvName.Columns[4].Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;

        }
        public void getGrid_QuantityReq(GridControl glName)
        {
            /*อันนี้ลบทิ้งได้เลยนะ*/
            sql = new StringBuilder();
            sql.Append("SELECT ROW_NUMBER() OVER(order by OIDSMPLDT asc) as No, SMPLQuantityRequired.OIDSMPLDT, SMPLQuantityRequired.OIDSMPL, SMPLRequest.SMPLNo, SMPLRequest.SMPLRevise, SMPLRequest.SMPLItem, SMPLRequest.SMPLPatternNo, SMPLRequest.PatternSizeZone,ProductColor.ColorNo, ProductColor.ColorName, ProductSize.SizeNo, ProductSize.SizeName, SMPLQuantityRequired.Quantity, Unit.UnitName FROM SMPLQuantityRequired INNER JOIN SMPLRequest ON SMPLQuantityRequired.OIDSMPL = SMPLRequest.OIDSMPL INNER JOIN ProductColor ON SMPLQuantityRequired.OIDCOLOR = ProductColor.OIDCOLOR INNER JOIN ProductSize ON SMPLQuantityRequired.OIDSIZE = ProductSize.OIDSIZE INNER JOIN Unit ON SMPLQuantityRequired.OIDUnit = Unit.OIDUNIT/*Where*/Order by SMPLQuantityRequired.OIDSMPLDT");
            getGc(sql, glName, MDS());
        }
        public void getGrid_FBListSample(GridControl gcName,string Where)
        {
            sql = new StringBuilder();
            sql.Append("SELECT smplQR.OIDSMPL,smpl.SMPLNo, smpl.SMPLPatternNo,'' as Consumption, c.ColorName, s.SizeName, smplQR.Quantity,u.UnitName,smplQR.OIDSMPLDT,s.OIDSIZE,u.OIDUNIT FROM SMPLRequest smpl INNER JOIN SMPLQuantityRequired smplQR ON smpl.OIDSMPL = smplQR.OIDSMPL INNER JOIN ProductColor c ON smplQR.OIDCOLOR = c.OIDCOLOR INNER JOIN ProductSize s ON smplQR.OIDSIZE = s.OIDSIZE INNER JOIN Unit u ON smplQR.OIDUnit = u.OIDUNIT Where smplQR.OIDSMPL is not null " + Where + " Order By smplQR.OIDSMPL,smpl.SMPLPatternNo,c.ColorName");
            getGc(sql, gcName,MDS());
        }
        public string genSMPLNo()
        {
            string SMPLNo = string.Empty;
            sql = new StringBuilder();
            sql.Append("Select SUBSTRING(Season,1,4)+'S'+cast(OIDDEPT as nvarchar(10))+SUBSTRING( /*string*/'0000'+cast(SUBSTRING(SMPLNo,7,4)+1 as nvarchar(max)) ,/*start*/LEN('0000'+cast(SUBSTRING(SMPLNo,7,4)+1 as nvarchar(max)))-3 ,/*length*/4)+'-0'/*+cast(SUBSTRING(SMPLNo,12,1)+1 as nvarchar(max))*/ as SMPLNo From SMPLRequest Where OIDSMPL =(Select MAX(OIDSMPL) From SMPLRequest)");
            SMPLNo = getsb_oneParameter(sql,MDS(), "SMPLNo");
            return SMPLNo;
        }

        // Tab : Fabric
        public string get_newOIDFB()
        {
            string sql = "SELECT CASE WHEN ISNULL(MAX(OIDSMPLFB), '') = '' THEN 1 ELSE MAX(OIDSMPLFB) + 1 END AS newOIDFB FROM SMPLRequestFabric";
            string newOIDFB = get_oneParameter(sql, MDS(), "newOIDFB");
            return newOIDFB;
        }
        public void getListofFabric(GridControl gc, string OIDSMPL)
        {
            string sqlFB = "Select fb.OIDSMPLFB as No , VendFBCode,fb.Composition,FBWeight,c.ColorName as ColorName,SMPLotNo,v.Name as Supplier,i.Code as NAVCode From SMPLRequestFabric fb inner join SMPLQuantityRequired q on q.OIDSMPLDT = fb.OIDSMPLDT inner join SMPLRequest smpl on smpl.OIDSMPL = q.OIDSMPL inner join ProductColor c on c.OIDCOLOR = fb.OIDCOLOR inner join Items i on i.OIDITEM = fb.OIDITEM inner join Vendor v on v.OIDVEND = fb.OIDVEND Where smpl.OIDSMPL = "+OIDSMPL+" ";
            getDgv(sqlFB,gc,MDS());
        }

        // Tab : Material
        public void getListofMaterial(GridControl gc,string dosetOIDSMPL)
        {
            string sqlMat = "Select m.OIDSMPLMT as MatID,q.OIDSMPL as SampleID,d.Name as WorkStation,VendMTCode,SMPLotNo,v.Name as Vendor,c.ColorName as MatColor,s.SizeName as MatSize,m.Composition,Details,Price,cr.Currency as Currency,i.Code as NAVCode,m.Situation,Comment,Remark,m.PathFile ,Consumption/*,m.OIDUNIT */ From SMPLRequestMaterial m inner join SMPLQuantityRequired q on q.OIDSMPLDT = m.OIDSMPLDT inner join Departments d on d.OIDDEPT = m.OIDDEPT inner join Vendor v on v.OIDVEND = m.OIDVEND left join ProductColor c on c.OIDCOLOR = m.MTColor inner join ProductSize s on s.OIDSIZE = m.MTSize left join Currency cr on cr.OIDCURR = m.OIDCURR left join Items i on i.OIDITEM = m.OIDITEM Where q.OIDSMPL = " + dosetOIDSMPL + " ";
            getDgv(sqlMat, gc, MDS());
        }
        public void get_gl_WorkStationMat(GridLookUpEdit gl)
        {
            string sql = "Select OIDDEPT,brn.Name as BranName,dep.Name as Department From Departments dep inner join Branchs brn on brn.OIDBranch = dep.OIDBRANCH Where DepartmentType in(1,4,5) order by OIDDEPT";
            getGl(sql,MDS(),gl, "OIDDEPT", "Department");
        }
    }
}
