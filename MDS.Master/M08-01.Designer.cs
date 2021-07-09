namespace MDS.Master
{
    partial class M08_01
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(M08_01));
            this.layoutControl1 = new DevExpress.XtraLayout.LayoutControl();
            this.txeBranchID = new DevExpress.XtraEditors.TextEdit();
            this.txeBranch = new DevExpress.XtraEditors.TextEdit();
            this.btnAdd = new DevExpress.XtraEditors.SimpleButton();
            this.txeLineName = new DevExpress.XtraEditors.TextEdit();
            this.Root = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlItem1 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem3 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem2 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem4 = new DevExpress.XtraLayout.LayoutControlItem();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).BeginInit();
            this.layoutControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.txeBranchID.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txeBranch.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txeLineName.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem4)).BeginInit();
            this.SuspendLayout();
            // 
            // layoutControl1
            // 
            this.layoutControl1.Controls.Add(this.txeBranchID);
            this.layoutControl1.Controls.Add(this.txeBranch);
            this.layoutControl1.Controls.Add(this.btnAdd);
            this.layoutControl1.Controls.Add(this.txeLineName);
            this.layoutControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.layoutControl1.Location = new System.Drawing.Point(0, 0);
            this.layoutControl1.Name = "layoutControl1";
            this.layoutControl1.Root = this.Root;
            this.layoutControl1.Size = new System.Drawing.Size(498, 52);
            this.layoutControl1.TabIndex = 0;
            this.layoutControl1.Text = "layoutControl1";
            // 
            // txeBranchID
            // 
            this.txeBranchID.Location = new System.Drawing.Point(388, 4);
            this.txeBranchID.Name = "txeBranchID";
            this.txeBranchID.Properties.Appearance.BackColor = System.Drawing.Color.White;
            this.txeBranchID.Properties.Appearance.ForeColor = System.Drawing.Color.Maroon;
            this.txeBranchID.Properties.Appearance.Options.UseBackColor = true;
            this.txeBranchID.Properties.Appearance.Options.UseForeColor = true;
            this.txeBranchID.Properties.Appearance.Options.UseTextOptions = true;
            this.txeBranchID.Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.txeBranchID.Size = new System.Drawing.Size(64, 20);
            this.txeBranchID.StyleController = this.layoutControl1;
            this.txeBranchID.TabIndex = 8;
            // 
            // txeBranch
            // 
            this.txeBranch.Location = new System.Drawing.Point(73, 4);
            this.txeBranch.Name = "txeBranch";
            this.txeBranch.Properties.Appearance.BackColor = System.Drawing.Color.White;
            this.txeBranch.Properties.Appearance.ForeColor = System.Drawing.Color.Black;
            this.txeBranch.Properties.Appearance.Options.UseBackColor = true;
            this.txeBranch.Properties.Appearance.Options.UseForeColor = true;
            this.txeBranch.Properties.ReadOnly = true;
            this.txeBranch.Size = new System.Drawing.Size(311, 20);
            this.txeBranch.StyleController = this.layoutControl1;
            this.txeBranch.TabIndex = 7;
            // 
            // btnAdd
            // 
            this.btnAdd.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("btnAdd.ImageOptions.Image")));
            this.btnAdd.Location = new System.Drawing.Point(456, 4);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(38, 44);
            this.btnAdd.StyleController = this.layoutControl1;
            this.btnAdd.TabIndex = 6;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // txeLineName
            // 
            this.txeLineName.Location = new System.Drawing.Point(73, 28);
            this.txeLineName.Name = "txeLineName";
            this.txeLineName.Properties.MaxLength = 200;
            this.txeLineName.Size = new System.Drawing.Size(379, 20);
            this.txeLineName.StyleController = this.layoutControl1;
            this.txeLineName.TabIndex = 4;
            // 
            // Root
            // 
            this.Root.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.Root.GroupBordersVisible = false;
            this.Root.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlItem1,
            this.layoutControlItem3,
            this.layoutControlItem2,
            this.layoutControlItem4});
            this.Root.Name = "Root";
            this.Root.Padding = new DevExpress.XtraLayout.Utils.Padding(2, 2, 2, 2);
            this.Root.Size = new System.Drawing.Size(498, 52);
            this.Root.TextVisible = false;
            // 
            // layoutControlItem1
            // 
            this.layoutControlItem1.Control = this.txeLineName;
            this.layoutControlItem1.Location = new System.Drawing.Point(0, 24);
            this.layoutControlItem1.Name = "layoutControlItem1";
            this.layoutControlItem1.Size = new System.Drawing.Size(452, 24);
            this.layoutControlItem1.Text = "Line Name";
            this.layoutControlItem1.TextSize = new System.Drawing.Size(57, 14);
            // 
            // layoutControlItem3
            // 
            this.layoutControlItem3.Control = this.btnAdd;
            this.layoutControlItem3.Location = new System.Drawing.Point(452, 0);
            this.layoutControlItem3.MaxSize = new System.Drawing.Size(42, 48);
            this.layoutControlItem3.MinSize = new System.Drawing.Size(42, 48);
            this.layoutControlItem3.Name = "layoutControlItem3";
            this.layoutControlItem3.Size = new System.Drawing.Size(42, 48);
            this.layoutControlItem3.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem3.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem3.TextVisible = false;
            // 
            // layoutControlItem2
            // 
            this.layoutControlItem2.Control = this.txeBranch;
            this.layoutControlItem2.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem2.Name = "layoutControlItem2";
            this.layoutControlItem2.Size = new System.Drawing.Size(384, 24);
            this.layoutControlItem2.Text = "Branch";
            this.layoutControlItem2.TextSize = new System.Drawing.Size(57, 14);
            // 
            // layoutControlItem4
            // 
            this.layoutControlItem4.Control = this.txeBranchID;
            this.layoutControlItem4.Location = new System.Drawing.Point(384, 0);
            this.layoutControlItem4.MaxSize = new System.Drawing.Size(68, 24);
            this.layoutControlItem4.MinSize = new System.Drawing.Size(68, 24);
            this.layoutControlItem4.Name = "layoutControlItem4";
            this.layoutControlItem4.Size = new System.Drawing.Size(68, 24);
            this.layoutControlItem4.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem4.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem4.TextVisible = false;
            // 
            // M08_01
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(498, 52);
            this.Controls.Add(this.layoutControl1);
            this.IconOptions.Image = ((System.Drawing.Image)(resources.GetObject("M08_01.IconOptions.Image")));
            this.MaximizeBox = false;
            this.Name = "M08_01";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Add Line Name";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.M08_01_FormClosed);
            this.Load += new System.EventHandler(this.M08_01_Load);
            this.Shown += new System.EventHandler(this.M08_01_Shown);
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).EndInit();
            this.layoutControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.txeBranchID.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txeBranch.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txeLineName.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem4)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraLayout.LayoutControl layoutControl1;
        private DevExpress.XtraLayout.LayoutControlGroup Root;
        private DevExpress.XtraEditors.SimpleButton btnAdd;
        private DevExpress.XtraEditors.TextEdit txeLineName;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem1;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem3;
        private DevExpress.XtraEditors.TextEdit txeBranch;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem2;
        private DevExpress.XtraEditors.TextEdit txeBranchID;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem4;
    }
}