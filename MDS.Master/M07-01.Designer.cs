namespace MDS.Master
{
    partial class M07_01
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(M07_01));
            this.layoutControl1 = new DevExpress.XtraLayout.LayoutControl();
            this.txeMaterialTypeID = new DevExpress.XtraEditors.TextEdit();
            this.txeMaterialType = new DevExpress.XtraEditors.TextEdit();
            this.btnAdd = new DevExpress.XtraEditors.SimpleButton();
            this.txeItemCode = new DevExpress.XtraEditors.TextEdit();
            this.Root = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlItem1 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem3 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem2 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem4 = new DevExpress.XtraLayout.LayoutControlItem();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).BeginInit();
            this.layoutControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.txeMaterialTypeID.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txeMaterialType.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txeItemCode.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem4)).BeginInit();
            this.SuspendLayout();
            // 
            // layoutControl1
            // 
            this.layoutControl1.Controls.Add(this.txeMaterialTypeID);
            this.layoutControl1.Controls.Add(this.txeMaterialType);
            this.layoutControl1.Controls.Add(this.btnAdd);
            this.layoutControl1.Controls.Add(this.txeItemCode);
            this.layoutControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.layoutControl1.Location = new System.Drawing.Point(0, 0);
            this.layoutControl1.Name = "layoutControl1";
            this.layoutControl1.Root = this.Root;
            this.layoutControl1.Size = new System.Drawing.Size(498, 52);
            this.layoutControl1.TabIndex = 0;
            this.layoutControl1.Text = "layoutControl1";
            // 
            // txeMaterialTypeID
            // 
            this.txeMaterialTypeID.Location = new System.Drawing.Point(388, 4);
            this.txeMaterialTypeID.Name = "txeMaterialTypeID";
            this.txeMaterialTypeID.Properties.Appearance.BackColor = System.Drawing.Color.White;
            this.txeMaterialTypeID.Properties.Appearance.ForeColor = System.Drawing.Color.Maroon;
            this.txeMaterialTypeID.Properties.Appearance.Options.UseBackColor = true;
            this.txeMaterialTypeID.Properties.Appearance.Options.UseForeColor = true;
            this.txeMaterialTypeID.Properties.Appearance.Options.UseTextOptions = true;
            this.txeMaterialTypeID.Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.txeMaterialTypeID.Size = new System.Drawing.Size(64, 20);
            this.txeMaterialTypeID.StyleController = this.layoutControl1;
            this.txeMaterialTypeID.TabIndex = 8;
            // 
            // txeMaterialType
            // 
            this.txeMaterialType.Location = new System.Drawing.Point(89, 4);
            this.txeMaterialType.Name = "txeMaterialType";
            this.txeMaterialType.Properties.Appearance.BackColor = System.Drawing.Color.White;
            this.txeMaterialType.Properties.Appearance.ForeColor = System.Drawing.Color.Black;
            this.txeMaterialType.Properties.Appearance.Options.UseBackColor = true;
            this.txeMaterialType.Properties.Appearance.Options.UseForeColor = true;
            this.txeMaterialType.Properties.ReadOnly = true;
            this.txeMaterialType.Size = new System.Drawing.Size(295, 20);
            this.txeMaterialType.StyleController = this.layoutControl1;
            this.txeMaterialType.TabIndex = 7;
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
            // txeItemCode
            // 
            this.txeItemCode.Location = new System.Drawing.Point(89, 28);
            this.txeItemCode.Name = "txeItemCode";
            this.txeItemCode.Properties.MaxLength = 200;
            this.txeItemCode.Size = new System.Drawing.Size(363, 20);
            this.txeItemCode.StyleController = this.layoutControl1;
            this.txeItemCode.TabIndex = 4;
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
            this.layoutControlItem1.Control = this.txeItemCode;
            this.layoutControlItem1.Location = new System.Drawing.Point(0, 24);
            this.layoutControlItem1.Name = "layoutControlItem1";
            this.layoutControlItem1.Size = new System.Drawing.Size(452, 24);
            this.layoutControlItem1.Text = "Item Code";
            this.layoutControlItem1.TextSize = new System.Drawing.Size(73, 14);
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
            this.layoutControlItem2.Control = this.txeMaterialType;
            this.layoutControlItem2.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem2.Name = "layoutControlItem2";
            this.layoutControlItem2.Size = new System.Drawing.Size(384, 24);
            this.layoutControlItem2.Text = "Material Type";
            this.layoutControlItem2.TextSize = new System.Drawing.Size(73, 14);
            // 
            // layoutControlItem4
            // 
            this.layoutControlItem4.Control = this.txeMaterialTypeID;
            this.layoutControlItem4.Location = new System.Drawing.Point(384, 0);
            this.layoutControlItem4.MaxSize = new System.Drawing.Size(68, 24);
            this.layoutControlItem4.MinSize = new System.Drawing.Size(68, 24);
            this.layoutControlItem4.Name = "layoutControlItem4";
            this.layoutControlItem4.Size = new System.Drawing.Size(68, 24);
            this.layoutControlItem4.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem4.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem4.TextVisible = false;
            // 
            // M07_01
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(498, 52);
            this.Controls.Add(this.layoutControl1);
            this.IconOptions.Image = ((System.Drawing.Image)(resources.GetObject("M07_01.IconOptions.Image")));
            this.MaximizeBox = false;
            this.Name = "M07_01";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Add Item Code";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.M07_01_FormClosed);
            this.Load += new System.EventHandler(this.M07_01_Load);
            this.Shown += new System.EventHandler(this.M07_01_Shown);
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).EndInit();
            this.layoutControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.txeMaterialTypeID.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txeMaterialType.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txeItemCode.Properties)).EndInit();
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
        private DevExpress.XtraEditors.TextEdit txeItemCode;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem1;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem3;
        private DevExpress.XtraEditors.TextEdit txeMaterialType;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem2;
        private DevExpress.XtraEditors.TextEdit txeMaterialTypeID;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem4;
    }
}