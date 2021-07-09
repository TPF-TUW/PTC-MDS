namespace MDS.Development
{
    partial class DEV01_M10
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DEV01_M10));
            this.layoutControl1 = new DevExpress.XtraLayout.LayoutControl();
            this.btnAddCustomer = new DevExpress.XtraEditors.SimpleButton();
            this.txeSizeName = new DevExpress.XtraEditors.TextEdit();
            this.txeSizeNo = new DevExpress.XtraEditors.TextEdit();
            this.Root = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlItem2 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem3 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem4 = new DevExpress.XtraLayout.LayoutControlItem();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).BeginInit();
            this.layoutControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.txeSizeName.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txeSizeNo.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem4)).BeginInit();
            this.SuspendLayout();
            // 
            // layoutControl1
            // 
            this.layoutControl1.Controls.Add(this.btnAddCustomer);
            this.layoutControl1.Controls.Add(this.txeSizeName);
            this.layoutControl1.Controls.Add(this.txeSizeNo);
            this.layoutControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.layoutControl1.Location = new System.Drawing.Point(0, 0);
            this.layoutControl1.Name = "layoutControl1";
            this.layoutControl1.Root = this.Root;
            this.layoutControl1.Size = new System.Drawing.Size(337, 57);
            this.layoutControl1.TabIndex = 0;
            this.layoutControl1.Text = "layoutControl1";
            // 
            // btnAddCustomer
            // 
            this.btnAddCustomer.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("btnAddCustomer.ImageOptions.Image")));
            this.btnAddCustomer.Location = new System.Drawing.Point(293, 4);
            this.btnAddCustomer.Name = "btnAddCustomer";
            this.btnAddCustomer.Size = new System.Drawing.Size(40, 44);
            this.btnAddCustomer.StyleController = this.layoutControl1;
            this.btnAddCustomer.TabIndex = 7;
            this.btnAddCustomer.Click += new System.EventHandler(this.btnAddCustomer_Click);
            // 
            // txeSizeName
            // 
            this.txeSizeName.Location = new System.Drawing.Point(72, 28);
            this.txeSizeName.Name = "txeSizeName";
            this.txeSizeName.Properties.MaxLength = 50;
            this.txeSizeName.Size = new System.Drawing.Size(217, 20);
            this.txeSizeName.StyleController = this.layoutControl1;
            this.txeSizeName.TabIndex = 6;
            // 
            // txeSizeNo
            // 
            this.txeSizeNo.Location = new System.Drawing.Point(72, 4);
            this.txeSizeNo.Name = "txeSizeNo";
            this.txeSizeNo.Properties.MaxLength = 10;
            this.txeSizeNo.Size = new System.Drawing.Size(217, 20);
            this.txeSizeNo.StyleController = this.layoutControl1;
            this.txeSizeNo.TabIndex = 5;
            // 
            // Root
            // 
            this.Root.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.Root.GroupBordersVisible = false;
            this.Root.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlItem2,
            this.layoutControlItem3,
            this.layoutControlItem4});
            this.Root.Name = "Root";
            this.Root.Padding = new DevExpress.XtraLayout.Utils.Padding(2, 2, 2, 2);
            this.Root.Size = new System.Drawing.Size(337, 57);
            this.Root.TextVisible = false;
            // 
            // layoutControlItem2
            // 
            this.layoutControlItem2.Control = this.txeSizeNo;
            this.layoutControlItem2.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem2.Name = "layoutControlItem2";
            this.layoutControlItem2.Size = new System.Drawing.Size(289, 24);
            this.layoutControlItem2.Text = "Size No.";
            this.layoutControlItem2.TextSize = new System.Drawing.Size(56, 14);
            // 
            // layoutControlItem3
            // 
            this.layoutControlItem3.Control = this.txeSizeName;
            this.layoutControlItem3.Location = new System.Drawing.Point(0, 24);
            this.layoutControlItem3.Name = "layoutControlItem3";
            this.layoutControlItem3.Size = new System.Drawing.Size(289, 29);
            this.layoutControlItem3.Text = "Size Name";
            this.layoutControlItem3.TextSize = new System.Drawing.Size(56, 14);
            // 
            // layoutControlItem4
            // 
            this.layoutControlItem4.Control = this.btnAddCustomer;
            this.layoutControlItem4.Location = new System.Drawing.Point(289, 0);
            this.layoutControlItem4.MaxSize = new System.Drawing.Size(44, 48);
            this.layoutControlItem4.MinSize = new System.Drawing.Size(44, 48);
            this.layoutControlItem4.Name = "layoutControlItem4";
            this.layoutControlItem4.Size = new System.Drawing.Size(44, 53);
            this.layoutControlItem4.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem4.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem4.TextVisible = false;
            // 
            // DEV01_M10
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(337, 57);
            this.Controls.Add(this.layoutControl1);
            this.Name = "DEV01_M10";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Size";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.DEV01_M10_FormClosed);
            this.Load += new System.EventHandler(this.DEV01_M10_Load);
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).EndInit();
            this.layoutControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.txeSizeName.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txeSizeNo.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem4)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraLayout.LayoutControl layoutControl1;
        private DevExpress.XtraLayout.LayoutControlGroup Root;
        private DevExpress.XtraEditors.SimpleButton btnAddCustomer;
        private DevExpress.XtraEditors.TextEdit txeSizeName;
        private DevExpress.XtraEditors.TextEdit txeSizeNo;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem2;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem3;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem4;
    }
}