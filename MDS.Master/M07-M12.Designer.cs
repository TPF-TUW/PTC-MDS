namespace MDS.Master
{
    partial class M07_M12
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(M07_M12));
            this.layoutControl1 = new DevExpress.XtraLayout.LayoutControl();
            this.btnAddCustomer = new DevExpress.XtraEditors.SimpleButton();
            this.txeName = new DevExpress.XtraEditors.TextEdit();
            this.txeCode = new DevExpress.XtraEditors.TextEdit();
            this.cbeType = new DevExpress.XtraEditors.GridLookUpEdit();
            this.gridLookUpEdit1View = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.Root = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlItem1 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem2 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem3 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem4 = new DevExpress.XtraLayout.LayoutControlItem();
            this.txeEmail = new DevExpress.XtraEditors.TextEdit();
            this.layoutControlItem5 = new DevExpress.XtraLayout.LayoutControlItem();
            this.txeAddr3 = new DevExpress.XtraEditors.TextEdit();
            this.layoutControlItem6 = new DevExpress.XtraLayout.LayoutControlItem();
            this.txeTel = new DevExpress.XtraEditors.TextEdit();
            this.layoutControlItem7 = new DevExpress.XtraLayout.LayoutControlItem();
            this.txeAddr2 = new DevExpress.XtraEditors.TextEdit();
            this.layoutControlItem8 = new DevExpress.XtraLayout.LayoutControlItem();
            this.txeAddr1 = new DevExpress.XtraEditors.TextEdit();
            this.layoutControlItem9 = new DevExpress.XtraLayout.LayoutControlItem();
            this.txeCountry = new DevExpress.XtraEditors.TextEdit();
            this.Country = new DevExpress.XtraLayout.LayoutControlItem();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).BeginInit();
            this.layoutControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.txeName.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txeCode.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.cbeType.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridLookUpEdit1View)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txeEmail.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem5)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txeAddr3.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem6)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txeTel.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem7)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txeAddr2.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem8)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txeAddr1.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem9)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txeCountry.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Country)).BeginInit();
            this.SuspendLayout();
            // 
            // layoutControl1
            // 
            this.layoutControl1.Controls.Add(this.txeCountry);
            this.layoutControl1.Controls.Add(this.txeAddr1);
            this.layoutControl1.Controls.Add(this.txeAddr2);
            this.layoutControl1.Controls.Add(this.txeTel);
            this.layoutControl1.Controls.Add(this.txeAddr3);
            this.layoutControl1.Controls.Add(this.txeEmail);
            this.layoutControl1.Controls.Add(this.btnAddCustomer);
            this.layoutControl1.Controls.Add(this.txeName);
            this.layoutControl1.Controls.Add(this.txeCode);
            this.layoutControl1.Controls.Add(this.cbeType);
            this.layoutControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.layoutControl1.Location = new System.Drawing.Point(0, 0);
            this.layoutControl1.Name = "layoutControl1";
            this.layoutControl1.Root = this.Root;
            this.layoutControl1.Size = new System.Drawing.Size(427, 226);
            this.layoutControl1.TabIndex = 0;
            this.layoutControl1.Text = "layoutControl1";
            // 
            // btnAddCustomer
            // 
            this.btnAddCustomer.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("btnAddCustomer.ImageOptions.Image")));
            this.btnAddCustomer.Location = new System.Drawing.Point(383, 28);
            this.btnAddCustomer.Name = "btnAddCustomer";
            this.btnAddCustomer.Size = new System.Drawing.Size(40, 44);
            this.btnAddCustomer.StyleController = this.layoutControl1;
            this.btnAddCustomer.TabIndex = 7;
            this.btnAddCustomer.Click += new System.EventHandler(this.btnAddCustomer_Click);
            // 
            // txeName
            // 
            this.txeName.Location = new System.Drawing.Point(98, 28);
            this.txeName.Name = "txeName";
            this.txeName.Properties.MaxLength = 50;
            this.txeName.Size = new System.Drawing.Size(281, 20);
            this.txeName.StyleController = this.layoutControl1;
            this.txeName.TabIndex = 5;
            // 
            // txeCode
            // 
            this.txeCode.Location = new System.Drawing.Point(98, 4);
            this.txeCode.Name = "txeCode";
            this.txeCode.Properties.MaxLength = 20;
            this.txeCode.Size = new System.Drawing.Size(325, 20);
            this.txeCode.StyleController = this.layoutControl1;
            this.txeCode.TabIndex = 4;
            // 
            // cbeType
            // 
            this.cbeType.Location = new System.Drawing.Point(98, 52);
            this.cbeType.Name = "cbeType";
            this.cbeType.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.cbeType.Properties.MaxLength = 20;
            this.cbeType.Properties.NullText = "";
            this.cbeType.Properties.PopupView = this.gridLookUpEdit1View;
            this.cbeType.Size = new System.Drawing.Size(281, 20);
            this.cbeType.StyleController = this.layoutControl1;
            this.cbeType.TabIndex = 6;
            // 
            // gridLookUpEdit1View
            // 
            this.gridLookUpEdit1View.FocusRectStyle = DevExpress.XtraGrid.Views.Grid.DrawFocusRectStyle.RowFocus;
            this.gridLookUpEdit1View.Name = "gridLookUpEdit1View";
            this.gridLookUpEdit1View.OptionsSelection.EnableAppearanceFocusedCell = false;
            this.gridLookUpEdit1View.OptionsView.ShowGroupPanel = false;
            // 
            // Root
            // 
            this.Root.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.Root.GroupBordersVisible = false;
            this.Root.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlItem1,
            this.layoutControlItem2,
            this.layoutControlItem3,
            this.layoutControlItem4,
            this.layoutControlItem5,
            this.layoutControlItem6,
            this.layoutControlItem7,
            this.layoutControlItem8,
            this.layoutControlItem9,
            this.Country});
            this.Root.Name = "Root";
            this.Root.Padding = new DevExpress.XtraLayout.Utils.Padding(2, 2, 2, 2);
            this.Root.Size = new System.Drawing.Size(427, 226);
            this.Root.TextVisible = false;
            // 
            // layoutControlItem1
            // 
            this.layoutControlItem1.Control = this.txeCode;
            this.layoutControlItem1.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem1.Name = "layoutControlItem1";
            this.layoutControlItem1.Size = new System.Drawing.Size(423, 24);
            this.layoutControlItem1.Text = "Vendor Code";
            this.layoutControlItem1.TextSize = new System.Drawing.Size(82, 14);
            // 
            // layoutControlItem2
            // 
            this.layoutControlItem2.Control = this.txeName;
            this.layoutControlItem2.Location = new System.Drawing.Point(0, 24);
            this.layoutControlItem2.Name = "layoutControlItem2";
            this.layoutControlItem2.Size = new System.Drawing.Size(379, 24);
            this.layoutControlItem2.Text = "Vendor Name";
            this.layoutControlItem2.TextSize = new System.Drawing.Size(82, 14);
            // 
            // layoutControlItem3
            // 
            this.layoutControlItem3.Control = this.cbeType;
            this.layoutControlItem3.Location = new System.Drawing.Point(0, 48);
            this.layoutControlItem3.Name = "layoutControlItem3";
            this.layoutControlItem3.Size = new System.Drawing.Size(379, 24);
            this.layoutControlItem3.Text = "Vendor Type";
            this.layoutControlItem3.TextSize = new System.Drawing.Size(82, 14);
            // 
            // layoutControlItem4
            // 
            this.layoutControlItem4.Control = this.btnAddCustomer;
            this.layoutControlItem4.Location = new System.Drawing.Point(379, 24);
            this.layoutControlItem4.MaxSize = new System.Drawing.Size(44, 48);
            this.layoutControlItem4.MinSize = new System.Drawing.Size(44, 48);
            this.layoutControlItem4.Name = "layoutControlItem4";
            this.layoutControlItem4.Size = new System.Drawing.Size(44, 198);
            this.layoutControlItem4.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem4.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem4.TextVisible = false;
            // 
            // txeEmail
            // 
            this.txeEmail.Location = new System.Drawing.Point(98, 196);
            this.txeEmail.Name = "txeEmail";
            this.txeEmail.Size = new System.Drawing.Size(281, 20);
            this.txeEmail.StyleController = this.layoutControl1;
            this.txeEmail.TabIndex = 8;
            // 
            // layoutControlItem5
            // 
            this.layoutControlItem5.Control = this.txeEmail;
            this.layoutControlItem5.Location = new System.Drawing.Point(0, 192);
            this.layoutControlItem5.Name = "layoutControlItem5";
            this.layoutControlItem5.Size = new System.Drawing.Size(379, 30);
            this.layoutControlItem5.Text = "Email";
            this.layoutControlItem5.TextSize = new System.Drawing.Size(82, 14);
            // 
            // txeAddr3
            // 
            this.txeAddr3.Location = new System.Drawing.Point(98, 124);
            this.txeAddr3.Name = "txeAddr3";
            this.txeAddr3.Size = new System.Drawing.Size(281, 20);
            this.txeAddr3.StyleController = this.layoutControl1;
            this.txeAddr3.TabIndex = 9;
            // 
            // layoutControlItem6
            // 
            this.layoutControlItem6.Control = this.txeAddr3;
            this.layoutControlItem6.Location = new System.Drawing.Point(0, 120);
            this.layoutControlItem6.Name = "layoutControlItem6";
            this.layoutControlItem6.Size = new System.Drawing.Size(379, 24);
            this.layoutControlItem6.Text = "Address 3";
            this.layoutControlItem6.TextSize = new System.Drawing.Size(82, 14);
            // 
            // txeTel
            // 
            this.txeTel.Location = new System.Drawing.Point(98, 172);
            this.txeTel.Name = "txeTel";
            this.txeTel.Size = new System.Drawing.Size(281, 20);
            this.txeTel.StyleController = this.layoutControl1;
            this.txeTel.TabIndex = 10;
            // 
            // layoutControlItem7
            // 
            this.layoutControlItem7.Control = this.txeTel;
            this.layoutControlItem7.Location = new System.Drawing.Point(0, 168);
            this.layoutControlItem7.Name = "layoutControlItem7";
            this.layoutControlItem7.Size = new System.Drawing.Size(379, 24);
            this.layoutControlItem7.Text = "Telephone No.";
            this.layoutControlItem7.TextSize = new System.Drawing.Size(82, 14);
            // 
            // txeAddr2
            // 
            this.txeAddr2.Location = new System.Drawing.Point(98, 100);
            this.txeAddr2.Name = "txeAddr2";
            this.txeAddr2.Size = new System.Drawing.Size(281, 20);
            this.txeAddr2.StyleController = this.layoutControl1;
            this.txeAddr2.TabIndex = 11;
            // 
            // layoutControlItem8
            // 
            this.layoutControlItem8.Control = this.txeAddr2;
            this.layoutControlItem8.Location = new System.Drawing.Point(0, 96);
            this.layoutControlItem8.Name = "layoutControlItem8";
            this.layoutControlItem8.Size = new System.Drawing.Size(379, 24);
            this.layoutControlItem8.Text = "Address 2";
            this.layoutControlItem8.TextSize = new System.Drawing.Size(82, 14);
            // 
            // txeAddr1
            // 
            this.txeAddr1.Location = new System.Drawing.Point(98, 76);
            this.txeAddr1.Name = "txeAddr1";
            this.txeAddr1.Size = new System.Drawing.Size(281, 20);
            this.txeAddr1.StyleController = this.layoutControl1;
            this.txeAddr1.TabIndex = 12;
            // 
            // layoutControlItem9
            // 
            this.layoutControlItem9.Control = this.txeAddr1;
            this.layoutControlItem9.Location = new System.Drawing.Point(0, 72);
            this.layoutControlItem9.Name = "layoutControlItem9";
            this.layoutControlItem9.Size = new System.Drawing.Size(379, 24);
            this.layoutControlItem9.Text = "Address 1";
            this.layoutControlItem9.TextSize = new System.Drawing.Size(82, 14);
            // 
            // txeCountry
            // 
            this.txeCountry.EditValue = "Thailand";
            this.txeCountry.Location = new System.Drawing.Point(98, 148);
            this.txeCountry.Name = "txeCountry";
            this.txeCountry.Size = new System.Drawing.Size(281, 20);
            this.txeCountry.StyleController = this.layoutControl1;
            this.txeCountry.TabIndex = 13;
            // 
            // Country
            // 
            this.Country.Control = this.txeCountry;
            this.Country.Location = new System.Drawing.Point(0, 144);
            this.Country.Name = "Country";
            this.Country.Size = new System.Drawing.Size(379, 24);
            this.Country.TextSize = new System.Drawing.Size(82, 14);
            // 
            // M07_M12
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(427, 226);
            this.Controls.Add(this.layoutControl1);
            this.IconOptions.Image = ((System.Drawing.Image)(resources.GetObject("M07_M12.IconOptions.Image")));
            this.MaximizeBox = false;
            this.Name = "M07_M12";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Supplier (Vendor)";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.M07_M12_FormClosed);
            this.Load += new System.EventHandler(this.M07_M12_Load);
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).EndInit();
            this.layoutControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.txeName.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txeCode.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.cbeType.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridLookUpEdit1View)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txeEmail.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem5)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txeAddr3.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem6)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txeTel.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem7)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txeAddr2.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem8)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txeAddr1.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem9)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txeCountry.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Country)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraLayout.LayoutControl layoutControl1;
        private DevExpress.XtraLayout.LayoutControlGroup Root;
        private DevExpress.XtraEditors.TextEdit txeCode;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem1;
        private DevExpress.XtraEditors.SimpleButton btnAddCustomer;
        private DevExpress.XtraEditors.TextEdit txeName;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem2;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem3;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem4;
        private DevExpress.XtraEditors.GridLookUpEdit cbeType;
        private DevExpress.XtraGrid.Views.Grid.GridView gridLookUpEdit1View;
        private DevExpress.XtraEditors.TextEdit txeAddr1;
        private DevExpress.XtraEditors.TextEdit txeAddr2;
        private DevExpress.XtraEditors.TextEdit txeTel;
        private DevExpress.XtraEditors.TextEdit txeAddr3;
        private DevExpress.XtraEditors.TextEdit txeEmail;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem5;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem6;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem7;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem8;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem9;
        private DevExpress.XtraEditors.TextEdit txeCountry;
        private DevExpress.XtraLayout.LayoutControlItem Country;
    }
}