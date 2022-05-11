namespace sample_5_4 {
  partial class PivotOverview {
    /// <summary>
    /// Required designer variable.
    /// </summary>
    private System.ComponentModel.IContainer components = null;

    /// <summary>
    /// Clean up any resources being used.
    /// </summary>
    /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
    protected override void Dispose(bool disposing) {
      if (disposing && (components != null)) {
        components.Dispose();
      }
      base.Dispose(disposing);
    }

    #region Windows Form Designer generated code

    /// <summary>
    /// Required method for Designer support - do not modify
    /// the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent() {
      this.components = new System.ComponentModel.Container();
      this.roomBookingCollection = new DevExpress.Xpo.XPCollection(this.components);
      this.session = new DevExpress.Xpo.Session(this.components);
      this.splitContainerControl1 = new DevExpress.XtraEditors.SplitContainerControl();
      this.printExportItem = new DevExpress.XtraBars.BarButtonItem();
      this.ribbonControl1 = new DevExpress.XtraBars.Ribbon.RibbonControl();
      this.ribbonPageCategory1 = new DevExpress.XtraBars.Ribbon.RibbonPageCategory();
      this.ribbonPage1 = new DevExpress.XtraBars.Ribbon.RibbonPage();
      this.ribbonPageGroup1 = new DevExpress.XtraBars.Ribbon.RibbonPageGroup();
      ((System.ComponentModel.ISupportInitialize)(this.roomBookingCollection)).BeginInit();
      ((System.ComponentModel.ISupportInitialize)(this.session)).BeginInit();
      ((System.ComponentModel.ISupportInitialize)(this.splitContainerControl1)).BeginInit();
      this.splitContainerControl1.SuspendLayout();
      ((System.ComponentModel.ISupportInitialize)(this.ribbonControl1)).BeginInit();
      this.SuspendLayout();
      // 
      // roomBookingCollection
      // 
      this.roomBookingCollection.ObjectType = typeof(EventsDB.RoomBooking);
      this.roomBookingCollection.Session = this.session;
      // 
      // splitContainerControl1
      // 
      this.splitContainerControl1.Dock = System.Windows.Forms.DockStyle.Fill;
      this.splitContainerControl1.FixedPanel = DevExpress.XtraEditors.SplitFixedPanel.Panel2;
      this.splitContainerControl1.Horizontal = false;
      this.splitContainerControl1.Location = new System.Drawing.Point(0, 150);
      this.splitContainerControl1.Name = "splitContainerControl1";
      this.splitContainerControl1.Panel1.Text = "Panel1";
      this.splitContainerControl1.Panel2.Text = "Panel2";
      this.splitContainerControl1.Size = new System.Drawing.Size(701, 237);
      this.splitContainerControl1.SplitterPosition = 156;
      this.splitContainerControl1.TabIndex = 1;
      this.splitContainerControl1.Text = "splitContainerControl1";
      // 
      // printExportItem
      // 
      this.printExportItem.Caption = "Print && Export";
      this.printExportItem.Id = 0;
      this.printExportItem.ImageOptions.SvgImage = global::sample_5_4.Properties.Resources.printexport;
      this.printExportItem.Name = "printExportItem";
      // 
      // ribbonControl1
      // 
      this.ribbonControl1.ExpandCollapseItem.Id = 0;
      this.ribbonControl1.Items.AddRange(new DevExpress.XtraBars.BarItem[] {
            this.ribbonControl1.ExpandCollapseItem,
            this.printExportItem,
            this.ribbonControl1.SearchEditItem});
      this.ribbonControl1.Location = new System.Drawing.Point(0, 0);
      this.ribbonControl1.MaxItemId = 1;
      this.ribbonControl1.Name = "ribbonControl1";
      this.ribbonControl1.PageCategories.AddRange(new DevExpress.XtraBars.Ribbon.RibbonPageCategory[] {
            this.ribbonPageCategory1});
      this.ribbonControl1.Size = new System.Drawing.Size(701, 150);
      // 
      // ribbonPageCategory1
      // 
      this.ribbonPageCategory1.Appearance.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
      this.ribbonPageCategory1.Appearance.Options.UseBackColor = true;
      this.ribbonPageCategory1.Name = "ribbonPageCategory1";
      this.ribbonPageCategory1.Pages.AddRange(new DevExpress.XtraBars.Ribbon.RibbonPage[] {
            this.ribbonPage1});
      this.ribbonPageCategory1.Text = "Pivot Overview";
      // 
      // ribbonPage1
      // 
      this.ribbonPage1.Groups.AddRange(new DevExpress.XtraBars.Ribbon.RibbonPageGroup[] {
            this.ribbonPageGroup1});
      this.ribbonPage1.Name = "ribbonPage1";
      this.ribbonPage1.Text = "Data Operations";
      // 
      // ribbonPageGroup1
      // 
      this.ribbonPageGroup1.ItemLinks.Add(this.printExportItem);
      this.ribbonPageGroup1.Name = "ribbonPageGroup1";
      this.ribbonPageGroup1.Text = "Tools";
      // 
      // PivotOverview
      // 
      this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.Controls.Add(this.splitContainerControl1);
      this.Controls.Add(this.ribbonControl1);
      this.Name = "PivotOverview";
      this.Size = new System.Drawing.Size(701, 387);
      ((System.ComponentModel.ISupportInitialize)(this.roomBookingCollection)).EndInit();
      ((System.ComponentModel.ISupportInitialize)(this.session)).EndInit();
      ((System.ComponentModel.ISupportInitialize)(this.splitContainerControl1)).EndInit();
      this.splitContainerControl1.ResumeLayout(false);
      ((System.ComponentModel.ISupportInitialize)(this.ribbonControl1)).EndInit();
      this.ResumeLayout(false);
      this.PerformLayout();

    }

    #endregion
    private DevExpress.Xpo.Session session;
    private DevExpress.Xpo.XPCollection roomBookingCollection;
    private DevExpress.XtraEditors.SplitContainerControl splitContainerControl1;
    private DevExpress.XtraBars.BarButtonItem printExportItem;
    private DevExpress.XtraBars.Ribbon.RibbonControl ribbonControl1;
    private DevExpress.XtraBars.Ribbon.RibbonPage ribbonPage1;
    private DevExpress.XtraBars.Ribbon.RibbonPageGroup ribbonPageGroup1;
    private DevExpress.XtraBars.Ribbon.RibbonPageCategory ribbonPageCategory1;
  }
}