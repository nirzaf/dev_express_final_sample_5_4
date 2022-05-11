namespace sample_5_4 {
  partial class CustomerList {
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
      this.customerCollection = new DevExpress.Xpo.XPCollection(this.components);
      this.unitOfWork = new DevExpress.Xpo.UnitOfWork(this.components);
      this.saveChangesItem = new DevExpress.XtraBars.BarButtonItem();
      this.discardChangesItem = new DevExpress.XtraBars.BarButtonItem();
      this.printExportItem = new DevExpress.XtraBars.BarButtonItem();
      this.ribbonControl1 = new DevExpress.XtraBars.Ribbon.RibbonControl();
      this.ribbonPageCategory1 = new DevExpress.XtraBars.Ribbon.RibbonPageCategory();
      this.ribbonPage1 = new DevExpress.XtraBars.Ribbon.RibbonPage();
      this.ribbonPageGroup1 = new DevExpress.XtraBars.Ribbon.RibbonPageGroup();
      this.ribbonPageGroup2 = new DevExpress.XtraBars.Ribbon.RibbonPageGroup();
      ((System.ComponentModel.ISupportInitialize)(this.customerCollection)).BeginInit();
      ((System.ComponentModel.ISupportInitialize)(this.unitOfWork)).BeginInit();
      ((System.ComponentModel.ISupportInitialize)(this.ribbonControl1)).BeginInit();
      this.SuspendLayout();
      // 
      // customerCollection
      // 
      this.customerCollection.ObjectType = typeof(EventsDB.Customer);
      this.customerCollection.Session = this.unitOfWork;
      // 
      // saveChangesItem
      // 
      this.saveChangesItem.Caption = "Save Changes";
      this.saveChangesItem.Id = 1;
      this.saveChangesItem.ImageOptions.SvgImage = global::sample_5_4.Properties.Resources.savechanges;
      this.saveChangesItem.Name = "saveChangesItem";
      // 
      // discardChangesItem
      // 
      this.discardChangesItem.Caption = "Discard Changes";
      this.discardChangesItem.Id = 2;
      this.discardChangesItem.ImageOptions.SvgImage = global::sample_5_4.Properties.Resources.discardchanges;
      this.discardChangesItem.Name = "discardChangesItem";
      // 
      // printExportItem
      // 
      this.printExportItem.Caption = "Print && Export";
      this.printExportItem.Id = 3;
      this.printExportItem.ImageOptions.SvgImage = global::sample_5_4.Properties.Resources.printexport;
      this.printExportItem.Name = "printExportItem";
      // 
      // ribbonControl1
      // 
      this.ribbonControl1.ExpandCollapseItem.Id = 0;
      this.ribbonControl1.Items.AddRange(new DevExpress.XtraBars.BarItem[] {
            this.ribbonControl1.ExpandCollapseItem,
            this.ribbonControl1.SearchEditItem,
            this.saveChangesItem,
            this.discardChangesItem,
            this.printExportItem});
      this.ribbonControl1.Location = new System.Drawing.Point(0, 0);
      this.ribbonControl1.MaxItemId = 1;
      this.ribbonControl1.Name = "ribbonControl1";
      this.ribbonControl1.PageCategories.AddRange(new DevExpress.XtraBars.Ribbon.RibbonPageCategory[] {
            this.ribbonPageCategory1});
      this.ribbonControl1.Size = new System.Drawing.Size(569, 150);
      // 
      // ribbonPageCategory1
      // 
      this.ribbonPageCategory1.Appearance.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
      this.ribbonPageCategory1.Appearance.Options.UseBackColor = true;
      this.ribbonPageCategory1.Name = "ribbonPageCategory1";
      this.ribbonPageCategory1.Pages.AddRange(new DevExpress.XtraBars.Ribbon.RibbonPage[] {
            this.ribbonPage1});
      this.ribbonPageCategory1.Text = "Customers";
      // 
      // ribbonPage1
      // 
      this.ribbonPage1.Groups.AddRange(new DevExpress.XtraBars.Ribbon.RibbonPageGroup[] {
            this.ribbonPageGroup1,
            this.ribbonPageGroup2});
      this.ribbonPage1.Name = "ribbonPage1";
      this.ribbonPage1.Text = "Data Operations";
      // 
      // ribbonPageGroup1
      // 
      this.ribbonPageGroup1.ItemLinks.Add(this.saveChangesItem);
      this.ribbonPageGroup1.ItemLinks.Add(this.discardChangesItem);
      this.ribbonPageGroup1.Name = "ribbonPageGroup1";
      this.ribbonPageGroup1.Text = "Persistence";
      // 
      // ribbonPageGroup2
      // 
      this.ribbonPageGroup2.ItemLinks.Add(this.printExportItem);
      this.ribbonPageGroup2.Name = "ribbonPageGroup2";
      this.ribbonPageGroup2.Text = "Tools";
      // 
      // CustomerList
      // 
      this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.Controls.Add(this.ribbonControl1);
      this.Name = "CustomerList";
      this.Size = new System.Drawing.Size(569, 338);
      this.Load += new System.EventHandler(this.CustomerList_Load);
      ((System.ComponentModel.ISupportInitialize)(this.customerCollection)).EndInit();
      ((System.ComponentModel.ISupportInitialize)(this.unitOfWork)).EndInit();
      ((System.ComponentModel.ISupportInitialize)(this.ribbonControl1)).EndInit();
      this.ResumeLayout(false);
      this.PerformLayout();

    }

    #endregion
    private DevExpress.Xpo.UnitOfWork unitOfWork;
    private DevExpress.Xpo.XPCollection customerCollection;
    private DevExpress.XtraBars.BarButtonItem saveChangesItem;
    private DevExpress.XtraBars.BarButtonItem discardChangesItem;
    private DevExpress.XtraBars.BarButtonItem printExportItem;
    private DevExpress.XtraBars.Ribbon.RibbonControl ribbonControl1;
    private DevExpress.XtraBars.Ribbon.RibbonPage ribbonPage1;
    private DevExpress.XtraBars.Ribbon.RibbonPageGroup ribbonPageGroup1;
    private DevExpress.XtraBars.Ribbon.RibbonPageGroup ribbonPageGroup2;
    private DevExpress.XtraBars.Ribbon.RibbonPageCategory ribbonPageCategory1;
  }
}