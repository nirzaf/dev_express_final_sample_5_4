namespace sample_5_4 {
  partial class MainForm {
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
      this.ribbon = new DevExpress.XtraBars.Ribbon.RibbonControl();
      this.backstageViewControl1 = new DevExpress.XtraBars.Ribbon.BackstageViewControl();
      this.exitItem = new DevExpress.XtraBars.Ribbon.BackstageViewButtonItem();
      this.dashboardItem = new DevExpress.XtraBars.BarButtonItem();
      this.customersItem = new DevExpress.XtraBars.BarButtonItem();
      this.eventsItem = new DevExpress.XtraBars.BarButtonItem();
      this.roomBookingsItem = new DevExpress.XtraBars.BarButtonItem();
      this.pivotOverviewItem = new DevExpress.XtraBars.BarButtonItem();
      this.roomBookingReportItem = new DevExpress.XtraBars.BarButtonItem();
      this.barStaticItem1 = new DevExpress.XtraBars.BarStaticItem();
      this.skinRibbonGalleryBarItem1 = new DevExpress.XtraBars.SkinRibbonGalleryBarItem();
      this.skinPaletteRibbonGalleryBarItem1 = new DevExpress.XtraBars.SkinPaletteRibbonGalleryBarItem();
      this.ribbonPage1 = new DevExpress.XtraBars.Ribbon.RibbonPage();
      this.ribbonPageGroup3 = new DevExpress.XtraBars.Ribbon.RibbonPageGroup();
      this.ribbonPageGroup1 = new DevExpress.XtraBars.Ribbon.RibbonPageGroup();
      this.ribbonPageGroup2 = new DevExpress.XtraBars.Ribbon.RibbonPageGroup();
      this.ribbonPage2 = new DevExpress.XtraBars.Ribbon.RibbonPage();
      this.ribbonPageGroup4 = new DevExpress.XtraBars.Ribbon.RibbonPageGroup();
      this.ribbonStatusBar1 = new DevExpress.XtraBars.Ribbon.RibbonStatusBar();
      this.documentManager1 = new DevExpress.XtraBars.Docking2010.DocumentManager(this.components);
      this.tabbedView = new DevExpress.XtraBars.Docking2010.Views.Tabbed.TabbedView(this.components);
      ((System.ComponentModel.ISupportInitialize)(this.ribbon)).BeginInit();
      ((System.ComponentModel.ISupportInitialize)(this.backstageViewControl1)).BeginInit();
      ((System.ComponentModel.ISupportInitialize)(this.documentManager1)).BeginInit();
      ((System.ComponentModel.ISupportInitialize)(this.tabbedView)).BeginInit();
      this.SuspendLayout();
      // 
      // ribbon
      // 
      this.ribbon.ApplicationButtonDropDownControl = this.backstageViewControl1;
      this.ribbon.ExpandCollapseItem.Id = 0;
      this.ribbon.Items.AddRange(new DevExpress.XtraBars.BarItem[] {
            this.ribbon.ExpandCollapseItem,
            this.ribbon.SearchEditItem,
            this.dashboardItem,
            this.customersItem,
            this.eventsItem,
            this.roomBookingsItem,
            this.pivotOverviewItem,
            this.roomBookingReportItem,
            this.barStaticItem1,
            this.skinRibbonGalleryBarItem1,
            this.skinPaletteRibbonGalleryBarItem1});
      this.ribbon.Location = new System.Drawing.Point(0, 0);
      this.ribbon.MaxItemId = 11;
      this.ribbon.MdiMergeStyle = DevExpress.XtraBars.Ribbon.RibbonMdiMergeStyle.Always;
      this.ribbon.Name = "ribbon";
      this.ribbon.OptionsPageCategories.Alignment = DevExpress.XtraBars.Ribbon.RibbonPageCategoryAlignment.Right;
      this.ribbon.Pages.AddRange(new DevExpress.XtraBars.Ribbon.RibbonPage[] {
            this.ribbonPage1,
            this.ribbonPage2});
      this.ribbon.Size = new System.Drawing.Size(934, 158);
      this.ribbon.StatusBar = this.ribbonStatusBar1;
      this.ribbon.Merge += new DevExpress.XtraBars.Ribbon.RibbonMergeEventHandler(this.ribbon_Merge);
      // 
      // backstageViewControl1
      // 
      this.backstageViewControl1.Items.Add(this.exitItem);
      this.backstageViewControl1.Location = new System.Drawing.Point(346, 238);
      this.backstageViewControl1.Name = "backstageViewControl1";
      this.backstageViewControl1.OwnerControl = this.ribbon;
      this.backstageViewControl1.Size = new System.Drawing.Size(480, 150);
      this.backstageViewControl1.TabIndex = 10;
      this.backstageViewControl1.VisibleInDesignTime = true;
      // 
      // exitItem
      // 
      this.exitItem.Caption = "Exit";
      this.exitItem.Name = "exitItem";
      this.exitItem.ItemClick += new DevExpress.XtraBars.Ribbon.BackstageViewItemEventHandler(this.exitItem_ItemClick);
      // 
      // dashboardItem
      // 
      this.dashboardItem.Caption = "Dashboard";
      this.dashboardItem.Id = 1;
      this.dashboardItem.ImageOptions.SvgImage = global::sample_5_4.Properties.Resources.dashboard;
      this.dashboardItem.Name = "dashboardItem";
      // 
      // customersItem
      // 
      this.customersItem.Caption = "Customers";
      this.customersItem.Id = 2;
      this.customersItem.ImageOptions.SvgImage = global::sample_5_4.Properties.Resources.customers;
      this.customersItem.Name = "customersItem";
      // 
      // eventsItem
      // 
      this.eventsItem.Caption = "Events";
      this.eventsItem.Id = 4;
      this.eventsItem.ImageOptions.SvgImage = global::sample_5_4.Properties.Resources._event;
      this.eventsItem.Name = "eventsItem";
      // 
      // roomBookingsItem
      // 
      this.roomBookingsItem.Caption = "Room Bookings";
      this.roomBookingsItem.Id = 5;
      this.roomBookingsItem.ImageOptions.SvgImage = global::sample_5_4.Properties.Resources.roombooking_alt;
      this.roomBookingsItem.Name = "roomBookingsItem";
      // 
      // pivotOverviewItem
      // 
      this.pivotOverviewItem.Caption = "Pivot Overview";
      this.pivotOverviewItem.Id = 6;
      this.pivotOverviewItem.ImageOptions.SvgImage = global::sample_5_4.Properties.Resources.pivotoverview;
      this.pivotOverviewItem.Name = "pivotOverviewItem";
      // 
      // roomBookingReportItem
      // 
      this.roomBookingReportItem.Caption = "Room Booking Report";
      this.roomBookingReportItem.Id = 7;
      this.roomBookingReportItem.ImageOptions.SvgImage = global::sample_5_4.Properties.Resources.report;
      this.roomBookingReportItem.Name = "roomBookingReportItem";
      // 
      // barStaticItem1
      // 
      this.barStaticItem1.Caption = "Ready";
      this.barStaticItem1.Id = 8;
      this.barStaticItem1.Name = "barStaticItem1";
      // 
      // skinRibbonGalleryBarItem1
      // 
      this.skinRibbonGalleryBarItem1.Caption = "skinRibbonGalleryBarItem1";
      this.skinRibbonGalleryBarItem1.Id = 9;
      this.skinRibbonGalleryBarItem1.Name = "skinRibbonGalleryBarItem1";
      // 
      // skinPaletteRibbonGalleryBarItem1
      // 
      this.skinPaletteRibbonGalleryBarItem1.Caption = "skinPaletteRibbonGalleryBarItem1";
      this.skinPaletteRibbonGalleryBarItem1.Id = 10;
      this.skinPaletteRibbonGalleryBarItem1.Name = "skinPaletteRibbonGalleryBarItem1";
      // 
      // ribbonPage1
      // 
      this.ribbonPage1.Groups.AddRange(new DevExpress.XtraBars.Ribbon.RibbonPageGroup[] {
            this.ribbonPageGroup3,
            this.ribbonPageGroup1,
            this.ribbonPageGroup2});
      this.ribbonPage1.Name = "ribbonPage1";
      this.ribbonPage1.Text = "Navigation";
      // 
      // ribbonPageGroup3
      // 
      this.ribbonPageGroup3.ItemLinks.Add(this.dashboardItem);
      this.ribbonPageGroup3.Name = "ribbonPageGroup3";
      this.ribbonPageGroup3.Text = "Global";
      // 
      // ribbonPageGroup1
      // 
      this.ribbonPageGroup1.ItemLinks.Add(this.customersItem);
      this.ribbonPageGroup1.ItemLinks.Add(this.eventsItem);
      this.ribbonPageGroup1.ItemLinks.Add(this.roomBookingsItem);
      this.ribbonPageGroup1.Name = "ribbonPageGroup1";
      this.ribbonPageGroup1.Text = "Data";
      // 
      // ribbonPageGroup2
      // 
      this.ribbonPageGroup2.ItemLinks.Add(this.pivotOverviewItem);
      this.ribbonPageGroup2.ItemLinks.Add(this.roomBookingReportItem);
      this.ribbonPageGroup2.Name = "ribbonPageGroup2";
      this.ribbonPageGroup2.Text = "Data Analysis";
      // 
      // ribbonPage2
      // 
      this.ribbonPage2.Groups.AddRange(new DevExpress.XtraBars.Ribbon.RibbonPageGroup[] {
            this.ribbonPageGroup4});
      this.ribbonPage2.Name = "ribbonPage2";
      this.ribbonPage2.Text = "Style";
      // 
      // ribbonPageGroup4
      // 
      this.ribbonPageGroup4.ItemLinks.Add(this.skinRibbonGalleryBarItem1);
      this.ribbonPageGroup4.ItemLinks.Add(this.skinPaletteRibbonGalleryBarItem1);
      this.ribbonPageGroup4.Name = "ribbonPageGroup4";
      this.ribbonPageGroup4.Text = "Skins";
      // 
      // ribbonStatusBar1
      // 
      this.ribbonStatusBar1.ItemLinks.Add(this.barStaticItem1);
      this.ribbonStatusBar1.Location = new System.Drawing.Point(0, 657);
      this.ribbonStatusBar1.Name = "ribbonStatusBar1";
      this.ribbonStatusBar1.Ribbon = this.ribbon;
      this.ribbonStatusBar1.Size = new System.Drawing.Size(934, 24);
      // 
      // documentManager1
      // 
      this.documentManager1.MdiParent = this;
      this.documentManager1.MenuManager = this.ribbon;
      this.documentManager1.RibbonAndBarsMergeStyle = DevExpress.XtraBars.Docking2010.Views.RibbonAndBarsMergeStyle.Always;
      this.documentManager1.View = this.tabbedView;
      this.documentManager1.ViewCollection.AddRange(new DevExpress.XtraBars.Docking2010.Views.BaseView[] {
            this.tabbedView});
      // 
      // tabbedView
      // 
      this.tabbedView.QueryControl += new DevExpress.XtraBars.Docking2010.Views.QueryControlEventHandler(this.tabbedView_QueryControl);
      // 
      // MainForm
      // 
      this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.ClientSize = new System.Drawing.Size(934, 681);
      this.Controls.Add(this.ribbonStatusBar1);
      this.Controls.Add(this.backstageViewControl1);
      this.Controls.Add(this.ribbon);
      this.IsMdiContainer = true;
      this.Name = "MainForm";
      this.Ribbon = this.ribbon;
      this.StatusBar = this.ribbonStatusBar1;
      this.Text = "Event Management";
      this.Load += new System.EventHandler(this.MainForm_Load);
      ((System.ComponentModel.ISupportInitialize)(this.ribbon)).EndInit();
      ((System.ComponentModel.ISupportInitialize)(this.backstageViewControl1)).EndInit();
      ((System.ComponentModel.ISupportInitialize)(this.documentManager1)).EndInit();
      ((System.ComponentModel.ISupportInitialize)(this.tabbedView)).EndInit();
      this.ResumeLayout(false);
      this.PerformLayout();

    }

    #endregion
    private DevExpress.XtraBars.Ribbon.RibbonControl ribbon;
    private DevExpress.XtraBars.Ribbon.RibbonPage ribbonPage1;
    private DevExpress.XtraBars.Ribbon.RibbonPageGroup ribbonPageGroup1;
    private DevExpress.XtraBars.Ribbon.RibbonStatusBar ribbonStatusBar1;
    private DevExpress.XtraBars.Ribbon.BackstageViewControl backstageViewControl1;
    private DevExpress.XtraBars.Ribbon.BackstageViewButtonItem exitItem;
    private DevExpress.XtraBars.BarButtonItem dashboardItem;
    private DevExpress.XtraBars.BarButtonItem customersItem;
    private DevExpress.XtraBars.BarButtonItem eventsItem;
    private DevExpress.XtraBars.BarButtonItem roomBookingsItem;
    private DevExpress.XtraBars.Ribbon.RibbonPageGroup ribbonPageGroup3;
    private DevExpress.XtraBars.Ribbon.RibbonPageGroup ribbonPageGroup2;
    private DevExpress.XtraBars.BarButtonItem pivotOverviewItem;
    private DevExpress.XtraBars.BarButtonItem roomBookingReportItem;
    private DevExpress.XtraBars.BarStaticItem barStaticItem1;
    private DevExpress.XtraBars.Docking2010.DocumentManager documentManager1;
    private DevExpress.XtraBars.Docking2010.Views.Tabbed.TabbedView tabbedView;
    private DevExpress.XtraBars.SkinRibbonGalleryBarItem skinRibbonGalleryBarItem1;
    private DevExpress.XtraBars.SkinPaletteRibbonGalleryBarItem skinPaletteRibbonGalleryBarItem1;
    private DevExpress.XtraBars.Ribbon.RibbonPage ribbonPage2;
    private DevExpress.XtraBars.Ribbon.RibbonPageGroup ribbonPageGroup4;
  }
}

