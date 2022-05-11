namespace sample_5_4 {
  partial class Dashboard {
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
      DevExpress.XtraBars.Docking2010.Views.Widget.WidgetDockingContainer widgetDockingContainer1 = new DevExpress.XtraBars.Docking2010.Views.Widget.WidgetDockingContainer();
      DevExpress.XtraBars.Docking2010.Views.Widget.WidgetDockingContainer widgetDockingContainer2 = new DevExpress.XtraBars.Docking2010.Views.Widget.WidgetDockingContainer();
      DevExpress.XtraBars.Docking2010.Views.Widget.WidgetDockingContainer widgetDockingContainer3 = new DevExpress.XtraBars.Docking2010.Views.Widget.WidgetDockingContainer();
      DevExpress.XtraBars.Docking2010.Views.Widget.WidgetDockingContainer widgetDockingContainer4 = new DevExpress.XtraBars.Docking2010.Views.Widget.WidgetDockingContainer();
      DevExpress.XtraBars.Docking2010.Views.Widget.WidgetDockingContainer widgetDockingContainer5 = new DevExpress.XtraBars.Docking2010.Views.Widget.WidgetDockingContainer();
      DevExpress.XtraBars.Docking2010.Views.Widget.WidgetDockingContainer widgetDockingContainer6 = new DevExpress.XtraBars.Docking2010.Views.Widget.WidgetDockingContainer();
      this.eventsCurrentYearDocument = new DevExpress.XtraBars.Docking2010.Views.Widget.Document(this.components);
      this.roomBookingsCurrentYearDocument = new DevExpress.XtraBars.Docking2010.Views.Widget.Document(this.components);
      this.venuesDocument = new DevExpress.XtraBars.Docking2010.Views.Widget.Document(this.components);
      this.venueDetailsDocument = new DevExpress.XtraBars.Docking2010.Views.Widget.Document(this.components);
      this.roomBookingCollection = new DevExpress.Xpo.XPCollection(this.components);
      this.session = new DevExpress.Xpo.Session(this.components);
      this.eventCollection = new DevExpress.Xpo.XPCollection(this.components);
      this.singleVenueChartLink = new DevExpress.XtraPrinting.PrintableComponentLink(this.components);
      this.printExportItem = new DevExpress.XtraBars.BarButtonItem();
      this.ribbonControl1 = new DevExpress.XtraBars.Ribbon.RibbonControl();
      this.ribbonPageCategory1 = new DevExpress.XtraBars.Ribbon.RibbonPageCategory();
      this.ribbonPage1 = new DevExpress.XtraBars.Ribbon.RibbonPage();
      this.ribbonPageGroup1 = new DevExpress.XtraBars.Ribbon.RibbonPageGroup();
      this.documentManager1 = new DevExpress.XtraBars.Docking2010.DocumentManager(this.components);
      this.widgetView = new DevExpress.XtraBars.Docking2010.Views.Widget.WidgetView(this.components);
      ((System.ComponentModel.ISupportInitialize)(this.eventsCurrentYearDocument)).BeginInit();
      ((System.ComponentModel.ISupportInitialize)(this.roomBookingsCurrentYearDocument)).BeginInit();
      ((System.ComponentModel.ISupportInitialize)(this.venuesDocument)).BeginInit();
      ((System.ComponentModel.ISupportInitialize)(this.venueDetailsDocument)).BeginInit();
      ((System.ComponentModel.ISupportInitialize)(this.roomBookingCollection)).BeginInit();
      ((System.ComponentModel.ISupportInitialize)(this.session)).BeginInit();
      ((System.ComponentModel.ISupportInitialize)(this.eventCollection)).BeginInit();
      ((System.ComponentModel.ISupportInitialize)(this.ribbonControl1)).BeginInit();
      ((System.ComponentModel.ISupportInitialize)(this.documentManager1)).BeginInit();
      ((System.ComponentModel.ISupportInitialize)(this.widgetView)).BeginInit();
      this.SuspendLayout();
      // 
      // eventsCurrentYearDocument
      // 
      this.eventsCurrentYearDocument.Caption = "Events / Current Year";
      this.eventsCurrentYearDocument.ControlName = "EventsCurrentYear";
      // 
      // roomBookingsCurrentYearDocument
      // 
      this.roomBookingsCurrentYearDocument.Caption = "Room Bookings / Current Year";
      this.roomBookingsCurrentYearDocument.ColumnIndex = 1;
      this.roomBookingsCurrentYearDocument.ControlName = "RoomBookingsCurrentYear";
      // 
      // venuesDocument
      // 
      this.venuesDocument.Caption = "Venues";
      this.venuesDocument.ControlName = "Venues";
      this.venuesDocument.RowIndex = 1;
      // 
      // venueDetailsDocument
      // 
      this.venueDetailsDocument.Caption = "Venue Details";
      this.venueDetailsDocument.ColumnIndex = 1;
      this.venueDetailsDocument.ControlName = "VenueDetails";
      this.venueDetailsDocument.RowIndex = 1;
      // 
      // roomBookingCollection
      // 
      this.roomBookingCollection.ObjectType = typeof(EventsDB.RoomBooking);
      this.roomBookingCollection.Session = this.session;
      // 
      // eventCollection
      // 
      this.eventCollection.ObjectType = typeof(EventsDB.Event);
      this.eventCollection.Session = this.session;
      // 
      // singleVenueChartLink
      // 
      this.singleVenueChartLink.PaperKind = System.Drawing.Printing.PaperKind.A4;
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
            this.ribbonControl1.SearchEditItem,
            this.printExportItem});
      this.ribbonControl1.Location = new System.Drawing.Point(0, 0);
      this.ribbonControl1.MaxItemId = 1;
      this.ribbonControl1.Name = "ribbonControl1";
      this.ribbonControl1.PageCategories.AddRange(new DevExpress.XtraBars.Ribbon.RibbonPageCategory[] {
            this.ribbonPageCategory1});
      this.ribbonControl1.Size = new System.Drawing.Size(931, 150);
      // 
      // ribbonPageCategory1
      // 
      this.ribbonPageCategory1.Appearance.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
      this.ribbonPageCategory1.Appearance.Options.UseBackColor = true;
      this.ribbonPageCategory1.Name = "ribbonPageCategory1";
      this.ribbonPageCategory1.Pages.AddRange(new DevExpress.XtraBars.Ribbon.RibbonPage[] {
            this.ribbonPage1});
      this.ribbonPageCategory1.Text = "Dashboard";
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
      // documentManager1
      // 
      this.documentManager1.ContainerControl = this;
      this.documentManager1.MenuManager = this.ribbonControl1;
      this.documentManager1.View = this.widgetView;
      this.documentManager1.ViewCollection.AddRange(new DevExpress.XtraBars.Docking2010.Views.BaseView[] {
            this.widgetView});
      // 
      // widgetView
      // 
      this.widgetView.Documents.AddRange(new DevExpress.XtraBars.Docking2010.Views.BaseDocument[] {
            this.eventsCurrentYearDocument,
            this.roomBookingsCurrentYearDocument,
            this.venuesDocument,
            this.venueDetailsDocument});
      this.widgetView.FreeLayoutProperties.FreeLayoutItems.AddRange(new DevExpress.XtraBars.Docking2010.Views.Widget.Document[] {
            this.eventsCurrentYearDocument,
            this.roomBookingsCurrentYearDocument,
            this.venuesDocument,
            this.venueDetailsDocument});
      this.widgetView.LayoutMode = DevExpress.XtraBars.Docking2010.Views.Widget.LayoutMode.FreeLayout;
      widgetDockingContainer2.Element = this.eventsCurrentYearDocument;
      widgetDockingContainer3.Element = this.roomBookingsCurrentYearDocument;
      widgetDockingContainer1.Nodes.AddRange(new DevExpress.XtraBars.Docking2010.Views.Tabbed.DockingContainer[] {
            widgetDockingContainer2,
            widgetDockingContainer3});
      widgetDockingContainer5.Element = this.venuesDocument;
      widgetDockingContainer6.Element = this.venueDetailsDocument;
      widgetDockingContainer4.Nodes.AddRange(new DevExpress.XtraBars.Docking2010.Views.Tabbed.DockingContainer[] {
            widgetDockingContainer5,
            widgetDockingContainer6});
      this.widgetView.RootContainer.Nodes.AddRange(new DevExpress.XtraBars.Docking2010.Views.Tabbed.DockingContainer[] {
            widgetDockingContainer1,
            widgetDockingContainer4});
      this.widgetView.RootContainer.Orientation = System.Windows.Forms.Orientation.Vertical;
      this.widgetView.QueryControl += new DevExpress.XtraBars.Docking2010.Views.QueryControlEventHandler(this.widgetView_QueryControl);
      // 
      // Dashboard
      // 
      this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.Controls.Add(this.ribbonControl1);
      this.Name = "Dashboard";
      this.Size = new System.Drawing.Size(931, 686);
      ((System.ComponentModel.ISupportInitialize)(this.eventsCurrentYearDocument)).EndInit();
      ((System.ComponentModel.ISupportInitialize)(this.roomBookingsCurrentYearDocument)).EndInit();
      ((System.ComponentModel.ISupportInitialize)(this.venuesDocument)).EndInit();
      ((System.ComponentModel.ISupportInitialize)(this.venueDetailsDocument)).EndInit();
      ((System.ComponentModel.ISupportInitialize)(this.roomBookingCollection)).EndInit();
      ((System.ComponentModel.ISupportInitialize)(this.session)).EndInit();
      ((System.ComponentModel.ISupportInitialize)(this.eventCollection)).EndInit();
      ((System.ComponentModel.ISupportInitialize)(this.ribbonControl1)).EndInit();
      ((System.ComponentModel.ISupportInitialize)(this.documentManager1)).EndInit();
      ((System.ComponentModel.ISupportInitialize)(this.widgetView)).EndInit();
      this.ResumeLayout(false);
      this.PerformLayout();

    }

    #endregion
    private DevExpress.Xpo.Session session;
    private DevExpress.Xpo.XPCollection eventCollection;
    private DevExpress.Xpo.XPCollection roomBookingCollection;
    private DevExpress.XtraPrinting.PrintableComponentLink singleVenueChartLink;
    private DevExpress.XtraBars.BarButtonItem printExportItem;
    private DevExpress.XtraBars.Ribbon.RibbonControl ribbonControl1;
    private DevExpress.XtraBars.Ribbon.RibbonPage ribbonPage1;
    private DevExpress.XtraBars.Ribbon.RibbonPageGroup ribbonPageGroup1;
    private DevExpress.XtraBars.Ribbon.RibbonPageCategory ribbonPageCategory1;
    private DevExpress.XtraBars.Docking2010.DocumentManager documentManager1;
    private DevExpress.XtraBars.Docking2010.Views.Widget.WidgetView widgetView;
    private DevExpress.XtraBars.Docking2010.Views.Widget.Document eventsCurrentYearDocument;
    private DevExpress.XtraBars.Docking2010.Views.Widget.Document roomBookingsCurrentYearDocument;
    private DevExpress.XtraBars.Docking2010.Views.Widget.Document venuesDocument;
    private DevExpress.XtraBars.Docking2010.Views.Widget.Document venueDetailsDocument;
  }
}