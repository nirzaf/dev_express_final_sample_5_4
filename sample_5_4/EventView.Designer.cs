namespace sample_5_4 {
  partial class EventView {
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
      DevExpress.XtraScheduler.TimeRuler timeRuler1 = new DevExpress.XtraScheduler.TimeRuler();
      DevExpress.XtraScheduler.TimeRuler timeRuler2 = new DevExpress.XtraScheduler.TimeRuler();
      DevExpress.XtraScheduler.TimeRuler timeRuler3 = new DevExpress.XtraScheduler.TimeRuler();
      this.commitChangesItem = new DevExpress.XtraBars.BarButtonItem();
      this.discardChangesItem = new DevExpress.XtraBars.BarButtonItem();
      this.printExportItem = new DevExpress.XtraBars.BarButtonItem();
      this.storage = new DevExpress.XtraScheduler.SchedulerStorage(this.components);
      this.dayScrollerPanel = new DevExpress.XtraEditors.PanelControl();
      this.dayScrollbar = new DevExpress.XtraEditors.HScrollBar();
      this.eventCollection = new DevExpress.Xpo.XPCollection(this.components);
      this.unitOfWork = new DevExpress.Xpo.UnitOfWork(this.components);
      this.splitContainerControl1 = new DevExpress.XtraEditors.SplitContainerControl();
      this.eventGrid = new DevExpress.XtraGrid.GridControl();
      this.eventGridView = new DevExpress.XtraGrid.Views.Grid.GridView();
      this.colName = new DevExpress.XtraGrid.Columns.GridColumn();
      this.colStartDate = new DevExpress.XtraGrid.Columns.GridColumn();
      this.colEndDate = new DevExpress.XtraGrid.Columns.GridColumn();
      this.gridColumn1 = new DevExpress.XtraGrid.Columns.GridColumn();
      this.gridColumn3 = new DevExpress.XtraGrid.Columns.GridColumn();
      this.ribbonControl1 = new DevExpress.XtraBars.Ribbon.RibbonControl();
      this.ribbonPageCategory1 = new DevExpress.XtraBars.Ribbon.RibbonPageCategory();
      this.ribbonPage1 = new DevExpress.XtraBars.Ribbon.RibbonPage();
      this.ribbonPageGroup2 = new DevExpress.XtraBars.Ribbon.RibbonPageGroup();
      this.ribbonPageGroup1 = new DevExpress.XtraBars.Ribbon.RibbonPageGroup();
      this.scheduler = new DevExpress.XtraScheduler.SchedulerControl();
      this.customerCollection = new DevExpress.Xpo.XPCollection(this.components);
      this.venueCollection = new DevExpress.Xpo.XPCollection(this.components);
      ((System.ComponentModel.ISupportInitialize)(this.storage)).BeginInit();
      ((System.ComponentModel.ISupportInitialize)(this.dayScrollerPanel)).BeginInit();
      this.dayScrollerPanel.SuspendLayout();
      ((System.ComponentModel.ISupportInitialize)(this.eventCollection)).BeginInit();
      ((System.ComponentModel.ISupportInitialize)(this.unitOfWork)).BeginInit();
      ((System.ComponentModel.ISupportInitialize)(this.splitContainerControl1)).BeginInit();
      this.splitContainerControl1.SuspendLayout();
      ((System.ComponentModel.ISupportInitialize)(this.eventGrid)).BeginInit();
      ((System.ComponentModel.ISupportInitialize)(this.eventGridView)).BeginInit();
      ((System.ComponentModel.ISupportInitialize)(this.ribbonControl1)).BeginInit();
      ((System.ComponentModel.ISupportInitialize)(this.scheduler)).BeginInit();
      ((System.ComponentModel.ISupportInitialize)(this.customerCollection)).BeginInit();
      ((System.ComponentModel.ISupportInitialize)(this.venueCollection)).BeginInit();
      this.SuspendLayout();
      // 
      // commitChangesItem
      // 
      this.commitChangesItem.Caption = "Save Changes";
      this.commitChangesItem.Id = 0;
      this.commitChangesItem.ImageOptions.SvgImage = global::sample_5_4.Properties.Resources.savechanges;
      this.commitChangesItem.Name = "commitChangesItem";
      // 
      // discardChangesItem
      // 
      this.discardChangesItem.Caption = "Discard Changes";
      this.discardChangesItem.Id = 1;
      this.discardChangesItem.ImageOptions.SvgImage = global::sample_5_4.Properties.Resources.discardchanges;
      this.discardChangesItem.Name = "discardChangesItem";
      // 
      // printExportItem
      // 
      this.printExportItem.Caption = "Print && Export";
      this.printExportItem.Id = 2;
      this.printExportItem.ImageOptions.SvgImage = global::sample_5_4.Properties.Resources.printexport;
      this.printExportItem.Name = "printExportItem";
      // 
      // storage
      // 
      this.storage.Appointments.CustomFieldMappings.Add(new DevExpress.XtraScheduler.AppointmentCustomFieldMapping("CustomCapacity", "Capacity", DevExpress.XtraScheduler.FieldValueType.Integer));
      this.storage.Appointments.CustomFieldMappings.Add(new DevExpress.XtraScheduler.AppointmentCustomFieldMapping("CustomRoomSetup", "RoomSetup"));
      this.storage.Appointments.Mappings.AppointmentId = "Oid";
      this.storage.Appointments.Mappings.End = "EndTime";
      this.storage.Appointments.Mappings.Start = "StartTime";
      this.storage.Appointments.Mappings.Subject = "Name";
      this.storage.EnableReminders = false;
      // 
      // dayScrollerPanel
      // 
      this.dayScrollerPanel.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder;
      this.dayScrollerPanel.Controls.Add(this.dayScrollbar);
      this.dayScrollerPanel.Dock = System.Windows.Forms.DockStyle.Bottom;
      this.dayScrollerPanel.Location = new System.Drawing.Point(0, 459);
      this.dayScrollerPanel.Margin = new System.Windows.Forms.Padding(0);
      this.dayScrollerPanel.Name = "dayScrollerPanel";
      this.dayScrollerPanel.Padding = new System.Windows.Forms.Padding(0, 0, 19, 0);
      this.dayScrollerPanel.Size = new System.Drawing.Size(1115, 16);
      this.dayScrollerPanel.TabIndex = 2;
      // 
      // dayScrollbar
      // 
      this.dayScrollbar.Dock = System.Windows.Forms.DockStyle.Fill;
      this.dayScrollbar.LargeChange = 1;
      this.dayScrollbar.Location = new System.Drawing.Point(0, 0);
      this.dayScrollbar.Margin = new System.Windows.Forms.Padding(0);
      this.dayScrollbar.Maximum = 10;
      this.dayScrollbar.Name = "dayScrollbar";
      this.dayScrollbar.Size = new System.Drawing.Size(1096, 16);
      this.dayScrollbar.TabIndex = 1;
      this.dayScrollbar.Scroll += new System.Windows.Forms.ScrollEventHandler(this.dayScrollbar_Scroll);
      // 
      // eventCollection
      // 
      this.eventCollection.ObjectType = typeof(EventsDB.Event);
      this.eventCollection.Session = this.unitOfWork;
      // 
      // splitContainerControl1
      // 
      this.splitContainerControl1.Dock = System.Windows.Forms.DockStyle.Fill;
      this.splitContainerControl1.Horizontal = false;
      this.splitContainerControl1.Location = new System.Drawing.Point(0, 150);
      this.splitContainerControl1.Name = "splitContainerControl1";
      this.splitContainerControl1.Panel1.Controls.Add(this.eventGrid);
      this.splitContainerControl1.Panel1.Text = "Panel1";
      this.splitContainerControl1.Panel2.Controls.Add(this.scheduler);
      this.splitContainerControl1.Panel2.Controls.Add(this.dayScrollerPanel);
      this.splitContainerControl1.Panel2.Text = "Panel2";
      this.splitContainerControl1.Size = new System.Drawing.Size(1115, 608);
      this.splitContainerControl1.SplitterPosition = 123;
      this.splitContainerControl1.TabIndex = 4;
      this.splitContainerControl1.Text = "splitContainerControl1";
      // 
      // eventGrid
      // 
      this.eventGrid.DataSource = this.eventCollection;
      this.eventGrid.Dock = System.Windows.Forms.DockStyle.Fill;
      this.eventGrid.Location = new System.Drawing.Point(0, 0);
      this.eventGrid.MainView = this.eventGridView;
      this.eventGrid.MenuManager = this.ribbonControl1;
      this.eventGrid.Name = "eventGrid";
      this.eventGrid.Size = new System.Drawing.Size(1115, 123);
      this.eventGrid.TabIndex = 0;
      this.eventGrid.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.eventGridView});
      this.eventGrid.Load += new System.EventHandler(this.eventGrid_Load);
      // 
      // eventGridView
      // 
      this.eventGridView.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.colName,
            this.colStartDate,
            this.colEndDate,
            this.gridColumn1,
            this.gridColumn3});
      this.eventGridView.GridControl = this.eventGrid;
      this.eventGridView.Name = "eventGridView";
      this.eventGridView.OptionsBehavior.Editable = false;
      this.eventGridView.OptionsSelection.EnableAppearanceFocusedCell = false;
      this.eventGridView.OptionsView.ShowGroupPanel = false;
      this.eventGridView.SortInfo.AddRange(new DevExpress.XtraGrid.Columns.GridColumnSortInfo[] {
            new DevExpress.XtraGrid.Columns.GridColumnSortInfo(this.colStartDate, DevExpress.Data.ColumnSortOrder.Ascending)});
      // 
      // colName
      // 
      this.colName.FieldName = "Name";
      this.colName.Name = "colName";
      this.colName.Visible = true;
      this.colName.VisibleIndex = 0;
      // 
      // colStartDate
      // 
      this.colStartDate.FieldName = "StartDate";
      this.colStartDate.Name = "colStartDate";
      this.colStartDate.Visible = true;
      this.colStartDate.VisibleIndex = 1;
      // 
      // colEndDate
      // 
      this.colEndDate.FieldName = "EndDate";
      this.colEndDate.Name = "colEndDate";
      this.colEndDate.Visible = true;
      this.colEndDate.VisibleIndex = 2;
      // 
      // gridColumn1
      // 
      this.gridColumn1.Caption = "Venue";
      this.gridColumn1.FieldName = "Venue.Name";
      this.gridColumn1.Name = "gridColumn1";
      this.gridColumn1.Visible = true;
      this.gridColumn1.VisibleIndex = 3;
      // 
      // gridColumn3
      // 
      this.gridColumn3.Caption = "Customer";
      this.gridColumn3.FieldName = "Customer.Name";
      this.gridColumn3.Name = "gridColumn3";
      this.gridColumn3.Visible = true;
      this.gridColumn3.VisibleIndex = 4;
      // 
      // ribbonControl1
      // 
      this.ribbonControl1.ExpandCollapseItem.Id = 0;
      this.ribbonControl1.Items.AddRange(new DevExpress.XtraBars.BarItem[] {
            this.ribbonControl1.ExpandCollapseItem,
            this.ribbonControl1.SearchEditItem,
            this.commitChangesItem,
            this.discardChangesItem,
            this.printExportItem});
      this.ribbonControl1.Location = new System.Drawing.Point(0, 0);
      this.ribbonControl1.MaxItemId = 1;
      this.ribbonControl1.Name = "ribbonControl1";
      this.ribbonControl1.PageCategories.AddRange(new DevExpress.XtraBars.Ribbon.RibbonPageCategory[] {
            this.ribbonPageCategory1});
      this.ribbonControl1.Size = new System.Drawing.Size(1115, 150);
      // 
      // ribbonPageCategory1
      // 
      this.ribbonPageCategory1.Appearance.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
      this.ribbonPageCategory1.Appearance.Options.UseBackColor = true;
      this.ribbonPageCategory1.Name = "ribbonPageCategory1";
      this.ribbonPageCategory1.Pages.AddRange(new DevExpress.XtraBars.Ribbon.RibbonPage[] {
            this.ribbonPage1});
      this.ribbonPageCategory1.Text = "Events";
      // 
      // ribbonPage1
      // 
      this.ribbonPage1.Groups.AddRange(new DevExpress.XtraBars.Ribbon.RibbonPageGroup[] {
            this.ribbonPageGroup2,
            this.ribbonPageGroup1});
      this.ribbonPage1.Name = "ribbonPage1";
      this.ribbonPage1.Text = "Data Operations";
      // 
      // ribbonPageGroup2
      // 
      this.ribbonPageGroup2.ItemLinks.Add(this.commitChangesItem);
      this.ribbonPageGroup2.ItemLinks.Add(this.discardChangesItem);
      this.ribbonPageGroup2.Name = "ribbonPageGroup2";
      this.ribbonPageGroup2.Text = "Persistence";
      // 
      // ribbonPageGroup1
      // 
      this.ribbonPageGroup1.ItemLinks.Add(this.printExportItem, true);
      this.ribbonPageGroup1.Name = "ribbonPageGroup1";
      this.ribbonPageGroup1.Text = "Tools";
      // 
      // scheduler
      // 
      this.scheduler.DataStorage = this.storage;
      this.scheduler.Dock = System.Windows.Forms.DockStyle.Fill;
      this.scheduler.Location = new System.Drawing.Point(0, 0);
      this.scheduler.MenuManager = this.ribbonControl1;
      this.scheduler.Name = "scheduler";
      this.scheduler.Size = new System.Drawing.Size(1115, 459);
      this.scheduler.Start = new System.DateTime(2020, 7, 16, 0, 0, 0, 0);
      this.scheduler.TabIndex = 3;
      this.scheduler.Text = "schedulerControl1";
      this.scheduler.Views.AgendaView.Enabled = false;
      this.scheduler.Views.DayView.AppointmentDisplayOptions.SnapToCellsMode = DevExpress.XtraScheduler.AppointmentSnapToCellsMode.Never;
      this.scheduler.Views.DayView.NavigationButtonVisibility = DevExpress.XtraScheduler.NavigationButtonVisibility.Never;
      this.scheduler.Views.DayView.ShowAllDayArea = false;
      this.scheduler.Views.DayView.TimeMarkerVisibility = DevExpress.XtraScheduler.TimeMarkerVisibility.Never;
      this.scheduler.Views.DayView.TimeRulers.Add(timeRuler1);
      this.scheduler.Views.DayView.TimeScale = System.TimeSpan.Parse("01:00:00");
      this.scheduler.Views.FullWeekView.TimeRulers.Add(timeRuler2);
      this.scheduler.Views.GanttView.Enabled = false;
      this.scheduler.Views.MonthView.Enabled = false;
      this.scheduler.Views.TimelineView.Enabled = false;
      this.scheduler.Views.WeekView.Enabled = false;
      this.scheduler.Views.WorkWeekView.Enabled = false;
      this.scheduler.Views.WorkWeekView.TimeRulers.Add(timeRuler3);
      this.scheduler.VisibleIntervalChanged += new System.EventHandler(this.scheduler_VisibleIntervalChanged);
      this.scheduler.InitAppointmentDisplayText += new DevExpress.XtraScheduler.AppointmentDisplayTextEventHandler(this.scheduler_InitAppointmentDisplayText);
      this.scheduler.AppointmentViewInfoCustomizing += new DevExpress.XtraScheduler.AppointmentViewInfoCustomizingEventHandler(this.scheduler_AppointmentViewInfoCustomizing);
      this.scheduler.PopupMenuShowing += new DevExpress.XtraScheduler.PopupMenuShowingEventHandler(this.scheduler_PopupMenuShowing);
      this.scheduler.EditAppointmentFormShowing += new DevExpress.XtraScheduler.AppointmentFormEventHandler(this.scheduler_EditAppointmentFormShowing);
      // 
      // customerCollection
      // 
      this.customerCollection.ObjectType = typeof(EventsDB.Customer);
      this.customerCollection.Session = this.unitOfWork;
      // 
      // venueCollection
      // 
      this.venueCollection.ObjectType = typeof(EventsDB.Venue);
      this.venueCollection.Session = this.unitOfWork;
      // 
      // EventView
      // 
      this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.Controls.Add(this.splitContainerControl1);
      this.Controls.Add(this.ribbonControl1);
      this.Name = "EventView";
      this.Size = new System.Drawing.Size(1115, 758);
      this.Load += new System.EventHandler(this.EventView_Load);
      ((System.ComponentModel.ISupportInitialize)(this.storage)).EndInit();
      ((System.ComponentModel.ISupportInitialize)(this.dayScrollerPanel)).EndInit();
      this.dayScrollerPanel.ResumeLayout(false);
      ((System.ComponentModel.ISupportInitialize)(this.eventCollection)).EndInit();
      ((System.ComponentModel.ISupportInitialize)(this.unitOfWork)).EndInit();
      ((System.ComponentModel.ISupportInitialize)(this.splitContainerControl1)).EndInit();
      this.splitContainerControl1.ResumeLayout(false);
      ((System.ComponentModel.ISupportInitialize)(this.eventGrid)).EndInit();
      ((System.ComponentModel.ISupportInitialize)(this.eventGridView)).EndInit();
      ((System.ComponentModel.ISupportInitialize)(this.ribbonControl1)).EndInit();
      ((System.ComponentModel.ISupportInitialize)(this.scheduler)).EndInit();
      ((System.ComponentModel.ISupportInitialize)(this.customerCollection)).EndInit();
      ((System.ComponentModel.ISupportInitialize)(this.venueCollection)).EndInit();
      this.ResumeLayout(false);
      this.PerformLayout();

    }

    #endregion

    private DevExpress.XtraScheduler.SchedulerStorage storage;
    private DevExpress.Xpo.XPCollection eventCollection;
    private DevExpress.Xpo.UnitOfWork unitOfWork;
    private DevExpress.XtraEditors.SplitContainerControl splitContainerControl1;
    private DevExpress.Xpo.XPCollection customerCollection;
    private DevExpress.Xpo.XPCollection venueCollection;
    private DevExpress.XtraBars.BarButtonItem commitChangesItem;
    private DevExpress.XtraBars.BarButtonItem discardChangesItem;
    private DevExpress.XtraEditors.HScrollBar dayScrollbar;
    private DevExpress.XtraEditors.PanelControl dayScrollerPanel;
    private DevExpress.XtraBars.BarButtonItem printExportItem;
    private DevExpress.XtraBars.Ribbon.RibbonControl ribbonControl1;
    private DevExpress.XtraBars.Ribbon.RibbonPage ribbonPage1;
    private DevExpress.XtraBars.Ribbon.RibbonPageGroup ribbonPageGroup1;
    private DevExpress.XtraBars.Ribbon.RibbonPageCategory ribbonPageCategory1;
    private DevExpress.XtraBars.Ribbon.RibbonPageGroup ribbonPageGroup2;
    private DevExpress.XtraGrid.GridControl eventGrid;
    private DevExpress.XtraGrid.Views.Grid.GridView eventGridView;
    private DevExpress.XtraGrid.Columns.GridColumn colName;
    private DevExpress.XtraGrid.Columns.GridColumn colStartDate;
    private DevExpress.XtraGrid.Columns.GridColumn colEndDate;
    private DevExpress.XtraGrid.Columns.GridColumn gridColumn1;
    private DevExpress.XtraGrid.Columns.GridColumn gridColumn3;
    private DevExpress.XtraScheduler.SchedulerControl scheduler;
  }
}