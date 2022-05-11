namespace sample_5_4 {
  partial class RoomBookingView {
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
      this.saveChangesItem = new DevExpress.XtraBars.BarButtonItem();
      this.discardChangesItem = new DevExpress.XtraBars.BarButtonItem();
      this.unitOfWork = new DevExpress.Xpo.UnitOfWork(this.components);
      this.splitContainerControl1 = new DevExpress.XtraEditors.SplitContainerControl();
      this.splitContainerControl2 = new DevExpress.XtraEditors.SplitContainerControl();
      this.venueCollection = new DevExpress.Xpo.XPCollection(this.components);
      this.schedulerStorage = new DevExpress.XtraScheduler.SchedulerStorage(this.components);
      this.ribbonControl1 = new DevExpress.XtraBars.Ribbon.RibbonControl();
      this.ribbonPageCategory1 = new DevExpress.XtraBars.Ribbon.RibbonPageCategory();
      this.ribbonPage1 = new DevExpress.XtraBars.Ribbon.RibbonPage();
      this.ribbonPageGroup1 = new DevExpress.XtraBars.Ribbon.RibbonPageGroup();
      ((System.ComponentModel.ISupportInitialize)(this.unitOfWork)).BeginInit();
      ((System.ComponentModel.ISupportInitialize)(this.splitContainerControl1)).BeginInit();
      this.splitContainerControl1.SuspendLayout();
      ((System.ComponentModel.ISupportInitialize)(this.splitContainerControl2)).BeginInit();
      this.splitContainerControl2.SuspendLayout();
      ((System.ComponentModel.ISupportInitialize)(this.venueCollection)).BeginInit();
      ((System.ComponentModel.ISupportInitialize)(this.schedulerStorage)).BeginInit();
      ((System.ComponentModel.ISupportInitialize)(this.ribbonControl1)).BeginInit();
      this.SuspendLayout();
      // 
      // saveChangesItem
      // 
      this.saveChangesItem.Caption = "Save Changes";
      this.saveChangesItem.Id = 0;
      this.saveChangesItem.ImageOptions.SvgImage = global::sample_5_4.Properties.Resources.savechanges;
      this.saveChangesItem.Name = "saveChangesItem";
      // 
      // discardChangesItem
      // 
      this.discardChangesItem.Caption = "Discard Changes";
      this.discardChangesItem.Id = 1;
      this.discardChangesItem.ImageOptions.SvgImage = global::sample_5_4.Properties.Resources.discardchanges;
      this.discardChangesItem.Name = "discardChangesItem";
      // 
      // splitContainerControl1
      // 
      this.splitContainerControl1.Dock = System.Windows.Forms.DockStyle.Fill;
      this.splitContainerControl1.Horizontal = false;
      this.splitContainerControl1.Location = new System.Drawing.Point(0, 150);
      this.splitContainerControl1.Name = "splitContainerControl1";
      this.splitContainerControl1.Panel1.Text = "Panel1";
      this.splitContainerControl1.Panel2.Controls.Add(this.splitContainerControl2);
      this.splitContainerControl1.Panel2.Text = "Panel2";
      this.splitContainerControl1.Size = new System.Drawing.Size(607, 210);
      this.splitContainerControl1.SplitterPosition = 95;
      this.splitContainerControl1.TabIndex = 4;
      this.splitContainerControl1.Text = "splitContainerControl1";
      // 
      // splitContainerControl2
      // 
      this.splitContainerControl2.Dock = System.Windows.Forms.DockStyle.Fill;
      this.splitContainerControl2.FixedPanel = DevExpress.XtraEditors.SplitFixedPanel.Panel2;
      this.splitContainerControl2.Location = new System.Drawing.Point(0, 0);
      this.splitContainerControl2.Name = "splitContainerControl2";
      this.splitContainerControl2.Panel1.Text = "Panel1";
      this.splitContainerControl2.Panel2.Text = "Panel2";
      this.splitContainerControl2.Size = new System.Drawing.Size(607, 105);
      this.splitContainerControl2.SplitterPosition = 178;
      this.splitContainerControl2.TabIndex = 0;
      this.splitContainerControl2.Text = "splitContainerControl2";
      // 
      // venueCollection
      // 
      this.venueCollection.ObjectType = typeof(EventsDB.Venue);
      this.venueCollection.Session = this.unitOfWork;
      // 
      // schedulerStorage
      // 
      this.schedulerStorage.Appointments.CustomFieldMappings.Add(new DevExpress.XtraScheduler.AppointmentCustomFieldMapping("CustomCapReq", "CapacityRequirement!"));
      this.schedulerStorage.Appointments.Mappings.End = "EndTime";
      this.schedulerStorage.Appointments.Mappings.ResourceId = "Room!Key";
      this.schedulerStorage.Appointments.Mappings.Start = "StartTime";
      this.schedulerStorage.EnableReminders = false;
      this.schedulerStorage.Resources.Mappings.Caption = "Name";
      this.schedulerStorage.Resources.Mappings.Id = "Oid";
      // 
      // ribbonControl1
      // 
      this.ribbonControl1.ExpandCollapseItem.Id = 0;
      this.ribbonControl1.Items.AddRange(new DevExpress.XtraBars.BarItem[] {
            this.ribbonControl1.ExpandCollapseItem,
            this.ribbonControl1.SearchEditItem,
            this.saveChangesItem,
            this.discardChangesItem});
      this.ribbonControl1.Location = new System.Drawing.Point(0, 0);
      this.ribbonControl1.MaxItemId = 1;
      this.ribbonControl1.Name = "ribbonControl1";
      this.ribbonControl1.PageCategories.AddRange(new DevExpress.XtraBars.Ribbon.RibbonPageCategory[] {
            this.ribbonPageCategory1});
      this.ribbonControl1.Size = new System.Drawing.Size(607, 150);
      // 
      // ribbonPageCategory1
      // 
      this.ribbonPageCategory1.Appearance.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
      this.ribbonPageCategory1.Appearance.Options.UseBackColor = true;
      this.ribbonPageCategory1.Name = "ribbonPageCategory1";
      this.ribbonPageCategory1.Pages.AddRange(new DevExpress.XtraBars.Ribbon.RibbonPage[] {
            this.ribbonPage1});
      this.ribbonPageCategory1.Text = "Room Bookings";
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
      this.ribbonPageGroup1.ItemLinks.Add(this.saveChangesItem);
      this.ribbonPageGroup1.ItemLinks.Add(this.discardChangesItem);
      this.ribbonPageGroup1.Name = "ribbonPageGroup1";
      this.ribbonPageGroup1.Text = "Persistence";
      // 
      // RoomBookingsView
      // 
      this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.Controls.Add(this.splitContainerControl1);
      this.Controls.Add(this.ribbonControl1);
      this.Name = "RoomBookingsView";
      this.Size = new System.Drawing.Size(607, 360);
      this.Load += new System.EventHandler(this.RoomBookingView_Load);
      ((System.ComponentModel.ISupportInitialize)(this.unitOfWork)).EndInit();
      ((System.ComponentModel.ISupportInitialize)(this.splitContainerControl1)).EndInit();
      this.splitContainerControl1.ResumeLayout(false);
      ((System.ComponentModel.ISupportInitialize)(this.splitContainerControl2)).EndInit();
      this.splitContainerControl2.ResumeLayout(false);
      ((System.ComponentModel.ISupportInitialize)(this.venueCollection)).EndInit();
      ((System.ComponentModel.ISupportInitialize)(this.schedulerStorage)).EndInit();
      ((System.ComponentModel.ISupportInitialize)(this.ribbonControl1)).EndInit();
      this.ResumeLayout(false);
      this.PerformLayout();

    }

    #endregion
    private DevExpress.Xpo.UnitOfWork unitOfWork;
    private DevExpress.XtraEditors.SplitContainerControl splitContainerControl1;
    private DevExpress.Xpo.XPCollection venueCollection;
    private DevExpress.XtraBars.BarButtonItem saveChangesItem;
    private DevExpress.XtraBars.BarButtonItem discardChangesItem;
    private DevExpress.XtraEditors.SplitContainerControl splitContainerControl2;
    private DevExpress.XtraScheduler.SchedulerStorage schedulerStorage;
    private DevExpress.XtraBars.Ribbon.RibbonControl ribbonControl1;
    private DevExpress.XtraBars.Ribbon.RibbonPage ribbonPage1;
    private DevExpress.XtraBars.Ribbon.RibbonPageGroup ribbonPageGroup1;
    private DevExpress.XtraBars.Ribbon.RibbonPageCategory ribbonPageCategory1;
  }
}