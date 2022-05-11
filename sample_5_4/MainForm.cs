using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DevExpress.XtraBars;
using DevExpress.XtraBars.Docking2010.Views;

namespace sample_5_4 {
  public partial class MainForm : DevExpress.XtraBars.Ribbon.RibbonForm {
    public MainForm() {
      InitializeComponent();
    }

    private void MainForm_Load(object sender, EventArgs e) {
      var buttonViews = new Dictionary<BarItem, ShowOrActivateDelegate> {
        { dashboardItem, ShowOrActivate<Dashboard>("Dashboard") },
        { customersItem, ShowOrActivate<CustomerList>("Customers") },
        { eventsItem, ShowOrActivate<EventView>("Events") },
        { roomBookingsItem, ShowOrActivate<RoomBookingView>("Room Bookings") },
        { pivotOverviewItem, ShowOrActivate<PivotOverview>("Pivot Overview") }
      };

      foreach (var barItem in buttonViews.Keys) {
        barItem.ItemClick += (s, ev) => {
          buttonViews[barItem]();
        };
      }
    }

    delegate BaseDocument ShowOrActivateDelegate(Action<BaseDocument> initDocument = null);

    ShowOrActivateDelegate ShowOrActivate<T>(string caption) =>
      (Action<BaseDocument> initDocument) => {
        var controlName = typeof(T).Name;
        var doc = tabbedView.Documents.FirstOrDefault(d => d.ControlName == controlName && d.Caption == caption);
        if (doc == null) {
          doc = tabbedView.AddDocument(caption, controlName);
          initDocument?.Invoke(doc);
        }
        if (doc != null)
          tabbedView.Controller.Activate(doc);
        return doc;
      };

    private void tabbedView_QueryControl(object sender, QueryControlEventArgs e) {
      e.Control = e.Document.ControlName switch
      {
        "Dashboard" => new Dashboard(),
        "CustomerList" => new CustomerList(),
        "EventView" => new EventView(),
        "RoomBookingView" => new RoomBookingView(),
        "PivotOverview" => new PivotOverview(),
        _ => new Control()
      };
    }

    private void exitItem_ItemClick(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e) {
      Close();
    }

    private void ribbon_Merge(object sender, DevExpress.XtraBars.Ribbon.RibbonMergeEventArgs e) {
      if (e.MergeOwner.MergedCategories.Count > 0) {
        var page = e.MergeOwner.MergedCategories[0].Pages[0];
        if (page != null)
          ribbon.SelectedPage = page;
      }
    }
  }
}
