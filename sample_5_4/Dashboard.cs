using DevExpress.XtraEditors;
using sample_5_4.DashboardControls;
using System.Windows.Forms;

namespace sample_5_4 {
  public partial class Dashboard : XtraUserControl {
    public Dashboard() {
      InitializeComponent();
    }

    private void widgetView_QueryControl(object sender, DevExpress.XtraBars.Docking2010.Views.QueryControlEventArgs e) {
      e.Control = e.Document.ControlName switch
      {
        "EventsCurrentYear" => new EventsCurrentYear(),
        "RoomBookingsCurrentYear" => new RoomBookingsCurrentYear(),
        "VenueDetails" => new VenueDetails(),
        "Venues" => new Venues(),
        _ => new Control()
      };
    }
  }
}
