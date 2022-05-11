using DevExpress.XtraEditors;
using EventsDB;
using DevExpress.XtraBars.Docking2010;
using System;
using System.Linq;
using DevExpress.XtraScheduler;
using DevExpress.LookAndFeel;

namespace sample_5_4 {
  public partial class EventView : XtraUserControl {
    public EventView() {
      InitializeComponent();

      eventGrid.BindingContext[eventGrid.DataSource].CurrentChanged += new EventHandler(eventGridView_CurrentChanged);
    }

    private void EventView_Load(object sender, System.EventArgs e) {
      TransactionalChangeHandler.Hookup<Event>(
        DocumentManager.FromControl(ParentForm).View, unitOfWork, commitChangesItem,
        discardChangesItem, eventCollection);
    }

    private void eventGrid_Load(object sender, System.EventArgs e) {
      ConfigureScheduler();
    }

    private Event GetFocusedEvent() {
      return (Event)eventGridView.GetFocusedRow();
    }

    private void ConfigureScheduler() {
      const int maxVisibleDays = 5;

      var ev = GetFocusedEvent();
      if (ev != null) {
        scheduler.Enabled = true;
        scheduler.LimitInterval.Start = ev.StartDate;
        scheduler.LimitInterval.End = ev.EndDate.AddDays(1);
        int days = ((int)(ev.EndDate - ev.StartDate).TotalDays) + 1;
        scheduler.DayView.DayCount = Math.Min(days, maxVisibleDays);
        if (days > maxVisibleDays) {
          dayScrollerPanel.Visible = true;
          dayScrollbar.Maximum = days - maxVisibleDays;
          dayScrollbar.Value = 0;
        }
        else
          dayScrollerPanel.Visible = false;

        scheduler.DayView.TopRowTime = new TimeSpan(
          ev.CapacityRequirements.Min(cr => cr.StartTime.TimeOfDay.Ticks)) -
          new TimeSpan(0, 30, 0);

        storage.Appointments.DataSource = ev.CapacityRequirements;
      }
      else {
        storage.Appointments.DataSource = null;
        scheduler.Enabled = false;
      }
    }

    void eventGridView_CurrentChanged(object sender, EventArgs e) {
      ConfigureScheduler();
    }

    private void dayScrollbar_Scroll(object sender, System.Windows.Forms.ScrollEventArgs e) {
      scheduler.Start = scheduler.LimitInterval.Start.AddDays(e.NewValue);
    }

    private void scheduler_VisibleIntervalChanged(object sender, EventArgs e) {
      dayScrollbar.Value = scheduler.Start > scheduler.LimitInterval.Start ?
      (int)(scheduler.Start - scheduler.LimitInterval.Start).TotalDays : 0;
    }

    private void scheduler_PopupMenuShowing(object sender, DevExpress.XtraScheduler.PopupMenuShowingEventArgs e) {
      // You can find out precisely which menu it is that's being shown now, like this:
      // if (e.Menu.Id == SchedulerMenuItemId.AppointmentMenu)

      e.Menu.RemoveMenuItem(SchedulerMenuItemId.EditSeries);
      e.Menu.RemoveMenuItem(SchedulerMenuItemId.StatusSubMenu);
      e.Menu.RemoveMenuItem(SchedulerMenuItemId.LabelSubMenu);
      e.Menu.RemoveMenuItem(SchedulerMenuItemId.NewAllDayEvent);
      e.Menu.RemoveMenuItem(SchedulerMenuItemId.NewRecurringAppointment);
      e.Menu.RemoveMenuItem(SchedulerMenuItemId.NewRecurringEvent);
      e.Menu.RemoveMenuItem(SchedulerMenuItemId.GotoDate);
      e.Menu.RemoveMenuItem(SchedulerMenuItemId.GotoThisDay);
      e.Menu.RemoveMenuItem(SchedulerMenuItemId.GotoToday);
      e.Menu.RemoveMenuItem(SchedulerMenuItemId.RestoreOccurrence);
      e.Menu.RemoveMenuItem(SchedulerMenuItemId.SwitchViewMenu);
      e.Menu.RemoveMenuItem(SchedulerMenuItemId.SplitAppointment);

      var item = e.Menu.GetMenuItemById(SchedulerMenuItemId.NewAppointment);
      if (item != null)
        item.Caption = "New Capacity &Requirement";
    }

    private void scheduler_EditAppointmentFormShowing(object sender, AppointmentFormEventArgs e) {
      DevExpress.XtraScheduler.SchedulerControl scheduler = ((DevExpress.XtraScheduler.SchedulerControl)(sender));
      sample_5_4.OutlookAppointmentForm form = new sample_5_4.OutlookAppointmentForm(scheduler, e.Appointment, e.OpenRecurrenceForm);
      try {
        e.DialogResult = form.ShowDialog();
        e.Handled = true;
      }
      finally {
        form.Dispose();
      }

    }

    private void scheduler_AppointmentViewInfoCustomizing(object sender, AppointmentViewInfoCustomizingEventArgs e) {
      e.ViewInfo.Appearance.BackColor = DXSkinColors.FillColors.Warning;

      if (e.ViewInfo.Appointment.GetRow(storage) is CapacityRequirement capReq &&
          capReq.RoomBooking != null)
        e.ViewInfo.Appearance.BackColor = DXSkinColors.FillColors.Success;
    }

    private void scheduler_InitAppointmentDisplayText(object sender, AppointmentDisplayTextEventArgs e) {
      if (e.ViewInfo.Appointment.GetRow(storage) is CapacityRequirement capReq &&
      capReq.RoomBooking != null)
        e.Text = $"{e.Appointment.Subject} ({capReq.RoomBooking.Room.Name})";
    }
  }
}