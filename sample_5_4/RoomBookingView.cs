using System.Linq;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using DevExpress.Xpo;
using EventsDB;
using System.Collections.Generic;
using EventsDB.Populate;
using DevExpress.XtraBars.Docking2010;
using DevExpress.XtraBars.Docking2010.Views;

namespace sample_5_4 {
  public partial class RoomBookingView : XtraUserControl {
    public RoomBookingView() {
      InitializeComponent();
    }

    private void RoomBookingView_Load(object sender, System.EventArgs e) {
      TransactionalChangeHandler.Hookup<RoomBooking>(
        DocumentManager.FromControl(ParentForm).View, unitOfWork, saveChangesItem,
        discardChangesItem, null, null, null, null, null, null,
        RoomBookingView_Closing,
        saveChangesItem_ItemClick,
        discardChangesItem_ItemClick);
    }

    #region Change Handling
    private void saveChangesItem_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e) {
      // use my own ValidateAndCommit
      ValidateAndCommit();
    }

    private bool ValidateAndCommit() {
      var checkables = TransactionalChangeHandler.GetCheckables(unitOfWork);
      if (TransactionalChangeHandler.Validate(checkables) == false)
        return false;

      var relatedRooms = checkables.
        OfType<RoomBooking>().
        Where(rb => rb.Room != null).
        Select(rb => rb.Room);

      var relatedRoomResults = DemoDatabase.BizRulezChecker.Check("save", relatedRooms).ToList();
      if (relatedRoomResults.Count > 0) {
        XtraMessageBox.Show("There have been errors during data validation.");
        return false;
      }

      unitOfWork.CommitChanges();
      return true;
    }

    private void RoomBookingView_Closing(object sender, DocumentCancelEventArgs e) {
      if (TransactionalChangeHandler.ObjectsToSave(unitOfWork)) {
        var answer = XtraMessageBox.Show("There are unsaved changes. Do you want to validate and save before closing the list?", "Question",
           MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
        if (answer == DialogResult.Yes) {
          // use my own ValidateAndCommit
          if (!ValidateAndCommit())
            e.Cancel = true;
        }
        else if (answer == DialogResult.Cancel)
          e.Cancel = true;
      }
    }

    private void discardChangesItem_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e) {
      if (TransactionalChangeHandler.ObjectsToSave(unitOfWork)) {
        if (XtraMessageBox.Show("This operation is going to discard all unsaved changes. Are you sure you want to proceed?", "Question",
          MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes) {
          //scheduler.BeginUpdate();
          try {
            // Use my own DiscardChanges 
            DiscardChanges(unitOfWork, null);
          }
          finally {
            //scheduler.EndUpdate();
          }
        }
      }
    }

    public void DiscardChanges(UnitOfWork unitOfWork, XPBaseCollection collection) {
      var deletedTLObjects = TransactionalChangeHandler.GetDeletedTLObjects<RoomBooking>(unitOfWork);

      foreach (var newObject in TransactionalChangeHandler.GetSavedNewObjects(unitOfWork)) {
        TransactionalChangeHandler.ResetAssociations(newObject);

        // special behavior for this form
        if (newObject is RoomBooking) {
          // remove the appointment from the collection, this doesn't
          // happen automatically
          var appt = schedulerStorage.Appointments.Items.Find(
            a => object.ReferenceEquals(a.GetRow(schedulerStorage), newObject));
          schedulerStorage.Appointments.Remove(appt);
        }
      }

      unitOfWork.ReloadChangedObjects();

      // special behavior: get the right collection first
      var roomBookingCollection = (IList<RoomBooking>)schedulerStorage.Appointments.DataSource;
      foreach (var item in deletedTLObjects)
        roomBookingCollection.Add(item);
    }


    #endregion
  }
}
