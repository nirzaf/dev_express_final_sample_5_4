using System;
using System.Linq;
using System.Windows.Forms;
using DevExpress.Xpo;
using DevExpress.XtraBars;
using DevExpress.XtraEditors;
using DevExpress.Xpo.Metadata;
using System.Collections.Generic;
using EventsDB;
using EventsDB.Populate;
using DevExpress.XtraBars.Docking2010.Views;

namespace sample_5_4 {
  public static class TransactionalChangeHandler {
    public static void Hookup<T>(BaseView view, UnitOfWork unitOfWork, BarButtonItem saveChangesItem, BarButtonItem discardChangesItem, XPBaseCollection topLevelCollection)
      where T : class {
      Hookup<T>(view, unitOfWork, saveChangesItem, discardChangesItem, topLevelCollection,
        null, null, null, null, null, null, null, null);
    }
    public static void Hookup<T>(BaseView view, UnitOfWork unitOfWork,
      BarButtonItem saveChangesItem, BarButtonItem discardChangesItem,
      XPBaseCollection topLevelCollection,
      SessionManipulationEventHandler afterCommitTransactionEventHandler,
      SessionManipulationEventHandler afterRollbackTransactionEventHandler,
      ObjectChangeEventHandler objectChangedEventHandler,
      ObjectManipulationEventHandler objectDeletedEventHandler,
      ObjectManipulationEventHandler objectSavedEventHandler,
      DocumentCancelEventHandler documentClosingHandler,
      ItemClickEventHandler saveChangesItemClickEventHandler,
      ItemClickEventHandler discardChangesItemClickEventHandler)
      where T : class {
      var container = new TransactionalChangeEventHandlerContainer<T>(
        view, unitOfWork, saveChangesItem,
        discardChangesItem, topLevelCollection
        );

      unitOfWork.AfterCommitTransaction +=
        afterCommitTransactionEventHandler ??
        container.SessionManipulationHandler;
      unitOfWork.AfterRollbackTransaction +=
        afterRollbackTransactionEventHandler ??
        container.SessionManipulationHandler;
      unitOfWork.ObjectChanged +=
        objectChangedEventHandler ?? container.ObjectChangeHandler;
      unitOfWork.ObjectDeleted +=
        objectDeletedEventHandler ??
        container.ObjectManipulationHandler;
      unitOfWork.ObjectSaved +=
        objectSavedEventHandler ?? container.ObjectManipulationHandler;

      view.DocumentClosing +=
        documentClosingHandler ?? container.DocumentClosingHandler;

      saveChangesItem.ItemClick +=
        saveChangesItemClickEventHandler ??
        container.SaveChangesItemClickHandler;
      discardChangesItem.ItemClick +=
        discardChangesItemClickEventHandler ??
        container.DiscardChangesItemClickHandler;
    }

    public static IEnumerable<XPBaseObject> GetCheckables(UnitOfWork unitOfWork) {
      var savedObjects = unitOfWork.GetObjectsToSave();
      var checkables = savedObjects.Cast<XPBaseObject>().Where(o => !o.IsDeleted);
      return checkables;
    }

    public static bool Validate(IEnumerable<XPBaseObject> checkables) {
      var results = DemoDatabase.BizRulezChecker.Check("save", checkables).ToList();
      if (results.Count > 0) {
        XtraMessageBox.Show("There have been errors during data validation.");
        return false;
      }
      return true;
    }

    public static bool ValidateAndCommit(UnitOfWork unitOfWork) {
      if (Validate(GetCheckables(unitOfWork))) {
        unitOfWork.CommitChanges();
        return true;
      }

      return false;
    }

    public static bool ObjectsToSave(UnitOfWork unitOfWork) {
      return unitOfWork.GetObjectsToSave().Count > 0 ||
        unitOfWork.GetObjectsToDelete().Count > 0;
    }

    public static List<T> GetDeletedTLObjects<T>(UnitOfWork unitOfWork)
                                where T : class {
      return unitOfWork.GetObjectsToDelete().Cast<XPBaseObject>().
                  OfType<T>().ToList();
    }

    public static IEnumerable<XPBaseObject> GetSavedNewObjects(UnitOfWork unitOfWork) {
      return unitOfWork.GetObjectsToSave().Cast<XPBaseObject>().
              Where(o => unitOfWork.IsNewObject(o));
    }

    public static void ResetAssociations(XPBaseObject newObject) {
      foreach (var member in newObject.ClassInfo.ObjectProperties.
                            Cast<XPMemberInfo>().
                            Where(mi => mi.IsAssociation)) {
        member.SetValue(newObject, null);
      }
    }

    public static void RemoveTLObjects<T>(XPBaseCollection collection, XPBaseObject ob)
      where T : class {
      var obAsT = ob as T;
      if (obAsT != null)
        collection.BaseRemove(ob);
    }

    public static void DiscardChanges<T>(UnitOfWork unitOfWork, XPBaseCollection collection) where T : class {
      // remember which top-level objects were going to be deleted
      // so we can restore them back into the UI collection
      var deletedTLObjects = GetDeletedTLObjects<T>(unitOfWork);

      // as long as the unit of work hasn't been committed, the newly added 
      // objects are recognized by the IsNewObject method
      // for the elements that had been added, find their association
      // properties and set their values to null, so that they are 
      // removed from the collections they had been added to
      foreach (var newObject in GetSavedNewObjects(unitOfWork)) {
        ResetAssociations(newObject);
        RemoveTLObjects<T>(collection, newObject);
      }

      unitOfWork.ReloadChangedObjects();

      // add the deleted objects back into the top level collection
      collection.BaseAddRange(deletedTLObjects);
    }


    private class TransactionalChangeEventHandlerContainer<T> where T : class {
      BaseView view;
      UnitOfWork unitOfWork;
      BarButtonItem saveChangesItem;
      BarButtonItem discardChangesItem;
      XPBaseCollection topLevelCollection;

      public TransactionalChangeEventHandlerContainer(BaseView view, UnitOfWork unitOfWork,
        BarButtonItem saveChangesItem, BarButtonItem discardChangesItem,
        XPBaseCollection topLevelCollection) {
        this.view = view;
        this.unitOfWork = unitOfWork;
        this.saveChangesItem = saveChangesItem;
        this.discardChangesItem = discardChangesItem;
        this.topLevelCollection = topLevelCollection;

        this.unitOfWork.TrackPropertiesModifications = true;

        EvalTransactionState();
      }

      public void ObjectManipulationHandler(object sender, ObjectManipulationEventArgs e) {
        EvalTransactionState();
      }

      public void SessionManipulationHandler(object sender, SessionManipulationEventArgs e) {
        EvalTransactionState();
      }

      void RemoveFromObjectToSaveIfNeeded(Session session, object o) {
        if (session.TrackPropertiesModifications && (session.IsObjectToSave(o) || session.IsObjectToDelete(o))) {
          var classInfo = session.GetClassInfo(o);

          if (classInfo.HasModifications(o)) {
            foreach (XPMemberInfo property in classInfo.PersistentProperties) {
              if (property.GetModified(o)) {
                if (!object.Equals(property.GetOldValue(o), property.GetValue(o))) {
                  return;
                }
              }
            }
          }
          session.RemoveFromLists(o);
          if (session.GetObjectsToSave().Count == 0 && session.GetObjectsToDelete().Count == 0) {
            session.DropChanges();
          }
        }
      }

      public void ObjectChangeHandler(object sender, ObjectChangeEventArgs e) {
        if (e.Reason == ObjectChangeReason.CancelEdit)
          RemoveFromObjectToSaveIfNeeded(unitOfWork, e.Object);
        EvalTransactionState();
      }

      void EvalTransactionState() {
        var objectsToSave = TransactionalChangeHandler.ObjectsToSave(unitOfWork);
        saveChangesItem.Enabled =
          discardChangesItem.Enabled = objectsToSave;
      }

      public void DocumentClosingHandler(object sender, DocumentCancelEventArgs e) {
        if (TransactionalChangeHandler.ObjectsToSave(unitOfWork)) {
          var answer = XtraMessageBox.Show("There are unsaved changes. Do you want to validate and save before closing the list?", "Question",
             MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
          if (answer == DialogResult.Yes) {
            if (!TransactionalChangeHandler.ValidateAndCommit(unitOfWork))
              e.Cancel = true;
          }
          else if (answer == DialogResult.Cancel)
            e.Cancel = true;
        }
      }

      public void SaveChangesItemClickHandler(object sender, DevExpress.XtraBars.ItemClickEventArgs e) {
        TransactionalChangeHandler.ValidateAndCommit(unitOfWork);
      }

      public void DiscardChangesItemClickHandler(object sender, DevExpress.XtraBars.ItemClickEventArgs e) {
        if (ObjectsToSave(unitOfWork)) {
          if (XtraMessageBox.Show("This operation is going to discard all unsaved changes. Are you sure you want to proceed?", "Question",
            MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes) {
            TransactionalChangeHandler.DiscardChanges<T>(unitOfWork, topLevelCollection);
          }
        }
      }
    }
  }
}