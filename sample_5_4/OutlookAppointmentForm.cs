using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using DevExpress.Utils;
using DevExpress.Utils.Internal;
using DevExpress.Utils.Menu;
using DevExpress.XtraBars;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Controls;
using DevExpress.XtraEditors.Native;
using DevExpress.XtraScheduler;
using DevExpress.XtraScheduler.iCalendar;
using DevExpress.XtraScheduler.Localization;
using DevExpress.XtraScheduler.Native;
using DevExpress.XtraScheduler.Printing;
using DevExpress.XtraScheduler.Printing.Native;
using DevExpress.XtraScheduler.UI;
using DevExpress.XtraScheduler.Commands;
using DevExpress.XtraScheduler.Services;
using DevExpress.XtraPrinting;
using DevExpress.Utils.CommonDialogs;
using CommonDialogsInternal = DevExpress.Utils.CommonDialogs.Internal;
using DevExpress.Utils.Svg;

namespace sample_5_4 {
  /// <summary>
  /// Summary description for AppointmentRibbonForm.
  /// </summary>
  public partial class OutlookAppointmentForm : DevExpress.XtraBars.Ribbon.RibbonForm, IDXManagerPopupMenu {
    #region Fields
    bool openRecurrenceForm;
    readonly ISchedulerStorage storage;
    readonly SchedulerControl control;
    SvgImage recurringSvgImage;
    SvgImage normalSvgImage;
    readonly AppointmentFormController controller;
    IDXMenuManager menuManager;
    bool supressCancelCore;
    #endregion

    [EditorBrowsable(EditorBrowsableState.Never)]
    public OutlookAppointmentForm() {
      InitializeComponent();
    }
    public OutlookAppointmentForm(DevExpress.XtraScheduler.SchedulerControl control, Appointment apt)
        : this(control, apt, false) {
    }
    public OutlookAppointmentForm(DevExpress.XtraScheduler.SchedulerControl control, Appointment apt, bool openRecurrenceForm) {
      Guard.ArgumentNotNull(control, "control");
      Guard.ArgumentNotNull(control.DataStorage, "control.DataStorage");
      Guard.ArgumentNotNull(apt, "apt");

      this.openRecurrenceForm = openRecurrenceForm;
      this.controller = CreateController(control, apt);
      //
      // Required for Windows Form Designer support
      //
      InitializeComponent();
      SetupPredefinedConstraints();

      LoadIcons();

      this.control = control;
      this.storage = control.DataStorage;

      this.edtResource.SchedulerControl = control;
      this.edtResource.Storage = this.storage;
      this.edtResources.SchedulerControl = control;

      this.riAppointmentResource.SchedulerControl = control;
      this.riAppointmentResource.Storage = this.storage;
      this.riAppointmentStatus.Storage = this.storage;

      this.riAppointmentLabel.Storage = this.storage;

      BindControllerToControls();

      LookAndFeel.ParentLookAndFeel = control.LookAndFeel;

      this.supressCancelCore = false;
    }
    #region Properties
    [Browsable(false)]
    public IDXMenuManager MenuManager { get { return this.menuManager; } private set { this.menuManager = value; } }
    protected internal AppointmentFormController Controller { get { return this.controller; } }
    protected internal SchedulerControl Control { get { return this.control; } }
    protected internal ISchedulerStorage Storage { get { return this.storage; } }
    protected internal bool IsNewAppointment { get { return this.controller != null ? this.controller.IsNewAppointment : true; } }
    protected internal SvgImage RecurringSvgImage { get { return this.recurringSvgImage; } }
    protected internal SvgImage NormalSvgImage { get { return this.normalSvgImage; } }
    [Browsable(false), EditorBrowsable(EditorBrowsableState.Never), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public SvgImage SvgImage { get { return IconOptions.SvgImage; } set { IconOptions.SvgImage = value; } }
    protected internal bool OpenRecurrenceForm { get { return this.openRecurrenceForm; } }
    [DXDescription("DevExpress.XtraScheduler.UI.AppointmentRibbonForm,ReadOnly")]
    [DXCategory(CategoryName.Behavior)]
    [DefaultValue(false)]
    public bool ReadOnly {
      get { return Controller.ReadOnly; }
      set {
        if (Controller.ReadOnly == value)
          return;
        Controller.ReadOnly = value;
      }
    }
    protected override FormShowMode ShowMode { get { return DevExpress.XtraEditors.FormShowMode.AfterInitialization; } }
    internal Point FormLocation { get; private set; }
    #endregion

    public virtual void LoadFormData(Appointment appointment) {
      //do nothing
    }
    public virtual bool SaveFormData(Appointment appointment) {
      return true;
    }
    public virtual bool IsAppointmentChanged(Appointment appointment) {
      return false;
    }
    public virtual void SetMenuManager(DevExpress.Utils.Menu.IDXMenuManager menuManager) {
      MenuManagerUtils.SetMenuManager(Controls, menuManager);
      this.menuManager = menuManager;
    }

    protected internal virtual void SetupPredefinedConstraints() {
      this.tbProgress.Properties.Minimum = AppointmentProcessValues.Min;
      this.tbProgress.Properties.Maximum = AppointmentProcessValues.Max;
      this.tbProgress.Properties.SmallChange = AppointmentProcessValues.Step;
      this.edtResources.Visible = true;
      if (!WindowsFormsSettings.GetAllowDefaultSvgImages()) {
        foreach (BarItem item in this.ribbonControl1.Items)
          item.ImageOptions.SvgImage = null;
      }
    }
    protected virtual void BindControllerToControls() {
      DataBindings.Add("Text", Controller, "Caption");
      BindControllerToIcon();
      BindProperties(this.tbSubject, "Text", "Subject");
      BindProperties(this.tbLocation, "Text", "Location");
      BindProperties(this.tbDescription, "Text", "Description");
      BindProperties(this.edtStartDate, "EditValue", "DisplayStartDate");
      BindProperties(this.edtStartDate, "Enabled", "IsDateTimeEditable");
      BindProperties(this.edtStartTime, "EditValue", "DisplayStartTime");
      BindProperties(this.edtStartTime, "Enabled", "IsTimeEnabled");
      BindProperties(this.edtEndDate, "EditValue", "DisplayEndDate", DataSourceUpdateMode.Never);
      BindProperties(this.edtEndDate, "Enabled", "IsDateTimeEditable", DataSourceUpdateMode.Never);
      BindProperties(this.edtEndTime, "EditValue", "DisplayEndTime", DataSourceUpdateMode.Never);
      BindProperties(this.edtEndTime, "Enabled", "IsTimeEnabled", DataSourceUpdateMode.Never);
      BindProperties(this.chkAllDay, "Checked", "AllDay");
      BindProperties(this.chkAllDay, "Enabled", "IsDateTimeEditable");

      BindProperties(this.lblResource, "Enabled", "CanEditResource");

      BindProperties(this.edtResources, "ResourceIds", "ResourceIds");
      BindProperties(this.edtResources, "Visible", "ResourceSharing");
      BindProperties(this.edtResources, "Enabled", "CanEditResource");

      BindProperties(this.edtResource, "ResourceId", "ResourceId");
      BindProperties(this.edtResource, "Enabled", "CanEditResource");
      BindToBoolPropertyAndInvert(this.edtResource, "Visible", "ResourceSharing");

      BindProperties(this.barLabel, "EditValue", "Label");

      BindProperties(this.barStatus, "EditValue", "Status");

      BindBoolToVisibility(this.barReminder, "Visibility", "ReminderVisible");
      BindProperties(this.barReminder, "Editvalue", "ReminderTimeBeforeStart");

      BindProperties(this.tbProgress, "Value", "PercentComplete");
      BindProperties(this.lblPercentCompleteValue, "Text", "PercentComplete", ObjectToStringConverter);
      BindProperties(this.progressPanel, "Visible", "ShouldEditTaskProgress");
      BindProperties(this.btnDelete, "Enabled", "CanDeleteAppointment");

      BindBoolToVisibility(this.btnRecurrence, "Visibility", "ShouldShowRecurrenceButton");
      BindProperties(this.btnRecurrence, "Down", "IsRecurrentAppointment");


      BindToBoolPropertyAndInvert(this.ribbonControl1, "Enabled", "ReadOnly");

      BindProperties(this.edtTimeZone, "Visible", "TimeZoneVisible");
      BindProperties(this.edtTimeZone, "EditValue", "TimeZoneId");
      BindProperties(this.edtTimeZone, "Enabled", "TimeZoneEnabled");

      BindBoolToVisibility(this.btnTimeZones, "Visibility", "TimeZonesEnabled");
      BindProperties(this.btnTimeZones, "Down", "TimeZoneVisible");
    }

    protected virtual void BindControllerToIcon() {
      Binding binding = new Binding("SvgImage", Controller, "AppointmentType");
      binding.Format += AppointmentTypeToIconConverter;
      DataBindings.Add(binding);
    }
    protected virtual void ObjectToStringConverter(object o, ConvertEventArgs e) {
      e.Value = e.Value.ToString();
    }
    protected virtual void AppointmentTypeToIconConverter(object o, ConvertEventArgs e) {
      AppointmentType type = (AppointmentType)e.Value;
      if (type == AppointmentType.Pattern)
        e.Value = RecurringSvgImage;
      else
        e.Value = NormalSvgImage;
    }
    protected virtual void BindProperties(Control target, string targetProperty, string sourceProperty) {
      BindProperties(target, targetProperty, sourceProperty, DataSourceUpdateMode.OnPropertyChanged);
    }
    protected virtual void BindProperties(Control target, string targetProperty, string sourceProperty, DataSourceUpdateMode updateMode) {
      target.DataBindings.Add(targetProperty, Controller, sourceProperty, true, updateMode);
      BindToIsReadOnly(target, updateMode);
    }
    protected virtual void BindProperties(Control target, string targetProperty, string sourceProperty, ConvertEventHandler objectToStringConverter) {
      Binding binding = new Binding(targetProperty, Controller, sourceProperty, true);
      binding.Format += objectToStringConverter;
      target.DataBindings.Add(binding);
    }
    protected virtual void BindToBoolPropertyAndInvert(Control target, string targetProperty, string sourceProperty) {
      target.DataBindings.Add(new BoolInvertBinding(targetProperty, Controller, sourceProperty));
      BindToIsReadOnly(target);
    }
    protected virtual void BindToIsReadOnly(Control control) {
      BindToIsReadOnly(control, DataSourceUpdateMode.OnPropertyChanged);
    }
    protected virtual void BindToIsReadOnly(Control control, DataSourceUpdateMode updateMode) {
      if ((!(control is BaseEdit)) || control.DataBindings["ReadOnly"] != null)
        return;
      control.DataBindings.Add("ReadOnly", Controller, "ReadOnly", true, updateMode);
    }

    protected virtual void BindProperties(DevExpress.XtraBars.BarItem target, string targetProperty, string sourceProperty) {
      BindProperties(target, targetProperty, sourceProperty, DataSourceUpdateMode.OnPropertyChanged);
    }
    protected virtual void BindProperties(DevExpress.XtraBars.BarItem target, string targetProperty, string sourceProperty, DataSourceUpdateMode updateMode) {
      target.DataBindings.Add(targetProperty, Controller, sourceProperty, true, updateMode);
    }
    protected virtual void BindProperties(DevExpress.XtraBars.BarItem target, string targetProperty, string sourceProperty, ConvertEventHandler objectToStringConverter) {
      Binding binding = new Binding(targetProperty, Controller, sourceProperty, true);
      binding.Format += objectToStringConverter;
      target.DataBindings.Add(binding);
    }
    protected virtual void BindToBoolPropertyAndInvert(DevExpress.XtraBars.BarItem target, string targetProperty, string sourceProperty) {
      target.DataBindings.Add(new BoolInvertBinding(targetProperty, Controller, sourceProperty));
    }
    protected virtual void BindBoolToVisibility(DevExpress.XtraBars.BarItem target, string targetProperty, string sourceProperty) {
      target.DataBindings.Add(new BoolToVisibilityBinding(targetProperty, Controller, sourceProperty, false));
    }
    protected virtual void BindBoolToVisibility(DevExpress.XtraBars.BarItem target, string targetProperty, string sourceProperty, bool invert) {
      target.DataBindings.Add(new BoolToVisibilityBinding(targetProperty, Controller, sourceProperty, invert));
    }

    protected override void OnLoad(EventArgs e) {
      base.OnLoad(e);
      if (Controller == null)
        return;
      SubscribeControlsEvents();
      LoadFormData(Controller.EditedAppointmentCopy);
    }
    protected virtual AppointmentFormController CreateController(SchedulerControl control, Appointment apt) {
      return new AppointmentFormController(control, apt);
    }
    protected internal virtual void LoadIcons() {
      Assembly asm = typeof(Appointment).Assembly;
      this.recurringSvgImage = ResourceImageHelper.CreateSvgImageFromResources(SchedulerSvgImageNames.NewRecurringAppointment, asm);
      this.normalSvgImage = ResourceImageHelper.CreateSvgImageFromResources(SchedulerSvgImageNames.NewAppointment, asm);
    }
    protected internal virtual void SubscribeControlsEvents() {
      this.edtEndDate.Validating += new CancelEventHandler(OnEdtEndDateValidating);
      this.edtEndDate.InvalidValue += new InvalidValueExceptionEventHandler(OnEdtEndDateInvalidValue);
      this.edtEndTime.Validating += new CancelEventHandler(OnEdtEndTimeValidating);
      this.edtEndTime.InvalidValue += new InvalidValueExceptionEventHandler(OnEdtEndTimeInvalidValue);
      this.riDuration.Validating += new CancelEventHandler(OnCbReminderValidating);
      this.edtStartDate.Validating += new CancelEventHandler(OnEdtStartDateValidating);
      this.edtStartDate.InvalidValue += new InvalidValueExceptionEventHandler(OnEdtStartDateInvalidValue);
      this.edtStartTime.Validating += new CancelEventHandler(OnEdtStartTimeValidating);
      this.edtStartTime.InvalidValue += new InvalidValueExceptionEventHandler(OnEdtStartTimeInvalidValue);
    }

    protected internal virtual void UnsubscribeControlsEvents() {
      this.edtEndDate.Validating -= new CancelEventHandler(OnEdtEndDateValidating);
      this.edtEndDate.InvalidValue -= new InvalidValueExceptionEventHandler(OnEdtEndDateInvalidValue);
      this.edtEndTime.Validating -= new CancelEventHandler(OnEdtEndTimeValidating);
      this.edtEndTime.InvalidValue -= new InvalidValueExceptionEventHandler(OnEdtEndTimeInvalidValue);
      this.riDuration.Validating -= new CancelEventHandler(OnCbReminderValidating);
      this.edtStartDate.Validating -= new CancelEventHandler(OnEdtStartDateValidating);
      this.edtStartDate.InvalidValue -= new InvalidValueExceptionEventHandler(OnEdtStartDateInvalidValue);
      this.edtStartTime.Validating -= new CancelEventHandler(OnEdtStartTimeValidating);
      this.edtStartTime.InvalidValue -= new InvalidValueExceptionEventHandler(OnEdtStartTimeInvalidValue);
    }

    protected internal virtual void OnEdtStartTimeInvalidValue(object sender, InvalidValueExceptionEventArgs e) {
      e.ErrorText = SchedulerLocalizer.GetString(SchedulerStringId.Msg_DateOutsideLimitInterval);
    }
    protected internal virtual void OnEdtStartTimeValidating(object sender, CancelEventArgs e) {
      e.Cancel = !Controller.ValidateLimitInterval(this.edtStartDate.DateTime.Date, this.edtStartTime.Time.TimeOfDay, this.edtEndDate.DateTime.Date, this.edtEndTime.Time.TimeOfDay);
    }
    protected internal virtual void OnEdtStartDateInvalidValue(object sender, InvalidValueExceptionEventArgs e) {
      e.ErrorText = SchedulerLocalizer.GetString(SchedulerStringId.Msg_DateOutsideLimitInterval);
    }
    protected internal virtual void OnEdtStartDateValidating(object sender, CancelEventArgs e) {
      e.Cancel = !Controller.ValidateLimitInterval(this.edtStartDate.DateTime.Date, this.edtStartTime.Time.TimeOfDay, this.edtEndDate.DateTime.Date, this.edtEndTime.Time.TimeOfDay);
    }
    protected internal virtual void OnEdtEndDateValidating(object sender, CancelEventArgs e) {
      e.Cancel = !IsValidInterval();
      if (!e.Cancel)
        this.edtEndDate.DataBindings["EditValue"].WriteValue();
    }
    protected internal virtual void OnEdtEndDateInvalidValue(object sender, InvalidValueExceptionEventArgs e) {
      if (!AppointmentFormControllerBase.ValidateInterval(this.edtStartDate.DateTime.Date, this.edtStartTime.Time.TimeOfDay, this.edtEndDate.DateTime.Date, this.edtEndTime.Time.TimeOfDay))
        e.ErrorText = SchedulerLocalizer.GetString(SchedulerStringId.Msg_InvalidEndDate);
      else
        e.ErrorText = SchedulerLocalizer.GetString(SchedulerStringId.Msg_DateOutsideLimitInterval);
    }
    protected internal virtual void OnEdtEndTimeValidating(object sender, CancelEventArgs e) {
      e.Cancel = !IsValidInterval();
      if (!e.Cancel)
        this.edtEndTime.DataBindings["EditValue"].WriteValue();
    }
    protected internal virtual void OnEdtEndTimeInvalidValue(object sender, InvalidValueExceptionEventArgs e) {
      if (!AppointmentFormControllerBase.ValidateInterval(this.edtStartDate.DateTime.Date, this.edtStartTime.Time.TimeOfDay, this.edtEndDate.DateTime.Date, this.edtEndTime.Time.TimeOfDay))
        e.ErrorText = SchedulerLocalizer.GetString(SchedulerStringId.Msg_InvalidEndDate);
      else
        e.ErrorText = SchedulerLocalizer.GetString(SchedulerStringId.Msg_DateOutsideLimitInterval);
    }
    protected internal virtual bool IsValidInterval() {
      return AppointmentFormControllerBase.ValidateInterval(this.edtStartDate.DateTime.Date, this.edtStartTime.Time.TimeOfDay, this.edtEndDate.DateTime.Date, this.edtEndTime.Time.TimeOfDay) &&
          Controller.ValidateLimitInterval(this.edtStartDate.DateTime.Date, this.edtStartTime.Time.TimeOfDay, this.edtEndDate.DateTime.Date, this.edtEndTime.Time.TimeOfDay);
    }
    protected internal virtual void OnOkButton() {
      Save(true);
    }
    protected virtual void OnSaveButton() {
      Save(false);
    }

    void Save(bool closeAfterSave) {
      if (!ValidateDateAndTime())
        return;
      if (!SaveFormData(Controller.EditedAppointmentCopy))
        return;
      if (!Controller.IsConflictResolved()) {
        ShowMessageBox(SchedulerLocalizer.GetString(SchedulerStringId.Msg_Conflict), Controller.GetMessageBoxCaption(SchedulerStringId.Msg_Conflict), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        return;
      }
      if (!Controller.IsTimeValid()) {
        ShowMessageBox(SchedulerLocalizer.GetString(SchedulerStringId.Msg_InvalidAppointmentTime), Controller.GetMessageBoxCaption(SchedulerStringId.Msg_InvalidAppointmentTime), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        return;
      }

      if (IsAppointmentChanged(Controller.EditedAppointmentCopy) || Controller.IsAppointmentChanged() || Controller.IsNewAppointment)
        Controller.ApplyChanges();
      if (closeAfterSave) {
        this.supressCancelCore = true;
        DialogResult = System.Windows.Forms.DialogResult.OK;
      }
    }
    private bool ValidateDateAndTime() {
      this.edtEndDate.DoValidate();
      this.edtEndTime.DoValidate();
      this.edtStartDate.DoValidate();
      this.edtStartTime.DoValidate();

      return String.IsNullOrEmpty(this.edtEndTime.ErrorText) && String.IsNullOrEmpty(this.edtEndDate.ErrorText) && String.IsNullOrEmpty(this.edtStartDate.ErrorText) && String.IsNullOrEmpty(this.edtStartTime.ErrorText);
    }
    protected virtual void OnSaveAsButton() {
      ISaveFileDialog fileDialog = CommonDialogsInternal.CommonDialogProvider.Instance.CreateDefaultSaveFileDialog() ?? CommonDialogsInternal.CommonDialogProvider.Instance.CreateStandardSaveFileDialog();
      fileDialog.Filter = "iCalendar files (*.ics)|*.ics";
      fileDialog.FilterIndex = 1;
      if (fileDialog.ShowDialog() != CommonDialogsInternal.DialogResult.OK)
        return;
      try {
        using (Stream stream = new FileInfo(fileDialog.FileName).OpenWrite())
          ExportAppointment(stream);
      }
      catch {
        ShowMessageBox("Error: could not export appointments", String.Empty, MessageBoxButtons.OK, MessageBoxIcon.Error);
      }
    }
    void ExportAppointment(Stream stream) {
      if (stream == null)
        return;

      AppointmentBaseCollection aptsToExport = new AppointmentBaseCollection();
      aptsToExport.Add(Controller.EditedAppointmentCopy);
      iCalendarExporter exporter = new iCalendarExporter(this.storage, aptsToExport);

      exporter.ProductIdentifier = "-//Developer Express Inc.";
      exporter.Export(stream);
    }
    protected internal virtual DialogResult ShowMessageBox(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon) {
      return XtraMessageBox.Show(this, text, caption, buttons, icon);
    }
    protected internal virtual void OnDeleteButton() {
      if (IsNewAppointment)
        return;

      Controller.DeleteAppointment();

      DialogResult = DialogResult.Abort;
      Close();
    }
    protected internal virtual void OnRecurrenceButton() {
      if (!Controller.ShouldShowRecurrenceButton)
        return;

      Appointment patternCopy = Controller.PrepareToRecurrenceEdit();

      DialogResult result;
      using (Form form = CreateAppointmentRecurrenceForm(patternCopy, Control.OptionsView.FirstDayOfWeek)) {
        result = ShowRecurrenceForm(form);
      }

      if (result == DialogResult.Abort) {
        Controller.RemoveRecurrence();
      }
      else if (result == DialogResult.OK) {
        Controller.ApplyRecurrence(patternCopy);
      }

      this.btnRecurrence.Down = Controller.IsRecurrentAppointment;
    }
    protected virtual void OnCloseButton() {
      Close();
    }

    private bool CancelCore() {
      bool result = true;

      if (DialogResult != System.Windows.Forms.DialogResult.Abort && Controller != null && Controller.IsAppointmentChanged() && !this.supressCancelCore) {
        DialogResult dialogResult = ShowMessageBox(SchedulerLocalizer.GetString(SchedulerStringId.Msg_SaveBeforeClose), Controller.GetMessageBoxCaption(SchedulerStringId.Msg_SaveBeforeClose), MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);

        if (dialogResult == System.Windows.Forms.DialogResult.Cancel)
          result = false;
        else if (dialogResult == System.Windows.Forms.DialogResult.Yes)
          Save(true);
      }

      return result;
    }

    protected virtual DialogResult ShowRecurrenceForm(Form form) {
      return FormTouchUIAdapter.ShowDialog(form, this);
    }
    protected internal virtual Form CreateAppointmentRecurrenceForm(Appointment patternCopy, FirstDayOfWeek firstDayOfWeek) {
      AppointmentRecurrenceForm form = new AppointmentRecurrenceForm(patternCopy, firstDayOfWeek, Controller);
      form.SetMenuManager(MenuManager);
      form.LookAndFeel.ParentLookAndFeel = LookAndFeel;
      form.RightToLeftLayout = RightToLeftLayout;
      form.RightToLeft = RightToLeft;
      form.ShowExceptionsRemoveMsgBox = this.controller.AreExceptionsPresent();
      return form;
    }
    protected override void OnShown(EventArgs e) {
      base.OnShown(e);
      if (this.openRecurrenceForm) {
        this.openRecurrenceForm = false;
        OnRecurrenceButton();
      }
    }
    protected internal virtual void OnCbReminderValidating(object sender, CancelEventArgs e) {
      TimeSpan span = (TimeSpan)this.barReminder.EditValue;
      e.Cancel = span.Ticks < 0 && span != TimeSpan.MinValue;
      if (!e.Cancel)
        this.barReminder.DataBindings["EditValue"].WriteValue();
    }

    protected internal virtual void OnNextButton() {
      if (CancelCore()) {
        this.supressCancelCore = true;
        OpenNextAppointmentCommand command = new OpenNextAppointmentCommand(Control);
        command.Execute();
        Close();
      }
    }

    protected internal virtual void OnPreviousButton() {
      if (CancelCore()) {
        this.supressCancelCore = true;
        OpenPrevAppointmentCommand command = new OpenPrevAppointmentCommand(Control);
        command.Execute();
        Close();
      }
    }

    protected internal virtual void OnTimeZonesButton() {
      Controller.TimeZoneVisible = !Controller.TimeZoneVisible;
    }

    protected virtual void OnApplicationButtonClick() {
      FormLocation = Location;
      this.appointmentBackstageControl.SetPrintDocument(Control.GetPrintPreviewDocument(new RibbonFormPreviewMemoPrintStyle(Controller.EditedAppointmentCopy)));
    }

    protected virtual void OnPrintButton() {
      using (PrintTool tool = new PrintTool(this.appointmentBackstageControl.Document))
        tool.Print(this.appointmentBackstageControl.SelectedPrinterName);
    }

    protected virtual void OnPrintOptionsButton() {
      using (PrintTool tool = new PrintTool(this.appointmentBackstageControl.Document))
        tool.PrintDialog(FindForm(), LookAndFeel);
    }

    private void btnSaveAndClose_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e) {
      OnOkButton();
    }

    private void barButtonDelete_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e) {
      OnDeleteButton();
    }

    private void barRecurrence_ItemClick(object sender, ItemClickEventArgs e) {
      OnRecurrenceButton();
    }

    private void bvbSave_ItemClick(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e) {
      OnSaveButton();
    }

    private void bvbSaveAs_ItemClick(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e) {
      OnSaveAsButton();
    }

    private void bvbClose_ItemClick(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e) {
      OnCloseButton();
    }

    private void btnSave_ItemClick(object sender, ItemClickEventArgs e) {
      OnSaveButton();
    }

    protected override void OnClosing(CancelEventArgs e) {
      e.Cancel = !CancelCore();
      base.OnClosing(e);
    }

    private void btnNext_ItemClick(object sender, ItemClickEventArgs e) {
      OnNextButton();
    }

    private void btnPrevious_ItemClick(object sender, ItemClickEventArgs e) {
      OnPreviousButton();
    }

    private void btnTimeZones_ItemClick(object sender, ItemClickEventArgs e) {
      OnTimeZonesButton();
    }

    private void ribbonControl1_ApplicationButtonClick(object sender, EventArgs e) {
      OnApplicationButtonClick();
    }

    private void schedulerPrint_PrintOptionsClick(object sender, System.EventArgs e) {
      OnPrintOptionsButton();
    }

    private void schedulerPrint_PrintClick(object sender, System.EventArgs e) {
      OnPrintButton();
    }

    protected override void Dispose(bool disposing) {
      if (disposing) {
        if (this.components != null) {
          this.components.Dispose();
        }
        if (LookAndFeel != null)
          LookAndFeel.ParentLookAndFeel = null;
      }
      base.Dispose(disposing);
    }
  }
}