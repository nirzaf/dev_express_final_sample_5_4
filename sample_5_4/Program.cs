using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using DevExpress.Skins;
using DevExpress.Xpo;
using DevExpress.Xpo.DB;
using pct.BizRulez;
using EventsDB.Validation;
using pct.BizRulez.Rules;
using EventsDB;
using EventsDB.Populate;

namespace sample_5_4 {
  static class Program {
    /// <summary>
    /// The main entry point for the application.
    /// </summary>
    [STAThread]
    static void Main() {
      Application.EnableVisualStyles();
      Application.SetCompatibleTextRenderingDefault(false);

      SkinManager.EnableMdiFormSkins();

      DemoDatabase.Populate();

      mainForm = new MainForm();
      Application.Run(MainForm);
    }

    static MainForm mainForm;
    public static MainForm MainForm {
      get {
        return mainForm;
      }
    }
  }
}
