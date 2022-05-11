namespace sample_5_4.DashboardControls {
  partial class Venues {
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

    #region Component Designer generated code

    /// <summary> 
    /// Required method for Designer support - do not modify 
    /// the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent() {
      this.placeholder1 = new sample_5_4.DashboardControls.Placeholder();
      this.SuspendLayout();
      // 
      // placeholder1
      // 
      this.placeholder1.Dock = System.Windows.Forms.DockStyle.Fill;
      this.placeholder1.Location = new System.Drawing.Point(0, 0);
      this.placeholder1.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
      this.placeholder1.Name = "placeholder1";
      this.placeholder1.Size = new System.Drawing.Size(150, 150);
      this.placeholder1.TabIndex = 0;
      // 
      // Venues
      // 
      this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.Controls.Add(this.placeholder1);
      this.Name = "Venues";
      this.ResumeLayout(false);

    }

    #endregion

    private Placeholder placeholder1;
  }
}
