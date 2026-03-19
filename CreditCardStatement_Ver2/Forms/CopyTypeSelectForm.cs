namespace CreditCardStatement_Ver2.Forms
{
  public partial class CopyTypeSelectForm : Form
  {
    private System.ComponentModel.IContainer components = null;

    public CopyTypeSelectForm()
    {
      InitializeComponent();
    }

    protected override void Dispose(bool disposing)
    {
      if (disposing && (components != null))
      {
        components.Dispose();
      }
      base.Dispose(disposing);
    }


    private void InitializeComponent()
    {
      this.components = new System.ComponentModel.Container();
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.ClientSize = new System.Drawing.Size(800, 450);
      this.Text = "CopyTypeSelectForm";
    }
  }
}
