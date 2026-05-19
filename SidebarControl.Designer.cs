namespace sign_templater
{
    partial class SidebarControl
    {
        private System.ComponentModel.IContainer components = null;

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
            this.SuspendLayout();
            // 
            // SidebarControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "SidebarControl";
            this.Size = new System.Drawing.Size(400, 738);
            this.Load += new System.EventHandler(this.SidebarControl_Load);
            this.ResumeLayout(false);

        }
    }
}
