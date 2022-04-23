using System.Windows.Forms;

namespace PPTTaskPaneAddIn.Views
{
    partial class TaskPaneUserControlPpt
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

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            this.panelContainer = new System.Windows.Forms.TableLayoutPanel();
            this.SuspendLayout();
            // 
            // panelContainer
            // 
            this.panelContainer.AutoSize = true;
            this.panelContainer.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.panelContainer.ColumnCount = 1;
            this.panelContainer.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.panelContainer.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelContainer.Location = new System.Drawing.Point(0, 0);
            this.panelContainer.Margin = new System.Windows.Forms.Padding(0);
            this.panelContainer.Name = "panelContainer";
            this.panelContainer.RowCount = 1;
            this.panelContainer.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.panelContainer.Size = new System.Drawing.Size(0, 0);
            this.panelContainer.TabIndex = 0;
            // 
            // MainTaskPaneUserControlPpt
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.Controls.Add(this.panelContainer);
            this.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.Margin = new System.Windows.Forms.Padding(0);
            this.Name = "MainTaskPaneUserControlPpt";
            this.Size = new System.Drawing.Size(0, 0);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public TabPage InformationTabPage;
        private TableLayoutPanel panelContainer;
    }
}