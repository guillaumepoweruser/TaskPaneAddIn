using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace PPTTaskPaneAddIn.Views
{
    [ComVisible(true)]
    public partial class TaskPaneUserControlPpt : UserControl
    {
        public SampleView TaskPaneSample { get; set; }
        public ElementHost WpfHost { get; set; }

        public TaskPaneUserControlPpt()
        {
            InitializeComponent();

            SuspendLayout();
            WpfHost = new ElementHost
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                Margin = new Padding(0)
            };
            panelContainer.Controls.Add(WpfHost, 0, 0);
            Controls.Add(panelContainer);
            ResumeLayout(false);
        }

        protected override void OnLoad(EventArgs e)
        {
            try
            {
                Initialize();
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message);
            }
        }

        private void Initialize()
        {
            TaskPaneSample = new SampleView();
            WpfHost.Child = TaskPaneSample;    
            WpfHost.Child.AddHandler(UIElement.PreviewMouseDownEvent, new RoutedEventHandler(PreviewMouseDown_Event), true);
        }

        private void PreviewMouseDown_Event(object sender, RoutedEventArgs e)
        {
            if (!WpfHost.Focused)
            {
                WpfHost.Focus();
            }
        }
    }
}
