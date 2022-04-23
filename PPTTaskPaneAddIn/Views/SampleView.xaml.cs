using System.Windows;
using System.Windows.Controls;

namespace PPTTaskPaneAddIn.Views
{
    /// <summary>
    /// Interaction logic for SampleView.xaml
    /// </summary>
    public partial class SampleView : UserControl
    {
        public SampleView()
        {
            InitializeComponent();
        }

        protected override void OnDpiChanged(DpiScale oldDpi, DpiScale newDpi)
        {
            base.OnDpiChanged(oldDpi, newDpi);
        }
    }
}
