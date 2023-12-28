using System.Windows;

namespace mStructural.WPF
{
    /// <summary>
    /// Interaction logic for ProductionExportForm.xaml
    /// </summary>
    public partial class ProductionExportForm : Window
    {
        public ProductionExportForm()
        {
            InitializeComponent();
        }

        private void OkBtn_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }
    }
}
