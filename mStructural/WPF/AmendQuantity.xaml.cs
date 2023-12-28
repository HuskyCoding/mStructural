using System.Windows;

namespace mStructural.WPF
{
    /// <summary>
    /// Interaction logic for AppendQuantity.xaml
    /// </summary>
    public partial class AppendQuantity : Window
    {
        public AppendQuantity()
        {
            InitializeComponent();
        }

        private void OkBtn_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }
    }
}
