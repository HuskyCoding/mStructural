using System;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;

namespace mStructural.WPF
{
    /// <summary>
    /// Interaction logic for Setting.xaml
    /// </summary>
    public partial class Setting : Window
    {
        public Setting()
        {
            InitializeComponent();

            // Extract default location from app setting
            appsetting appset = new appsetting();
            StartModelPath.Text = appset.StartModelPath;
            EdgeNoteTb.Text = appset.EdgeNote;
            DeleteTubeLaserBodyTb.Text = appset.DeleteTubeLaserBodyLoc;
            AutoViewGapTb.Text = appset.ViewGap.ToString();
            appset = null;
        }

        private void ConfirmBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // check if the directory exist
                if (!Directory.Exists(StartModelPath.Text))
                {
                    MessageBox.Show("Start Model Path is invalid.");
                    return;
                }

                // check if gap is numeric
                bool isDigit = false;
                double gap = 0;
                isDigit = double.TryParse(AutoViewGapTb.Text, out gap);

                if (!isDigit)
                {
                    MessageBox.Show("Auto view gap is not a number.");
                    return;
                }

                appsetting appset = new appsetting();
                appset.StartModelPath = StartModelPath.Text;
                appset.EdgeNote = EdgeNoteTb.Text;
                appset.DeleteTubeLaserBodyLoc = DeleteTubeLaserBodyTb.Text;
                appset.ViewGap = gap;
                appset.Save();
                DialogResult = true;
                appset = null;
                Close();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void StartModelPathBrowseBtn_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog fbd = new System.Windows.Forms.FolderBrowserDialog();
            System.Windows.Forms.DialogResult result = fbd.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                StartModelPath.Text = fbd.SelectedPath;
            }
        }

        private void dtlbTxtLocBrowseBtn_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.OpenFileDialog ofd = new System.Windows.Forms.OpenFileDialog();
            ofd.CheckFileExists = true;
            System.Windows.Forms.DialogResult result = ofd.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                DeleteTubeLaserBodyTb.Text = ofd.FileName;
            }
        }

        private void dtlbTxtLocOpenBtn_Click(object sender, RoutedEventArgs e)
        {
            Process.Start("explorer.exe", Path.GetDirectoryName(DeleteTubeLaserBodyTb.Text));
        }

        private void DefaultBtn_Click(object sender, RoutedEventArgs e)
        {
            string oneDrivePath = Environment.GetEnvironmentVariable("OneDrive");
            DeleteTubeLaserBodyTb.Text = oneDrivePath + @"\SolidworksLibraries\Macros\mStructural\Settings\dtlbstring.txt";
            AutoViewGapTb.Text = "10";
        }

        private void NumericInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            Regex regex = new Regex("[^0-9.]+");
            e.Handled = regex.IsMatch(e.Text);
        }
    }
}
