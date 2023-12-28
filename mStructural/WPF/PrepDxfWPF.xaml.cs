using mStructural.Function;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Collections;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;

namespace mStructural.WPF
{
    /// <summary>
    /// Interaction logic for EtchnNote.xaml
    /// </summary>
    public partial class PrepDxfWPF : Window
    {
        #region Private Variables
        private SldWorks swApp;
        private ModelDoc2 swModel;
        private DependancyUtil dependancyUtil;
        private Hashtable drawDep;
        private string pressStr = "";
        private string fullMatCode = "";
        private string basicMaterial = "";
        private string plateThickness = "";
        private bool bBlack = false;
        private bool isUpdateQty = false;
        private string topAssemblyPath = "";
        private bool isGenerated = false;
        private string selectedConfig = "";
        #endregion

        public PrepDxfWPF(SldWorks swapp, ModelDoc2 swmodel)
        {
            swApp = swapp;
            swModel = swmodel;
            dependancyUtil = new DependancyUtil();
            InitializeComponent();
        }

        private void OkBtn_Click(object sender, RoutedEventArgs e)
        {
            if(MatCodeCb.Text == "")
            {
                MessageBox.Show("Please select a material code");
                return;
            }

            if(ThicknessTb.Text == "")
            {
                MessageBox.Show("Plate thickness cannot be blank.");
                return;
            }

            if (pressCh.IsChecked == true)
            {
                pressStr = "P";
            }

            if(blackCh.IsChecked == true)
            {
                bBlack = true;
            }

            plateThickness = ThicknessTb.Text;

            if(MatGradeTb.Text == "")
            {
                fullMatCode = MatCodeCb.Text + " " + plateThickness + pressStr;
            }
            else
            {
                fullMatCode = MatCodeCb.Text + " " + MatGradeTb.Text + " " + plateThickness + pressStr;
            }

            switch (MatCodeCb.Text)
            {
                case "AL":
                case "ALC":
                    {
                        basicMaterial = "AL";
                    }
                    break;
                case "SM":
                case "SMC":
                    {
                        basicMaterial = "MS";
                    }
                    break;
                case "SS":
                    {
                        basicMaterial = "STS";
                    }
                    break;
                case "R":
                default:
                    {
                        basicMaterial = "OTHER";
                    }
                    break;
            }

            isUpdateQty = UpdateQtyCh.IsChecked ?? false;

            topAssemblyPath = SelectedPathTb.Text;

            PrepDxf prepDxf = new PrepDxf(swApp, swModel, fullMatCode, basicMaterial, plateThickness, bBlack, isUpdateQty, topAssemblyPath, selectedConfig);
            prepDxf.Run();

            Close();
        }

        private void ThicknessTb_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            Regex regex = new Regex("[^0-9.]+");
            e.Handled = regex.IsMatch(e.Text);
        }

        private void UpdateQtyCh_Checked(object sender, RoutedEventArgs e)
        {
            TopAssemblyCmb.IsEnabled = true;
            TopAssemblyConfigCmb.IsEnabled = true;

            if (!isGenerated)
            {
                // generate all dependency
                drawDep = dependancyUtil.getDependancy(swApp, swModel.GetPathName());
        
                foreach(DictionaryEntry entry in drawDep)
                {
                    TopAssemblyCmb.Items.Add(entry.Key.ToString());
                }

                // select top assembly
                string topAssembly = dependancyUtil.GetTopModelStrKey(swApp, swModel.GetPathName());
                TopAssemblyCmb.SelectedValue = topAssembly;

                var Entry =

                SelectedPathTb.Text = (string)drawDep[topAssembly];

                isGenerated = true;
            }
        }

        private void UpdateQtyCh_Unchecked(object sender, RoutedEventArgs e)
        {
            TopAssemblyCmb.IsEnabled = false;
            TopAssemblyConfigCmb.IsEnabled = false;
        }

        private void TopAssemblyCmb_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            string text = (sender as ComboBox).SelectedItem as string;
            string modelPath = (string)drawDep[text];
            SelectedPathTb.Text = modelPath;

            // clear config combobox items
            TopAssemblyConfigCmb.Items.Clear();

            // check model type
            int modelType = -1;
            switch (modelPath.Substring(modelPath.Length - 6).ToUpper())
            {
                case "SLDPRT":
                    {
                        modelType = (int)swDocumentTypes_e.swDocPART;
                    }
                    break;
                case "SLDASM":
                    {
                        modelType = (int)swDocumentTypes_e.swDocASSEMBLY;
                    }
                    break;
            }

            // open model
            int err = -1;
            int warn = -1;
            ModelDoc2 topAsmModel = swApp.OpenDoc6(modelPath, modelType, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref err, ref warn);

            // get configuration
            string[] configNames = topAsmModel.GetConfigurationNames();

            // add to dropdown
            foreach (string configName in configNames)
            {
                TopAssemblyConfigCmb.Items.Add(configName);
            }

            // select the current config
            TopAssemblyConfigCmb.SelectedValue = topAsmModel.GetActiveConfiguration().Name;
        }

        private void TopAssemblyConfigCmb_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            selectedConfig = (sender as ComboBox).SelectedItem as string;
        }
    }
}
