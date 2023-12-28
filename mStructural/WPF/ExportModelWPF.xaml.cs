using Microsoft.VisualBasic.Logging;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.IO;
using System.Text;
using System.Windows;

namespace mStructural.WPF
{
    /// <summary>
    /// Interaction logic for ExportModelWPF.xaml
    /// </summary>
    public partial class ExportModelWPF : Window
    {
        #region Private Variables
        private SldWorks swApp;
        private ModelDoc2 swModel;

        string[] configNames = null;
        #endregion

        // Constructor
        public ExportModelWPF(SldWorks swapp, ModelDoc2 swmodel)
        {
            InitializeComponent();

            swApp = swapp;
            swModel = swmodel;

            // get all configuration in opened model
            configNames = (string[])swModel.GetConfigurationNames();

            // add to list box
            foreach(string configName in configNames)
            {
                EDrawingLb.Items.Add(configName);
                TubeLaserLb.Items.Add(configName);
            }
        }

        private void ConfirmBtn_Click(object sender, RoutedEventArgs e)
        {
            int err = -1;
            int warn = -1;

            // clear selection first
            swModel.ClearSelection2(true);

            // get revision
            CustomPropertyManager swCusPropMgr = swModel.Extension.CustomPropertyManager[""];
            string val;
            string resVal;
            swCusPropMgr.Get4("Revision", false, out val, out resVal);

            // full path
            string modelPath = swModel.GetPathName();
            
            // model name
            string modelName = Path.GetFileNameWithoutExtension(modelPath);
            
            // directory
            string folderPath = Path.GetDirectoryName(modelPath);

            // check if the model is part or assembly
            string ext = "";
            switch (swModel.GetType())
            {
                case (int)swDocumentTypes_e.swDocASSEMBLY:
                    {
                        ext = ".easm";
                        break;
                    }
                case (int)swDocumentTypes_e.swDocPART:
                    {
                        ext = ".eprt";
                        break;
                    }
            }

            // save edrawing
            if (EdrawingCh.IsChecked == true)
            {
                // instantiate new string builder for config string
                StringBuilder sb = new StringBuilder();
                if(EDrawingLb.SelectedItems.Count > 0)
                {
                    foreach(var item in EDrawingLb.SelectedItems)
                    {
                        sb.Append(item.ToString());
                        sb.Append("\n");
                    }
                    string configString = sb.ToString();

                    // set config string to export option
                    swApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swEdrawingsSaveAsSelectionOption, (int)swEdrawingSaveAsOption_e.swEdrawingSaveSelected);
                    swApp.SetUserPreferenceStringListValue((int)swUserPreferenceStringListValue_e.swEmodelSelectionList, configString);

                    // get user setting
                    bool usersetting = swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swEDrawingsOkayToMeasure);

                    // change option
                    swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swEDrawingsOkayToMeasure, true);

                    // change folder path if save to production is checked
                    if (SaveToProductionCh.IsChecked == true)
                    {
                        folderPath = Path.GetDirectoryName (folderPath) + @"\Production";
                        if(!Directory.Exists (folderPath))
                        {
                            Directory.CreateDirectory (folderPath);
                        }
                    }

                    // create save string
                    string savePath = folderPath + @"\" + modelName + "_Rev" + resVal + ext;

                    // save the file
                    swModel.Extension.SaveAs3(savePath,
                        (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                        (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                        null, null, ref err, ref warn);

                    // change back option
                    swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swEDrawingsOkayToMeasure, usersetting);
                }
                else
                {
                    MessageBox.Show("No configuration is selected for Edrawing export.");
                }
            }

            // save step file
            if(TubeLaserCh.IsChecked == true)
            {
                if(TubeLaserLb.SelectedItems.Count > 0)
                {
                    // activate configuration
                    swModel.ShowConfiguration2(TubeLaserLb.SelectedItem.ToString());

                    // change folder path if save to production is checked
                    if (SaveToProductionCh.IsChecked == true)
                    {
                        folderPath = Path.GetDirectoryName(folderPath) + @"\Production\Tube Laser";
                        if (!Directory.Exists(folderPath))
                        {
                            Directory.CreateDirectory(folderPath);
                        }
                    }

                    // reconstruct modelname
                    if(modelName.Length > 13)
                    {
                        modelName = modelName.Substring(0, 10) + modelName.Substring(11, 3);
                        string savePath = folderPath + @"\" + modelName + resVal + ".STEP";

                        // save the file
                        swModel.Extension.SaveAs3(savePath,
                            (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                            (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                            null, null, ref err, ref warn);
                    }
                    else
                    {
                        MessageBox.Show("Invalid file name format for step file export, step file export failed.");
                    }
                }
                else
                {
                    MessageBox.Show("No configuration is selected for Tube Laser export.");
                }
            }

            Close();
            MessageBox.Show("Completed!");
        }
    }
}
