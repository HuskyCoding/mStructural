using mStructural.Classes;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Text;

namespace mStructural.Function
{
    public class ExportModel2
    {
        public StringBuilder Run(SldWorks swapp, string modelPath, bool isproduction, bool isedraw, bool istubelaser, ObservableCollection<ExportModelClass> emCol, string tlConfig, string rev)
        {
            StringBuilder sb = new StringBuilder();
            Log(sb, "Initiate Model Export.");

            int iDocType = -1;
            int err = -1;
            int warn = -1;
            string ext = "";

            switch (modelPath.Substring(modelPath.Length - 6).ToUpper())
            {
                case "SLDASM":
                    {
                        iDocType = (int)swDocumentTypes_e.swDocASSEMBLY;
                        ext = ".easm";
                        break;
                    }
                case "SLDPRT":
                    {
                        iDocType = (int)swDocumentTypes_e.swDocPART;
                        ext = ".eprt";
                        break;
                    }
            }

            ModelDoc2 swModel = swapp.OpenDoc6(modelPath, iDocType, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref err, ref warn);

            // model name
            string modelName = Path.GetFileNameWithoutExtension(modelPath);

            // directory
            string folderPath = Path.GetDirectoryName(modelPath);

            // save edrawing
            if (isedraw)
            {
                StringBuilder configSb = new StringBuilder();
                foreach(ExportModelClass emc in emCol)
                {
                    if (emc.IsChecked)
                    {
                        configSb.Append(emc.Configuration);
                        configSb.Append("\n");
                    }
                }

                string configString = configSb.ToString();

                // set config string to export option
                swapp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swEdrawingsSaveAsSelectionOption, (int)swEdrawingSaveAsOption_e.swEdrawingSaveSelected);
                swapp.SetUserPreferenceStringListValue((int)swUserPreferenceStringListValue_e.swEmodelSelectionList, configString);

                // get user setting
                bool usersetting = swapp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swEDrawingsOkayToMeasure);

                // change option
                swapp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swEDrawingsOkayToMeasure, true);

                if (isproduction)
                {
                    folderPath = Path.GetDirectoryName(folderPath) + @"\Production";
                    if (!Directory.Exists(folderPath))
                    {
                        Directory.CreateDirectory(folderPath);
                    }
                }

                // create save string
                string savePath = folderPath + @"\" + modelName + "_Rev" + rev + ext;

                // activate the main assembly for edrawing export
                swapp.ActivateDoc2(swModel.GetPathName(), true, ref err);

                // save the file
                swModel.Extension.SaveAs3(savePath,
                    (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                    (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                    null, null, ref err, ref warn);

                // change back option
                swapp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swEDrawingsOkayToMeasure, usersetting);

                // close the activated main assembly
                swapp.CloseDoc(swModel.GetPathName());

                Log(sb, "Exported edrawing(s)...");
            }

            // save tube laser step file
            if (istubelaser)
            {
                // activate configuration
                swModel.ShowConfiguration2(tlConfig);

                // change folder path if save to production is checked
                if (isproduction)
                {
                    folderPath = Path.GetDirectoryName(folderPath) + @"\Production\Tube Laser";
                    if (!Directory.Exists(folderPath))
                    {
                        Directory.CreateDirectory(folderPath);
                    }
                }

                // reconstruct modelname
                if (modelName.Length > 13)
                {
                    modelName = modelName.Substring(0, 10) + modelName.Substring(11, 3);
                    string savePath = folderPath + @"\" + modelName + rev + ".STEP";

                    // save the file
                    swModel.Extension.SaveAs3(savePath,
                        (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                        (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                        null, null, ref err, ref warn);

                    Log(sb, "Exported tube laser step file...");
                }
                else
                {
                    Log(sb, "Invalid file name format for step file export, step file export failed.");
                }
            }

            return sb;
        }
        private void Log(StringBuilder sb, string text)
        {
            sb.AppendLine(DateTime.Now.ToString() + " : " + text);
        }
    }
}
