using mStructural.Classes;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Text;
using System.Windows;

namespace mStructural.Function
{
    public class ExportDxf2
    {
        public StringBuilder Run(SldWorks swapp, string drawingpath, bool isproduction, ObservableCollection<ExportDxfClass> dxfcol, string rev)
        {
            StringBuilder sb = new StringBuilder();

            int err = -1;
            int warn = -1;

            // get drawing model
            ModelDoc2 swModel = swapp.OpenDoc6(drawingpath, (int)swDocumentTypes_e.swDocDRAWING, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref err, ref warn);

            // cast to drawing doc
            DrawingDoc swDraw = swModel as DrawingDoc;

            // set dxf export option to export active sheet only
            swapp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfMultiSheetOption, (int)swDxfMultisheet_e.swDxfActiveSheetOnly);

            // get current model name
            string sFileName = Path.GetFileNameWithoutExtension(swModel.GetPathName());

            // get current path name
            string sPathName = Path.GetDirectoryName(swModel.GetPathName());

            // find the \
            int index = sPathName.LastIndexOf("\\");

            string fpPathName = sPathName;
            string ppiPathName = sPathName;
            string ppePathName = sPathName;
            if (isproduction == true)
            {
                // add production and flat plates in path
                fpPathName = fpPathName.Substring(0, index) + "\\Production\\Flat Plates";

                // add production and pressed plates - inhouse fab in the path
                ppiPathName = ppiPathName.Substring(0, index) + "\\Production\\Pressed Plates - inhouse fab";

                // add production and pressed plates - external fab in the path
                ppePathName = ppePathName.Substring(0, index) + "\\Production\\Pressed Plates - external fab";
            }

            // process each sheet
            foreach(ExportDxfClass edc in dxfcol)
            {
                if (edc.Include)
                {
                    // activate the sheet
                    swDraw.ActivateSheet(edc.SheetName);

                    if (edc.SheetName.ToUpper().EndsWith("P"))
                    {
                        string dxfSaveName = ppePathName + "\\" + sFileName + "_" + rev + "_" + edc.SheetName.ToLower() + ".dxf";

                        MessageBox.Show(dxfSaveName);

                        // make sure the directory exist before save
                        Directory.CreateDirectory(ppePathName);

                        // get step file save name
                        // v1.7 changed to internal
                        string stpSaveName = ppiPathName + "\\" + sFileName + ".step";

                        // make sure the directory exist before save
                        Directory.CreateDirectory(ppiPathName);

                        // get the model in the view, first view will always be sheet format
                        View swView = swDraw.IGetFirstView();

                        // get next view which is the actual view
                        swView = swView.IGetNextView();

                        // if there is a view
                        while (swView != null)
                        {
                            #region v1.15 hide bendlines
                            // check if there is any bendline
                            if (swView.GetBendLineCount() > 0)
                            {
                                // get root drawing component
                                DrawingComponent viewDrawComp = swView.RootDrawingComponent;

                                // get all bendlines
                                object[] bendlines = swView.GetBendLines();

                                // loop for bendlines
                                foreach (object bendline in bendlines)
                                {
                                    // cast to sketchsegment
                                    SketchSegment skSegBendline = (SketchSegment)bendline;

                                    // get sketch from sketch segment
                                    Sketch skBendline = skSegBendline.GetSketch();

                                    // cast it to feature
                                    Feature featBendline = (Feature)skBendline;

                                    // select the bendline sketch
                                    bool bRet = swModel.Extension.SelectByID2(featBendline.Name + "@" + viewDrawComp.Name + "@" + swView.Name, "SKETCH", 0, 0, 0, false, 0, null, 0);

                                    // hide it
                                    swModel.BlankSketch();
                                }
                            }
                            #endregion

                            // get the reference model name of the view
                            string modelName = swView.GetReferencedModelName();

                            // open the model in the back
                            ModelDoc2 viewModel = swapp.OpenDoc6(modelName, (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", err, warn);

                            if (viewModel != null)
                            {
                                // if opened the model without any issue
                                // save the model
                                viewModel.Extension.SaveAs3(stpSaveName, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, null, ref err, ref warn);

                                // close the model 
                                swapp.CloseDoc(modelName);
                            }
                            
                            // get next view if no saving error
                            swView = swView.IGetNextView();
                        }

                        // save
                        swModel.Extension.SaveAs3(dxfSaveName, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, null, ref err, ref warn);
                    }
                    else
                    {
                        // get file save name
                        string saveName = fpPathName + "\\" + sFileName + "_" + rev + "_" + edc.SheetName + ".dxf";

                        // make sure the directory exist before save
                        Directory.CreateDirectory(fpPathName);

                        swModel.Extension.SaveAs3(saveName, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, null, ref err, ref warn);
                    }
                }
            }

            // activate first sheet
            swDraw.ActivateSheet(dxfcol[0].SheetName);

            // log
            Log(sb, "DXF exported.");

            return sb;
        }

        private void Log(StringBuilder sb, string text)
        {
            sb.AppendLine(DateTime.Now.ToString() + " : " + text);
        }
    }
}
