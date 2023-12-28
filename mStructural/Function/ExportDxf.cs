using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows;

namespace mStructural.Function
{
    public class ExportDxf
    {
        /// <summary>
        /// NOT USING
        /// </summary>

        #region Private Variables
        private SldWorks swApp;
        private Message msg;
        private Macros macros;
        private SwErrors swErrors;
        #endregion

        // constructor
        public ExportDxf(SWIntegration swIntegration)
        {
            swApp = swIntegration.SwApp;
            macros = new Macros(swApp);
            msg = new Message(swApp);
            swErrors = new SwErrors();
        }

        // Main method
        public void Run()
        {
            int err1 = -1;
            int warn1 = -1;
            int err2 = -1;
            int warn2 = -1;
            int errStep = -1;
            int warnStep = -1;
            bool bRet1;
            bool bRet2= false;

            // get active doc
            ModelDoc2 swModel = swApp.IActiveDoc2;

            bRet1 = macros.checkDocNullandType(swModel, swDocumentTypes_e.swDocDRAWING); if (!bRet1) return; // check

            // show confirm export message
            bool? toProduction;
            WPF.ProductionExportForm pef = new WPF.ProductionExportForm();
            bool? dialogResult = pef.ShowDialog();
            if (dialogResult == true)
            {
                toProduction = pef.ToProductionCh.IsChecked;
            }
            else
            {
                return;
            }

            //bool? amendQty;
            List<string> drawDependencises = new List<string>();
            WPF.AppendQuantity aq = new WPF.AppendQuantity();
            dialogResult = aq.ShowDialog();
            if(dialogResult == true)
            {
                //amendQty = aq.AppendQtyCh.IsChecked;

                // get unique dependencies of this drawing
                Hashtable drawDependsTable = getDependency(swModel.GetPathName());

                // Create a list and store the unique dependencies into it
                foreach (DictionaryEntry depend in drawDependsTable)
                {
                    drawDependencises.Add(depend.Value.ToString());
                }
            }
            else
            {
                //amendQty = true;
            }

            // initialize custom prop manager
            CustomPropertyManager swCusPropMgr = swModel.Extension.CustomPropertyManager[""];

            // get revision custom prop
            string val = "";
            string valout = "";
            bool bRet = swCusPropMgr.Get4("Revision", false, out val, out valout);

            // initialize string builder
            StringBuilder sb = new StringBuilder();

            // cast modeldoc to drawingdoc
            DrawingDoc swDraw = (DrawingDoc)swModel;

            // set dxf export option to export active sheet only
            bRet1 = swApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfMultiSheetOption, (int)swDxfMultisheet_e.swDxfActiveSheetOnly);

            // get current model name
            string sFileName = Path.GetFileNameWithoutExtension(swModel.GetPathName());

            // get current path name
            string sPathName = Path.GetDirectoryName(swModel.GetPathName());
            
            // find the \
            int index = sPathName.LastIndexOf("\\");

            string fpPathName = sPathName;
            string ppiPathName = sPathName;
            string ppePathName = sPathName;
            if (toProduction == true)
            {
                // add production and flat plates in path
                fpPathName = fpPathName.Substring(0, index) + "\\Production\\Flat Plates";

                // add production and pressed plates - inhouse fab in the path
                ppiPathName = ppiPathName.Substring(0,index) + "\\Production\\Pressed Plates - inhouse fab";

                // add production and pressed plates - external fab in the path
                ppePathName = ppePathName.Substring(0, index) + "\\Production\\Pressed Plates - external fab";
            }

            // get all sheet
            string[] vSheetName = (string[])swDraw.GetSheetNames();

            // process sheet
            for(int i = 0; i < vSheetName.Length; i++)
            {
                if (vSheetName[i].ToLower().Contains("dxf") || vSheetName[i].ToLower().Contains("th"))
                {
                    bRet1 = swDraw.ActivateSheet(vSheetName[i]);

                    // v1.8.0 modify the note if the text does not tally with the balloon
                    #region Check if note are tally with balloon
                    // get first view
                    View checkView = swDraw.IGetFirstView();
                    
                    // get next view which is the actual view
                    checkView = checkView.IGetNextView();
                    
                    // if there is view
                    while(checkView != null )
                    {
                        // get first note
                        Note checkNote = checkView.IGetFirstNote();
                        Note appendNote = checkNote;
                        string bomPNText = "";
                        string bomBNText = "";
                        string notePNText = "";
                        string noteBNText = "";
                        string originalNoteText = "";

                        // if there is note
                        while( checkNote != null)
                        {
                            string text = checkNote.GetText();

                            // check if it is balloon
                            if (checkNote.IsBomBalloon())
                            {
                                // get balloon text
                                bomPNText = text;
                                bomBNText = checkNote.GetBomBalloonText(false);
                            }
                            else if(text.Contains("PN:"))
                            {
                                // get copy
                                appendNote = checkNote;

                                // get note text
                                notePNText = text;
                                originalNoteText = notePNText;

                                // get position of newline
                                int newlineIndex = notePNText.IndexOf(System.Environment.NewLine);

                                // get only the part number
                                notePNText = notePNText.Substring(4, newlineIndex - 4);

                                // get position of quantity
                                int quantityIndex = originalNoteText.IndexOf("QTY:");

                                // get only the quantity number
                                noteBNText = originalNoteText.Substring(quantityIndex + 5, (originalNoteText.Length-4) - (quantityIndex+5));
                            }

                            // get next note
                            checkNote = checkNote.IGetNext();
                        }

                        // compare
                        if(bomPNText!= notePNText)
                        {
                            // change text if necessary
                            originalNoteText = originalNoteText.Replace("PN: " + notePNText, "PN: " + bomPNText);
                            appendNote.SetText(originalNoteText);

                            // log
                            sb.AppendLine("Amended part number item: " + bomPNText + " in View: " + checkView.Name + " in Sheet: " + vSheetName[i]);
                        }

                        /*if (amendQty == true)
                        {
                            // calculate quantity
                            CountTotalQuantityFromDrawingView countFunction = new CountTotalQuantityFromDrawingView(swApp);
                            int finalQty = countFunction.GetViewQuantity(swModel, checkView, drawDependencises);

                            // change text
                            originalNoteText = originalNoteText.Replace("QTY: " + noteBNText, "QTY: " + finalQty);
                            appendNote.SetText(originalNoteText);

                            // log
                            sb.AppendLine("Amended quantity of item: " + bomPNText + " in View: " + checkView.Name + " in Sheet: " + vSheetName[i]);
                        }*/

                        // get next view
                        checkView = checkView.IGetNextView();
                    }
                    #endregion

                    if (vSheetName[i].ToLower().EndsWith("p"))
                    {
                        // get file dxf save name
                        // v1.7 change to external
                        string dxfSaveName = ppePathName + "\\" + sFileName + "_"  +valout.ToUpper()+"_"+ vSheetName[i].ToLower() + ".dxf";

                        // make sure the directory exist before save
                        Directory.CreateDirectory(ppePathName);

                        // save
                        bRet1 = swModel.Extension.SaveAs3(dxfSaveName, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, false, ref err1, ref warn1);

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

                                // clear selection first
                                swModel.ClearSelection2(true);

                                // loop for bendlines
                                foreach(object bendline in bendlines)
                                {
                                    // cast to sketchsegment
                                    SketchSegment skSegBendline = (SketchSegment)bendline;
                                    
                                    // get sketch from sketch segment
                                    Sketch skBendline = skSegBendline.GetSketch();
                                    
                                    // cast it to feature
                                    Feature featBendline = (Feature)skBendline;

                                    // select the bendline sketch
                                    bRet = swModel.Extension.SelectByID2(featBendline.Name + "@" + viewDrawComp.Name + "@" + swView.Name, "SKETCH", 0, 0, 0, true, 0, null, 0);

                                    MessageBox.Show(bRet.ToString());
                                }

                                // hide it
                                swModel.BlankSketch();

                                // clear selection
                                swModel.ClearSelection2(true);
                            }
                            #endregion

                            // get the reference model name of the view
                            string modelName = swView.GetReferencedModelName();

                            // open the model in the back
                            ModelDoc2 viewModel = swApp.OpenDoc6(modelName, (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", errStep, warnStep);
                            
                            if (viewModel != null)
                            {
                                // if opened the model without any issue
                                // save the mode
                                bRet2 = viewModel.Extension.SaveAs3(stpSaveName, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, null, ref err2, ref warn2);
                                
                                // close the model 
                                swApp.CloseDoc(modelName);

                                // if saving failed
                                if (bRet2 == false) break;
                            }
                            else
                            {
                                // if failed to open the model
                                msg.ErrorMsg(swErrors.GetFileLoadError(errStep));
                                bRet2 = false;
                                return;
                            }
                            
                            // get next view if no saving error
                            swView = swView.IGetNextView();
                        }
                    }
                    else
                    {
                        // get file save name
                        string saveName = fpPathName + "\\" + sFileName + "_" + valout.ToUpper() +"_"+ vSheetName[i].ToLower() + ".dxf";

                        // make sure the directory exist before save
                        Directory.CreateDirectory(fpPathName);

                        // save
                        bRet1 = swModel.Extension.SaveAs3(saveName, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, null, ref err1, ref warn1);

                        bRet2 = true;
                    }

                    // if save failed
                    if (!bRet1)
                    {
                        msg.ErrorMsg(swErrors.GetFileSaveError(err1));
                    }

                    // if save failed
                    if (!bRet2)
                    {
                        msg.ErrorMsg(swErrors.GetFileSaveError(err2));
                    }
                }
            }

            // switch back to first sheet
            bRet1 = swDraw.ActivateSheet(vSheetName[0]);

            // Complete
            msg.InfoMsg("Save DXF completed!\r\n" + sb.ToString());
        }

        // method to get dependency
        private Hashtable getDependency(string file)
        {
            Hashtable dependencies = new Hashtable();
            string[] depends = (string[])swApp.GetDocumentDependencies2(file, true, true, false);
            if (depends != null)
            {
                int index = 0;
                while (index < depends.GetUpperBound(0))
                {
                    try
                    {
                        dependencies.Add(depends[index], depends[index + 1]);
                    }
                    catch { }
                    index += 2;
                }
            }

            return dependencies;
        }
    }
}
