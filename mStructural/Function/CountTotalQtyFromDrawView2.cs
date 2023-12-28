using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Diagnostics;
using System;
using System.Windows;

namespace mStructural.Function
{
    public class CountTotalQtyFromDrawView2
    {
        public int GetBodyQuantityFromView(SldWorks swApp, View swview, string topLevelAsmPath, string selectedConfig)
        {
            try
            {
                int finalQty = -1;
                int err = -1;
                int warn = -1;

                // get the modeldoc2 of the view
                ModelDoc2 viewModel = swview.ReferencedDocument;

                // get the path for this model
                string viewModelPath = viewModel.GetPathName();

                // get bodies fro the selected view
                object[] vBodies = (object[])swview.Bodies;

                // get the configuration using by the view
                string viewConfigName = swview.ReferencedConfiguration;

                // get the view body
                Body2 swBody = null;
                if (vBodies != null)
                {
                    if (vBodies.Length != 1)
                    {
                        MessageBox.Show(swview.Name + " contains more than a solid body.");
                        return -1;
                    }
                    else
                    {
                        swBody = (Body2)vBodies[0];
                    }
                }
                else
                {
                    // if not multi body, check if it is flat pattern
                    if (viewConfigName.Contains("FLAT-PATTERN"))
                    {
                        // cahnge to the same configuration as the view
                        viewModel.ShowConfiguration2(viewConfigName);

                        // activate the document
                        swApp.ActivateDoc3(viewModel.GetPathName(), true, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, ref err);

                        // get first feature for traverse
                        Feature swFeat = viewModel.IFirstFeature();

                        // set stopping point
                        bool bodyFound = false;

                        // variable for the name of the body folder to select later
                        string flatPatternBodyFolderName = "";
                        BodyFolder swBodyFolder = default(BodyFolder);

                        // traverse feature
                        while (swFeat != null)
                        {
                            // if it is cut list or solidbody folder
                            if (swFeat.GetTypeName2() == "CutListFolder" || swFeat.GetTypeName2() == "SolidBodyFolder")
                            {
                                // get the bodyfolder from feature
                                swBodyFolder = (BodyFolder)swFeat.GetSpecificFeature2();

                                // get all bodies from this body folder
                                object[] featBodies = (object[])swBodyFolder.GetBodies();
                                if (featBodies != null)
                                {
                                    // loop for the bodies
                                    foreach (object featBody in featBodies)
                                    {
                                        // cast the object to Body2
                                        Body2 swFeatBody = (Body2)featBody;

                                        // check visibility of the body
                                        if (swFeatBody.Visible)
                                        {
                                            // if it is visible, then this body folder contains the flat pattern body
                                            flatPatternBodyFolderName = swBodyFolder.GetFeature().Name;

                                            // set stopping point
                                            bodyFound = true;

                                            // stop loop
                                            break;
                                        }
                                    }
                                }
                            }

                            // if hit stopping point
                            if (bodyFound)
                            {
                                // stop traverse
                                break;
                            }

                            // if didnt hit stopping point, go to next feature
                            swFeat = swFeat.IGetNextFeature();
                        }

                        // get the parent config from the view model config
                        Configuration flatPatternConfig = (Configuration)viewModel.GetConfigurationByName(viewConfigName);
                        Configuration parentConfig = flatPatternConfig.GetParent();

                        // change to parent configuration
                        viewModel.ShowConfiguration2(parentConfig.Name);

                        // select the body folder 
                        SelectionMgr viewModelSelMgr = viewModel.ISelectionManager;
                        viewModel.Extension.SelectByID2(flatPatternBodyFolderName, "BDYFOLDER", 0, 0, 0, false, 0, null, 0);

                        // cast the selected object to feature and get the bodyfolder from feature
                        swFeat = (Feature)viewModelSelMgr.GetSelectedObject6(1, -1);
                        swBodyFolder = (BodyFolder)swFeat.GetSpecificFeature2();

                        // get bodies from the bodyfolder
                        object[] vFlatBodies = (object[])swBodyFolder.GetBodies();
                        swBody = (Body2)vFlatBodies[0];

                        // close activated doc
                        swApp.CloseDoc(viewModel.GetPathName());
                    }
                    else
                    {
                        // if not flat pattern show the specific configuration
                        viewModel.ShowConfiguration2(viewConfigName);

                        // cast modeldoc to partdoc to access getbodies method
                        PartDoc swPart = (PartDoc)viewModel;

                        // get all the bodies in this part
                        object[] partBodies = (object[])swPart.GetBodies2((int)swBodyType_e.swSolidBody, true);

                        // since it only have a body in the view and we will assume the index is 0
                        swBody = (Body2)partBodies[0];
                    }
                }

                if (topLevelAsmPath != "")
                {
                    // open the highest level assembly and traverse
                    ModelDoc2 tModelDoc = swApp.OpenDoc6(topLevelAsmPath, (int)swDocumentTypes_e.swDocASSEMBLY, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref err, ref warn);
                    //Configuration tRootConfig = tModelDoc.IGetActiveConfiguration();
                    Configuration tRootConfig = tModelDoc.GetConfigurationByName(selectedConfig);
                    Component2 tRootComp = tRootConfig.IGetRootComponent2();
                    finalQty = TraverseComponentForCount(tRootComp, swBody, 0, viewModelPath);
                }
                else
                {
                    PartDoc swPart = (PartDoc)viewModel;
                    object[] partBodies = (object[])swPart.GetBodies2((int)swBodyType_e.swSolidBody, true);
                    finalQty = 0;
                    foreach (object partBody in partBodies)
                    {
                        MathTransform swXfm;
                        bool bRet = swBody.GetCoincidenceTransform(partBody, out swXfm);
                        if (bRet)
                        {
                            finalQty++;
                        }
                    }
                    return finalQty;
                }

                return finalQty;
            }
            catch (Exception ex)
            {
                // Get stack trace for the exception with source file information
                var st = new StackTrace(ex, true);
                // Get the top stack frame
                var frame = st.GetFrame(0);
                // Get the line number from the stack frame
                var line = frame.GetFileLineNumber();
                MessageBox.Show("Line " + line.ToString() + " : " + ex.ToString());
                return -1;
            }
        }

        // method to Traverse for counting
        public int TraverseComponentForCount(Component2 swcomp, Body2 swbody, int count, string viewModelPath)
        {
            int Count = count;

            // get all child
            object[] childComp = (object[])swcomp.GetChildren();

            if (childComp != null)
            {
                // Traverse
                foreach (object child in childComp)
                {
                    // cast to component2 type object
                    Component2 swChildComp = (Component2)child;

                    // get all bodies
                    object vBodiesInfo;
                    object[] vBodies = (object[])swChildComp.GetBodies3((int)swBodyType_e.swSolidBody, out vBodiesInfo);

                    if (vBodies != null)
                    {
                        // get the modeldoc for this component
                        ModelDoc2 swChildModelDoc = swChildComp.IGetModelDoc();

                        if (swChildModelDoc != null)
                        {
                            // get the path name for this modeldoc
                            string childModelDocPath = swChildModelDoc.GetPathName();

                            MathTransform swXfm;
                            foreach (object vBody in vBodies)
                            {
                                bool bRet = swbody.GetCoincidenceTransform(vBody, out swXfm);
                                if (bRet)
                                {
                                    //if(childModelDocPath == viewModelPath)
                                    //{
                                    Count++;
                                    //}
                                }
                            }
                        }
                    }

                    Count = TraverseComponentForCount(swChildComp, swbody, Count, viewModelPath);
                }
            }

            return Count;
        }
    }
}
