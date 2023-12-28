using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Windows;

namespace mStructural.Function
{
    public class CountTotalQuantityFromDrawingView
    {
        #region Private Variables
        private SldWorks swApp;
        #endregion

        // constructor
        public CountTotalQuantityFromDrawingView(SldWorks swapp)
        {
            swApp = swapp;
        }

        // main method to get quantity
        public int GetViewQuantity(ModelDoc2 swmodel, View swview, List<string> drawdependencies)
        {
            int finalQty = -1;

            // get the modeldoc2 of the view
            ModelDoc2 viewModel = swview.ReferencedDocument;

            // get the path for this model
            string viewModelPath = viewModel.GetPathName();

            // get the related parent of the model 
            List<string> relatedParents = new List<string>();
            foreach(string drawDepend in drawdependencies)
            {
                Hashtable dependDependsTable = getDependency(drawDepend);
                if (dependDependsTable.ContainsValue(viewModelPath))
                {
                    relatedParents.Add(drawDepend);
                }
            }

            // create parent list with level hierarchy
            List<Tuple<int, string>> finalParents = new List<Tuple<int, string>>();
            foreach (string parent in relatedParents)
            {
                int err = -1;
                int warn = -1;
                ModelDoc2 parentModel = swApp.OpenDoc6(parent, (int)swDocumentTypes_e.swDocASSEMBLY, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref err, ref warn);
                Configuration parentConfig = parentModel.IGetActiveConfiguration();
                Component2 RootComp = parentConfig.IGetRootComponent2();
                TraverseComponent(RootComp, 1, viewModelPath, finalParents, parent);
            }

            // convert the list to array and sort it from lower to highest level
            Tuple<int, string>[] finalParentsArr = finalParents.ToArray();
            Array.Sort(finalParentsArr, new MyTupleComparer());

            // calculate qty in part
            int bodyCount = 0;

            // get the bodies for the selected view
            object[] vBodies = (object[])swview.Bodies;

            // get the configuration using by the view
            string viewConfigName = swview.ReferencedConfiguration;

            // if not multibody
            if (vBodies!=null)
            {
                if (vBodies.Length != 1)
                {
                    MessageBox.Show(swview.Name +  " contains more than a solid body.");
                    return -1;
                }
                else
                {
                    // since only support single body, get the first one
                    Body2 swBody = (Body2)vBodies[0];

                    // change configuration of the model to use the same configuration as view
                    viewModel.ShowConfiguration2(viewConfigName);

                    // get first feature 
                    Feature swFeat = viewModel.IFirstFeature();
                    
                    // this variable is to set a stopping point for the traverse if found match earlier
                    bool bodyFound = false;

                    // traverse feature
                    while (swFeat != null)
                    {
                        // if it is body folder
                        if(swFeat.GetTypeName2() == "CutListFolder" || swFeat.GetTypeName2() == "SolidBodyFolder")
                        {
                            // get the body folder from feature
                            BodyFolder swBodyFolder = (BodyFolder)swFeat.GetSpecificFeature2();

                            // get all bodies in the body folder
                            object[] featBodies = (object[])swBodyFolder.GetBodies();

                            // if there is body
                            if(featBodies!=null)
                            {
                                // loop each found body in the body folder
                                foreach(object featBody in featBodies)
                                {
                                    // cast to Body2
                                    Body2 swFeatBody = (Body2)featBody;

                                    // check the name of this body get from feature traversal
                                    // whether it is the same as the body used in view
                                    if(swFeatBody.Name == swBody.Name)
                                    {
                                        // if same, get the bodycount   
                                        bodyCount = swBodyFolder.GetBodyCount();

                                        // set stopping point
                                        bodyFound = true;

                                        // stop the loop
                                        break;
                                    }
                                }
                            }
                        }

                        // if hit stopping point, stop traverse
                        if (bodyFound) break;

                        // get next feature if no stopping point hit
                        swFeat = swFeat.IGetNextFeature();
                    }
                }
            }
            else
            {
                // if not multi body, check if it is flat pattern
                if (viewConfigName.Contains("FLAT-PATTERN"))
                {
                    int err = -1;
                    
                    // cahnge to the same configuration as the view
                    viewModel.ShowConfiguration2(viewConfigName);

                    // activate the document
                    swApp.ActivateDoc3(viewModel.GetPathName(),true,(int)swOpenDocOptions_e.swOpenDocOptions_Silent, ref err);
                    
                    // get first feature for traverse
                    Feature swFeat = viewModel.IFirstFeature();

                    // set stopping point
                    bool bodyFound = false;

                    // variable for the name of the body folder to select later
                    string flatPatternBodyFolderName = "";
                    BodyFolder swBodyFolder = default(BodyFolder) ;

                    // traverse feature
                    while (swFeat != null)
                    {
                        // if it is cut list or solidbody folder
                        if(swFeat.GetTypeName2() == "CutListFolder" || swFeat.GetTypeName2() == "SolidBodyFolder")
                        {
                            // get the bodyfolder from feature
                            swBodyFolder = (BodyFolder)swFeat.GetSpecificFeature2();
                            
                            // get all bodies from this body folder
                            object[]featBodies = (object[])swBodyFolder.GetBodies();
                            if(featBodies!=null)
                            {
                                // loop for the bodies
                                foreach(object featBody in featBodies)
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

                    // get the body count
                    bodyCount = swBodyFolder.GetBodyCount();
                    
                    // close the activated doc
                    swApp.CloseDoc(viewModel.GetPathName());
                }
                else
                {
                    bodyCount = 1;
                }

            }

            // Traverse sub assemblies
            int tErr = -1;
            int tWarn = -1;

            if(finalParents.Count == 0)
            {
                // this is highest level 
                finalQty = bodyCount;
            }
            else
            {
                // open the highest level assembly and traverse
                ModelDoc2 tModelDoc = swApp.OpenDoc6(finalParentsArr[finalParentsArr.Length-1].Item2, (int)swDocumentTypes_e.swDocASSEMBLY, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref tErr, ref tWarn);
                Configuration tRootConfig = tModelDoc.IGetActiveConfiguration();
                Component2 tRootComp = tRootConfig.IGetRootComponent2();
                finalQty = TraverseComponentForCount(tRootComp, 0, viewModelPath, viewConfigName);

                finalQty = finalQty * bodyCount;
            }


            return finalQty;
        }

        // method to count model in assembly level 
        private int TraverseComponentForCount(Component2 swcomp, int count, string viewmodelpath, string viewconfigname)
        {
            int Count = count;

            // get all child
            object[] childComp = (object[])swcomp.GetChildren();

            if (childComp != null)
            {
                // Traverse
                foreach (object child in childComp)
                {
                    // cast to Component2
                    Component2 swChildComp = (Component2)child;

                    // get the modeldoc2 of this child
                    ModelDoc2 childModel = (ModelDoc2)swChildComp.GetModelDoc2();
                    if (childModel != null)
                    {
                        // calculate how many times the model appears in the assembly
                        if (childModel.GetPathName() == viewmodelpath)
                        {
                            // if the view is flat pattern, count regardless of configuration
                            if (viewconfigname.Contains("FLAT-PATTERN"))
                            {
                                // get parent configuration
                                Configuration swConfig = childModel.IGetConfigurationByName(viewconfigname);
                                Configuration parentConfig = swConfig.GetParent();
                                string parentConfigName = parentConfig.Name;

                                // only increment if same as the parent config of the flat pattern
                                if(swChildComp.ReferencedConfiguration == parentConfigName)
                                {
                                    Count++;
                                }
                            }
                            // else only count if same configuration with view
                            else if(swChildComp.ReferencedConfiguration == viewconfigname)
                            {
                                Count++;
                            }
                        }
                    }

                    // traverse child
                    Count = TraverseComponentForCount(swChildComp, Count, viewmodelpath, viewconfigname);
                }
            }

            return Count;
        }

        // method to traverse assembly to create Tuple
        private void TraverseComponent(Component2 swcomp, int level, string viewmodelpath, List<Tuple<int,string>> finalparents, string parent)
        {
            // get all children
            object[] childComp = (object[])swcomp.GetChildren();

            if(childComp != null)
            {
                // traverse child component
                foreach(object child in childComp)
                {
                    // cast child object to Component2
                    Component2 swChildComp = (Component2)child;

                    // get modeldoc2 object from the component
                    ModelDoc2 childModel = (ModelDoc2)swChildComp.GetModelDoc2();

                    // check if it is the same as the model in view through path on disk
                    if(childModel != null)
                    {
                        if(childModel.GetPathName() == viewmodelpath)
                        {
                            // if it is the same, add it to the tuple
                            finalparents.Add(new Tuple<int, string>(level, parent));
                            return;
                        }
                    }

                    // traverse child comp
                    TraverseComponent(swChildComp, level + 1, viewmodelpath, finalparents, parent);
                }
            }
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

        // method for sorting
        public class MyTupleComparer : Comparer<Tuple<int, string>>
        {
            public override int Compare(Tuple<int, string> x, Tuple<int, string> y)
            {
                return x.Item1.CompareTo(y.Item1);
            }
        }
    }
}
