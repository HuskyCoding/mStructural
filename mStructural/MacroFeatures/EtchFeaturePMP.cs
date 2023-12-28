using mStructural.Function;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;

namespace mStructural.MacroFeatures
{
    public class EtchFeaturePMP: PropertyManagerPage2Handler9
    {
        #region Private Variables
        private SldWorks swApp;
        private ModelDoc2 swModel;
        private Feature swFeat;
        private MacroFeatureData swFeatData;

        private PropertyManagerPage2 pmp;
        private PropertyManagerPageGroup pmpGroup1;
        private PropertyManagerPageGroup pmpGroup2;
        private PropertyManagerPageSelectionbox pmpSb;
        private PropertyManagerPageLabel pmpLabel;
        private PropertyManagerPageTextbox pmpTb;

        private const int GroupId1 = 1;
        private const int GroupId2 = 2;
        private const int SelectionboxId = 3;
        private const int LabelId = 4;
        private const int TextboxId = 5;

        private Macros macros;
        private Message msg;

        private int State; // 0 is new, 1 is edit
        private string Text = "etch";
        private int mark = 1;

        // for selection list
        //private int preSelCount = 0;
        private List<Feature> preSelFeats;
        //private List<Feature> postSelFeats;
        #endregion

        // constructor
        public EtchFeaturePMP(SldWorks swapp, int state, Feature feat)
        {
            swApp = swapp;
            swModel = swApp.IActiveDoc2;
            macros = new Macros(swApp);
            if (!macros.checkDocNullandType(swModel, swDocumentTypes_e.swDocPART)) return; // check
            swFeat = feat;
            State = state;
            msg = new Message(swApp);
            preSelFeats = new List<Feature>();
            //postSelFeats = new List<Feature>();

            // pmp variables
            string pageTitle = "";
            string caption = "";
            string tip = null;
            int options = 0;
            int err = 0;
            int controlType = 0;
            int alignment = 0;

            // set page
            pageTitle = "Etch Feature";
            options = (int)swPropertyManagerButtonTypes_e.swPropertyManager_OkayButton +
                (int)swPropertyManagerButtonTypes_e.swPropertyManager_CancelButton +
                (int)swPropertyManagerPageOptions_e.swPropertyManagerOptions_LockedPage +
                (int)swPropertyManagerPageOptions_e.swPropertyManagerOptions_PushpinButton;
            pmp = (PropertyManagerPage2)swApp.CreatePropertyManagerPage(pageTitle, options, this, ref err);

            // make sure page is created successfully
            if(err == (int)swPropertyManagerPageStatus_e.swPropertyManagerPage_Okay)
            {
                // add Groups
                caption = "Features";
                options = (int)swAddGroupBoxOptions_e.swGroupBoxOptions_Visible +
                    (int)swAddGroupBoxOptions_e.swGroupBoxOptions_Expanded;
                pmpGroup1 = (PropertyManagerPageGroup)pmp.AddGroupBox(GroupId1, caption, options);

                caption = "Settings";
                pmpGroup2 = (PropertyManagerPageGroup)pmp.AddGroupBox(GroupId2, caption, options);

                // add selection box
                controlType = (int)swPropertyManagerPageControlType_e.swControlType_Selectionbox;
                caption = "";
                alignment = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_Indent;
                options = (int)swAddControlOptions_e.swControlOptions_Visible +
                    (int)swAddControlOptions_e.swControlOptions_Enabled;
                tip = "Select features";
                pmpSb = (PropertyManagerPageSelectionbox)pmpGroup1.AddControl(SelectionboxId, (short)controlType, caption, (short)alignment, options, tip);
                swSelectType_e[] filters = new swSelectType_e[2];
                filters[0] = swSelectType_e.swSelBODYFEATURES;
                filters[1] = swSelectType_e.swSelFACES;
                object filterObj = null;
                filterObj = filters;
                pmpSb.SingleEntityOnly = false;
                pmpSb.AllowMultipleSelectOfSameEntity = false;
                pmpSb.Height = 50;
                pmpSb.SetSelectionFilters(filters);

                // add label
                controlType = (int)swPropertyManagerPageControlType_e.swControlType_Label;
                caption = "Text";
                pmpLabel = (PropertyManagerPageLabel)pmpGroup2.AddControl2(LabelId, (short)controlType, caption, (short)alignment, options, tip);

                // add text for etching
                controlType = (int)swPropertyManagerPageControlType_e.swControlType_Textbox;
                tip = "Text to mark on entities";
                pmpTb = (PropertyManagerPageTextbox)pmpGroup2.AddControl(TextboxId, (short)controlType, caption, (short)(alignment), options, tip);
                pmpTb.Text = Text;

                // set selection if user editing feature
                if(State == 1)
                {
                    // get feature data
                    swFeatData = (MacroFeatureData)swFeat.GetDefinition();

                    // access to feature data
                    swFeatData.AccessSelections(swModel, null);

                    // get selection from this feature data
                    object sels;
                    object types;
                    object marks;
                    swFeatData.GetSelections(out sels, out types, out marks);
                    if (sels != null)
                    {
                        object[] selArr = (object[])sels;
                        foreach (object sel in selArr)
                        {
                            // cast selected object to feature
                            Feature swFeat = (Feature)sel;

                            // select the feature
                            swFeat.Select2(true, 0);

                            // add it to pre select feature list
                            preSelFeats.Add(swFeat);

                            // counter for pre select count
                            // preSelCount++;
                        }
                    }
                }
            }
            else
            {
                // if creating page failed
                msg.ErrorMsg("Failed to create Etch Feature Property Manager Page");
            }
        }

        // Method to show
        public void Show()
        {
            pmp.Show();
        }

        public void OnClose(int Reason)
        {
            if(Reason == (int)swPropertyManagerPageCloseReasons_e.swPropertyManagerPageClose_Okay)
            {
                if(State == 0) // new feature
                {
                    // setup parameter
                    string[] paramNames = new string[2];
                    int[] paramTypes = new int[2];
                    string[] paramValues = new string[2];

                    paramNames[0] = "Text";
                    paramTypes[0] = (int)swMacroFeatureParamType_e.swMacroFeatureParamTypeString;
                    paramValues[0] = Text;

                    paramNames[1] = "State";
                    paramTypes[1] = (int)swMacroFeatureParamType_e.swMacroFeatureParamTypeInteger;
                    paramValues[1] = State.ToString();

                    // icon path
                    string[] iconFiles = new string[3];
                    string assemblyFolder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                    iconFiles[0] = Path.Combine(assemblyFolder, "Macro Feature Icon\\EtchCut.bmp");
                    iconFiles[1] = Path.Combine(assemblyFolder, "Macro Feature Icon\\EtchCut_S.bmp");
                    iconFiles[2] = Path.Combine(assemblyFolder, "Macro Feature Icon\\EtchCut.bmp");

                    // instantiate feature manager
                    FeatureManager featMgr = swModel.FeatureManager;

                    // create feature
                    Feature mFeat = featMgr.InsertMacroFeature3("EtchFeature", "mStructural.MacroFeatures.EtchFeatureMF", null,
                        paramNames, paramTypes, paramValues, null, null, null,
                        iconFiles, (int)swMacroFeatureOptions_e.swMacroFeatureByDefault) ;
                }
                else if (State == 1) // edit feature
                {
                    // clear all edge and colour on old feature
                    // get material properties colour
                    double[] featColour = (double[])swModel.MaterialPropertyValues;

                    // For each feature that is removed
                    foreach (Feature deSelFeat in preSelFeats)
                    {
                        // get all faces from the feature
                        object[] vFaces = (object[])deSelFeat.GetFaces();
                        foreach (object vFace in vFaces)
                        {
                            // cast object as face2
                            Face2 swFace = (Face2)vFace;

                            // set colour of this deleted face
                            swFace.MaterialPropertyValues = featColour;

                            // get all edges from this face
                            object[] vEdges = (object[])swFace.GetEdges();
                            foreach (object vEdge in vEdges)
                            {
                                // cast object as Edge
                                Edge swEdge = (Edge)vEdge;

                                // cast edge as entity
                                Entity swEnt = (Entity)swEdge;

                                // delete the entity name
                                ((PartDoc)swModel).DeleteEntityName(swEnt);
                            }
                        }
                    }

                    // setup parameter
                    string[] paramNames = new string[2];
                    int[] paramTypes = new int[2];
                    string[] paramValues = new string[2];

                    paramNames[0] = "Text";
                    paramTypes[0] = (int)swMacroFeatureParamType_e.swMacroFeatureParamTypeString;
                    paramValues[0] = Text;

                    paramNames[1] = "State";
                    paramTypes[1] = (int)swMacroFeatureParamType_e.swMacroFeatureParamTypeInteger;
                    paramValues[1] = State.ToString();

                    // set the parameter to feature data
                    swFeatData.SetParameters(paramNames, paramTypes, paramValues);

                    // Instantiate selection manager
                    SelectionMgr swSelMgr = swModel.ISelectionManager;

                    // get selected object count
                    int selCount = swSelMgr.GetSelectedObjectCount2(mark);

                    // instantiate objects and dispatch wrappers
                    object[] selObj = new object[selCount];
                    DispatchWrapper[] dw = new DispatchWrapper[selCount];

                    // wrap each selected object to dispatchwrapper
                    // this step is to ensure if there is different type of object, 
                    // they can be passed into the feature data
                    for (int i = 1; i < selCount + 1; i++)
                    {
                        // get selected object
                        selObj[i - 1] = swSelMgr.GetSelectedObject6(i, mark);

                        // create dispatchwrapper from the selected object
                        dw[i - 1] = new DispatchWrapper(selObj[i - 1]);
                    }

                    // enumerate a list of integer as a placeholder for marks,
                    // set selection might failed if the array is null.
                    int[] marks = Enumerable.Repeat(0, selCount).ToArray();

                    // set selection
                    swFeatData.SetSelections(dw, marks);

                    // finalize modification
                    swFeat.ModifyDefinition(swFeatData, swModel, null);
                }
            }
            else
            {
                if(State == 1)
                {
                    // release feature data if cancel editing so that feature can roll forward
                    swFeatData.ReleaseSelectionAccess();
                }
            }
        }

        public void OnSelectionboxListChanged(int Id, int Count)
        {
            /*if(Id == SelectionboxId)
            {
                // if selection is less than the counter, means that some item has been removed from the list
                if(Count < preSelCount)
                {
                    MessageBox.Show("less than preselect");

                    // get material properties colour
                    double[] featColour = (double[])swModel.MaterialPropertyValues;

                    // instantiate selection manager
                    SelectionMgr swSelMgr = swModel.ISelectionManager;

                    // get total number of selection
                    int selCount = swSelMgr.GetSelectedObjectCount2(mark);
                    for(int i = 1;i < selCount + 1; i++)
                    {
                        // Cast each selected object as Feature
                        object selObj = swSelMgr.GetSelectedObject6(i, mark);
                        Feature swFeat = (Feature)selObj;

                        // add the feature to post selection list
                        postSelFeats.Add(swFeat);
                    }

                    // compare pre and post
                    List<Feature> diffList = preSelFeats.Except(postSelFeats).ToList();
                    
                    // For each feature that is removed
                    foreach (Feature deSelFeat in diffList)
                    {
                        // get all faces from the feature
                        object[] vFaces = (object[])deSelFeat.GetFaces();
                        foreach (object vFace in vFaces)
                        {
                            // cast object as face2
                            Face2 swFace = (Face2)vFace;

                            // set colour of this deleted face
                            swFace.MaterialPropertyValues = featColour;

                            // get all edges from this face
                            object[] vEdges = (object[])swFace.GetEdges();
                            foreach (object vEdge in vEdges)
                            {
                                // cast object as Edge
                                Edge swEdge = (Edge)vEdge;

                                // cast edge as entity
                                Entity swEnt = (Entity)swEdge;

                                // delete the entity name
                                ((PartDoc)swModel).DeleteEntityName(swEnt);
                            }
                        }
                    }

                    // recreate the pre selection list with the post selection list
                    preSelFeats = new List<Feature>(postSelFeats);
                }

                // set pre selection count
                preSelCount = Count;
            }*/
        }

        public bool OnSubmitSelection(int Id, object Selection, int SelType, ref string ItemText)
        {
            if(SelType == (int)swSelectType_e.swSelFACES)
            {
                SelectFeatureFromFace(Selection);
                return false;
            }
            else
            {
                return true;
            }
        }

        // Method to select feature from selected face
        private void SelectFeatureFromFace(object selection)
        {
            // cast selection to face
            Face swFace = (Face)selection;

            // get feature associated to the face
            Feature swFeat = swFace.IGetFeature();

            // select the feature
            swFeat.Select2(true, 0);
        }

        #region Unused Implementation
        public void AfterActivation()
        {
        }

        public void AfterClose()
        {
        }

        public bool OnHelp()
        {
            return true;
        }

        public bool OnPreviousPage()
        {
            return true;
        }

        public bool OnNextPage()
        {
            return true;
        }

        public bool OnPreview()
        {
            return true;
        }

        public void OnWhatsNew()
        {
        }

        public void OnUndo()
        {
        }

        public void OnRedo()
        {
        }

        public bool OnTabClicked(int Id)
        {
            return true;
        }

        public void OnGroupExpand(int Id, bool Expanded)
        {
        }

        public void OnGroupCheck(int Id, bool Checked)
        {
        }

        public void OnCheckboxCheck(int Id, bool Checked)
        {
        }

        public void OnOptionCheck(int Id)
        {
        }

        public void OnButtonPress(int Id)
        {
        }

        public void OnTextboxChanged(int Id, string Text)
        {
        }

        public void OnNumberboxChanged(int Id, double Value)
        {
        }

        public void OnComboboxEditChanged(int Id, string Text)
        {
        }

        public void OnComboboxSelectionChanged(int Id, int Item)
        {
        }

        public void OnListboxSelectionChanged(int Id, int Item)
        {
        }

        public void OnSelectionboxFocusChanged(int Id)
        {
        }

        public void OnSelectionboxCalloutCreated(int Id)
        {
        }

        public void OnSelectionboxCalloutDestroyed(int Id)
        {
        }

        public int OnActiveXControlCreated(int Id, bool Status)
        {
            return -1;
        }

        public void OnSliderPositionChanged(int Id, double Value)
        {
        }

        public void OnSliderTrackingCompleted(int Id, double Value)
        {
        }

        public bool OnKeystroke(int Wparam, int Message, int Lparam, int Id)
        {
            return true;
        }

        public void OnPopupMenuItem(int Id)
        {
        }

        public void OnPopupMenuItemUpdate(int Id, ref int retval)
        {
        }

        public void OnGainedFocus(int Id)
        {
        }

        public void OnLostFocus(int Id)
        {
        }

        public int OnWindowFromHandleControlCreated(int Id, bool Status)
        {
            return -1;
        }

        public void OnListboxRMBUp(int Id, int PosX, int PosY)
        {
        }

        public void OnNumberBoxTrackingCompleted(int Id, double Value)
        {
        }
        #endregion
    }
}
