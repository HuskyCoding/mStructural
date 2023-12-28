using mStructural.Function;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;

namespace mStructural.MacroFeatures
{
    [ComVisible(true)]
    public class ExperimentalPMP : PropertyManagerPage2Handler9
    {
        #region Private Variables
        private SldWorks swApp;
        private ModelDoc2 swModel;
        private Feature mFeat;
        private MacroFeatureData mFeatData;
        private PropertyManagerPage2 pmp;
        private PropertyManagerPageGroup pmp_group1;
        private PropertyManagerPageGroup pmp_group2;
        private PropertyManagerPageGroup pmp_group3;
        private PropertyManagerPageSelectionbox pmp_sb;
        private PropertyManagerPageNumberbox pmp_nb;
        private PropertyManagerPageCheckbox pmp_cb;
        private const int GroupID1 = 1;
        private const int GroupID2 = 2;
        private const int GroupID3 = 3;
        private const int SelectionID = 4;
        private const int NumberBoxID = 5;
        private const int CheckBoxID = 6;
        private Message msg;
        private Experimental etchFeat;
        private List<Body2> bodyList;
        private double depth;
        private int mark = 1;
        private int cutDirection;
        private int State; // 0 = new, 1 = edit
        #endregion

        // constructor
        public ExperimentalPMP(SldWorks swapp, int state, Feature feat)
        {
            // initialize
            swApp = swapp; // solidworks
            swModel = swApp.IActiveDoc2;
            msg = new Message(swApp); // message
            etchFeat = new Experimental(swApp);
            mFeat = feat;
            bodyList = new List<Body2>();

            swModel.ClearSelection2(true);

            State = state;

            // create property manager page error
            int errs =0;

            // options to create pmp
            int options = (int)swPropertyManagerPageOptions_e.swPropertyManagerOptions_OkayButton +
                (int)swPropertyManagerPageOptions_e.swPropertyManagerOptions_CancelButton;
            
            // create pmp
            pmp = swApp.ICreatePropertyManagerPage("Etch Feature", options, this, ref errs);

            if(errs == (int)swPropertyManagerPageStatus_e.swPropertyManagerPage_Okay)
            {
                string caption = "";
                int controlType = 0;
                int alignment = 0;
                string tip = "";
                swSelectType_e[] filters = null;
                object filterObj = null;

                #region Sketches
                // create group
                caption = "Sketches";
                options = (int)swAddGroupBoxOptions_e.swGroupBoxOptions_Visible + (int)swAddGroupBoxOptions_e.swGroupBoxOptions_Expanded;
                pmp_group1 = pmp.IAddGroupBox(GroupID1, caption, options);

                // create selection box for feature
                controlType = (int)swPropertyManagerPageControlType_e.swControlType_Selectionbox;
                alignment = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_Indent;
                options = (int)swAddControlOptions_e.swControlOptions_Visible +
                    (int)swAddControlOptions_e.swControlOptions_Enabled;
                tip = "Select sketch for etch cut.";
                pmp_sb = (PropertyManagerPageSelectionbox)pmp_group1.IAddControl(SelectionID, (short)controlType, caption, (short)alignment, options, tip);
                filters = new swSelectType_e[1];
                filters[0] = swSelectType_e.swSelSKETCHES;
                filterObj = filters;
                pmp_sb.Height = 50;
                pmp_sb.SingleEntityOnly = false;
                pmp_sb.AllowMultipleSelectOfSameEntity = false;
                pmp_sb.SetSelectionFilters(filterObj);
                pmp_sb.Mark = mark;
                depth = 0.0001;
                #endregion

                #region Depth
                // create group
                caption = "Settings";
                options = (int)swAddGroupBoxOptions_e.swGroupBoxOptions_Visible + (int)swAddGroupBoxOptions_e.swGroupBoxOptions_Expanded;
                pmp_group2 = pmp.IAddGroupBox(GroupID2, caption, options);

                // create number box for feature
                controlType = (int)swPropertyManagerPageControlType_e.swControlType_Numberbox;
                alignment = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_Indent;
                options = (int)swAddControlOptions_e.swControlOptions_Visible +
                    (int)swAddControlOptions_e.swControlOptions_Enabled;
                tip = "Etch cut depth.";
                pmp_nb = (PropertyManagerPageNumberbox)pmp_group2.IAddControl(NumberBoxID, (short)controlType, caption, (short)alignment, options, tip);
                pmp_nb.Height = 60;
                pmp_nb.Value = depth;
                pmp_nb.SetRange2((int)swNumberboxUnitType_e.swNumberBox_Length, 0.0001, 0.1, true, 0.0001, 0.0001, 0.0001);
                #endregion

                #region Flip Direction
                // create check box for flip direction
                caption = "Flip Direction";
                controlType = (int)swPropertyManagerPageControlType_e.swControlType_Checkbox;
                alignment = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_Indent;
                options = (int)swAddControlOptions_e.swControlOptions_Visible +
                    (int)swAddControlOptions_e.swControlOptions_Enabled;
                tip = "Define the direction of cut extrusion";
                pmp_cb = (PropertyManagerPageCheckbox)pmp_group2.IAddControl(CheckBoxID, (short)controlType, caption, (short)alignment, options, tip);
                pmp_cb.Checked = false;
                cutDirection = -1;
                #endregion

                if(State == 1)
                {
                    // populate sketches
                    mFeatData = (MacroFeatureData)mFeat.GetDefinition();
                    mFeatData.AccessSelections(swModel, null);
                    object sels;
                    object types;
                    object marks;
                    object views;
                    object xforms;
                    mFeatData.GetSelections3(out sels, out types, out marks, out views, out xforms);
                    if (sels != null)
                    {
                        object[] selArr = (object[])sels;
                        foreach (object sel in selArr)
                        {
                            Feature swFeat = (Feature)sel;
                            swFeat.Select2(true, mark);
                        }
                    }

                    // update variables
                    object pNames;
                    object pTypes;
                    object pValues;
                    mFeatData.GetParameters(out pNames, out pTypes, out pValues);
                    object[] valueArr = (object[])pValues;
                    depth = double.Parse(valueArr[0].ToString());
                    pmp_nb.Value= depth;
                    cutDirection = int.Parse(valueArr[1].ToString());
                    if(cutDirection == -1)
                    {
                        pmp_cb.Checked = false;
                    }
                    else if(cutDirection == 1) 
                    {
                        pmp_cb.Checked = true;
                    }

                    UpdatePreview(mark);
                }
            }
            else
            {
                msg.ErrorMsg("Error initializing the function");
                return;
            }
        }

        // entry
        public void Show()
        {
            pmp.Show2(0);
        }

        // Method to update Preview
        private void UpdatePreview(int mark)
        {
            ModelDoc2 swModel = swApp.IActiveDoc2;

            HidePreview(swModel);

            bodyList = etchFeat.UpdatePreview(swModel, depth, mark, cutDirection);

            // display bodies
            if (bodyList.Count > 0)
            {
                foreach (Body2 swBody in bodyList)
                {
                    swBody.Display3(swModel, 255, (int)swTempBodySelectOptions_e.swTempBodySelectOptionNone);
                    double[] props = new double[9];
                    props[0] = 1;
                    props[1] = 1;
                    props[2] = 0;
                    props[3] = 0;
                    props[4] = 0;
                    props[5] = 0;
                    props[6] = 1;
                    props[7] = 0.5;
                    props[8] = 1;
                    swBody.MaterialPropertyValues2 = props;
                }
            }
        }

        private void HidePreview(ModelDoc2 swModel)
        {
            // hide temp bodies
            if (bodyList.Count > 0)
            {
                for (int i = 0; i < bodyList.Count; i++)
                {
                    bodyList[i].Hide(swModel);
                    bodyList[i] = null;
                }
            }
        }

        private void SetSelectionData()
        {
            SelectionMgr swSelMgr = swModel.ISelectionManager;
            int selCount = swSelMgr.GetSelectedObjectCount2(mark);
            object[] selObj = new object[selCount];
            DispatchWrapper[] dw = new DispatchWrapper[selCount];
            for (int i = 1; i < selCount + 1; i++)
            {
                selObj[i - 1] = swSelMgr.GetSelectedObject6(i, mark);
                dw[i - 1] = new DispatchWrapper(selObj[i - 1]);
            }
            int[] marks = Enumerable.Repeat(0, selCount).ToArray();
            mFeatData.SetSelections(dw, marks);
        }

        private void SetParameterData()
        {
            // parameter
            string[] paramNames = new string[2];
            int[] paramTypes = new int[2];
            string[] paramValues = new string[2];

            paramNames[0] = "Depth";
            paramNames[1] = "CutDirection";
            paramTypes[0] = (int)swMacroFeatureParamType_e.swMacroFeatureParamTypeDouble;
            paramTypes[1] = (int)swMacroFeatureParamType_e.swMacroFeatureParamTypeInteger;
            paramValues[0] = depth.ToString();
            paramValues[1] = cutDirection.ToString();
            mFeatData.SetParameters(paramNames, paramTypes, paramValues);
        }

        public void OnClose(int Reason)
        {
            // hide preview
            HidePreview(swModel);
            if (Reason == (int)swPropertyManagerPageCloseReasons_e.swPropertyManagerPageClose_Cancel)
            {
                if(State == 1)
                {
                    mFeatData.ReleaseSelectionAccess();
                }
            }
            else if(Reason == (int)swPropertyManagerPageCloseReasons_e.swPropertyManagerPageClose_Okay)
            {
                if (State == 0) // new feature
                {
                    // collect all bodies
                    object[] vBodies = null;
                    List<Body2> editBodyList = new List<Body2>();
                    vBodies = (object[])((PartDoc)swModel).GetBodies2((int)swBodyType_e.swAllBodies, false);
                    if(vBodies.Length > 0)
                    {
                        foreach(object vBody in vBodies)
                        {
                            Body2 editBody = (Body2)vBody;
                            editBodyList.Add(editBody);
                        }
                    }

                    // parameter
                    string[] paramNames = new string[2];
                    int[] paramTypes = new int[2];
                    string[] paramValues = new string[2];

                    paramNames[0] = "Depth";
                    paramNames[1] = "CutDirection";
                    paramTypes[0] = (int)swMacroFeatureParamType_e.swMacroFeatureParamTypeDouble;
                    paramTypes[1] = (int)swMacroFeatureParamType_e.swMacroFeatureParamTypeInteger;
                    paramValues[0] = depth.ToString();
                    paramValues[1] = cutDirection.ToString();

                    // icon
                    string[] iconFiles = new string[3];
                    string assemblyFolder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                    iconFiles[0] = Path.Combine(assemblyFolder, "Macro Feature Icon\\EtchCut.bmp");
                    iconFiles[1] = Path.Combine(assemblyFolder, "Macro Feature Icon\\EtchCut_S.bmp");
                    iconFiles[2] = Path.Combine(assemblyFolder, "Macro Feature Icon\\EtchCut.bmp");

                    // create macro feature
                    FeatureManager swFeatMgr = swModel.FeatureManager;
                    mFeat = swFeatMgr.InsertMacroFeature3("EtchFeature", "mStructural.MacroFeatures.EtchFeatureMF", null,
                        paramNames, paramTypes, paramValues,
                        null, null, editBodyList.ToArray(), iconFiles, (int)swMacroFeatureOptions_e.swMacroFeatureByDefault);

                    //mFeatData = (MacroFeatureData)mFeat.GetDefinition();
                    //SetSelectionData();
                }
                else if(State == 1) // edit
                {
                    SetParameterData();
                    SetSelectionData();
                    bool bRet = mFeat.ModifyDefinition(mFeatData, swModel, null);
                    MessageBox.Show(bRet.ToString());
                }
            }
        }

        public void OnSelectionboxListChanged(int Id, int Count)
        {
            if(Id == SelectionID)
            {
                ModelDoc2 swModel = swApp.IActiveDoc2;
                SelectionMgr swSelMgr = swModel.ISelectionManager;
                int selCount = swSelMgr.GetSelectedObjectCount2(mark);
                for(int i = 1; i < selCount + 1; i++)
                {
                    object selObj = swSelMgr.GetSelectedObject6(i, mark);
                }
                UpdatePreview(mark);
            }
        }

        public void OnNumberboxChanged(int Id, double Value)
        {
            if(Id == NumberBoxID)
            {
                depth = Value;
                UpdatePreview(mark);
            }
        }

        public void OnCheckboxCheck(int Id, bool Checked)
        {
            if(Id == CheckBoxID)
            {
                if (Checked)
                {
                    cutDirection = 1;
                }
                else
                {
                    cutDirection = -1;
                }
                UpdatePreview(mark);
            }
        }

        #region Unused method
        public void AfterActivation()
        {
            throw new System.NotImplementedException();
        }

        public void AfterClose()
        {
            throw new System.NotImplementedException();
        }

        public bool OnHelp()
        {
            throw new System.NotImplementedException();
        }

        public bool OnPreviousPage()
        {
            throw new System.NotImplementedException();
        }

        public bool OnNextPage()
        {
            throw new System.NotImplementedException();
        }

        public bool OnPreview()
        {
            throw new System.NotImplementedException();
        }

        public void OnWhatsNew()
        {
            throw new System.NotImplementedException();
        }

        public void OnUndo()
        {
            throw new System.NotImplementedException();
        }

        public void OnRedo()
        {
            throw new System.NotImplementedException();
        }

        public bool OnTabClicked(int Id)
        {
            throw new System.NotImplementedException();
        }

        public void OnGroupExpand(int Id, bool Expanded)
        {
            throw new System.NotImplementedException();
        }

        public void OnGroupCheck(int Id, bool Checked)
        {
            throw new System.NotImplementedException();
        }

        public void OnOptionCheck(int Id)
        {
            throw new System.NotImplementedException();
        }

        public void OnButtonPress(int Id)
        {
            throw new System.NotImplementedException();
        }

        public void OnTextboxChanged(int Id, string Text)
        {
            throw new System.NotImplementedException();
        }

        public void OnComboboxEditChanged(int Id, string Text)
        {
            throw new System.NotImplementedException();
        }

        public void OnComboboxSelectionChanged(int Id, int Item)
        {
            throw new System.NotImplementedException();
        }

        public void OnListboxSelectionChanged(int Id, int Item)
        {
            throw new System.NotImplementedException();
        }

        public void OnSelectionboxFocusChanged(int Id)
        {
            throw new System.NotImplementedException();
        }

        public void OnSelectionboxCalloutCreated(int Id)
        {
            throw new System.NotImplementedException();
        }

        public void OnSelectionboxCalloutDestroyed(int Id)
        {
            throw new System.NotImplementedException();
        }

        public bool OnSubmitSelection(int Id, object Selection, int SelType, ref string ItemText)
        {
            return true;
        }

        public int OnActiveXControlCreated(int Id, bool Status)
        {
            throw new System.NotImplementedException();
        }

        public void OnSliderPositionChanged(int Id, double Value)
        {
            throw new System.NotImplementedException();
        }

        public void OnSliderTrackingCompleted(int Id, double Value)
        {
            throw new System.NotImplementedException();
        }

        public bool OnKeystroke(int Wparam, int Message, int Lparam, int Id)
        {
            throw new System.NotImplementedException();
        }

        public void OnPopupMenuItem(int Id)
        {
            throw new System.NotImplementedException();
        }

        public void OnPopupMenuItemUpdate(int Id, ref int retval)
        {
            throw new System.NotImplementedException();
        }

        public void OnGainedFocus(int Id)
        {
            throw new System.NotImplementedException();
        }

        public void OnLostFocus(int Id)
        {
            throw new System.NotImplementedException();
        }

        public int OnWindowFromHandleControlCreated(int Id, bool Status)
        {
            throw new System.NotImplementedException();
        }

        public void OnListboxRMBUp(int Id, int PosX, int PosY)
        {
            throw new System.NotImplementedException();
        }

        public void OnNumberBoxTrackingCompleted(int Id, double Value)
        {
            throw new System.NotImplementedException();
        }

        #endregion
    }
}
