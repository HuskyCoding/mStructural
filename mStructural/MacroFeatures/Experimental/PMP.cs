using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;

namespace mStructural.MacroFeatures
{
    [ComVisible(true)]
    public class PMP : PropertyManagerPage2Handler9
    {
        #region Private variables
        private PropertyManagerPage2 pmp;
        private PropertyManagerPageGroup pmp_group;
        private PropertyManagerPageSelectionbox pmp_sb;

        private const int GroupID = 1;
        private const int SelectionboxID = 2;

        private SldWorks swApp;
        private ModelDoc2 swModel;
        private Feature mFeat;
        private MacroFeatureData mFeatData;
        private int State;
        #endregion

        // constructor
        public PMP(SldWorks swapp, int state, Feature feat)
        {
            swApp = swapp;
            swModel = swApp.IActiveDoc2;
            State = state; // 0 is new, 1 is edit
            mFeat = feat;

            string pageTitle = "";
            string caption = "";
            string tip = null;
            int options = 0;
            int err = 0;
            int controlType = 0;
            int alignment = 0;

            // set page
            pageTitle = "My PMP";
            options = (int)swPropertyManagerButtonTypes_e.swPropertyManager_OkayButton +
                (int)swPropertyManagerButtonTypes_e.swPropertyManager_CancelButton +
                (int)swPropertyManagerPageOptions_e.swPropertyManagerOptions_LockedPage +
                (int)swPropertyManagerPageOptions_e.swPropertyManagerOptions_PushpinButton;
            pmp = (PropertyManagerPage2)swapp.CreatePropertyManagerPage(pageTitle,options,this,ref err);
            
            // make sure page is created successfully
            if(err == (int)swPropertyManagerPageStatus_e.swPropertyManagerPage_Okay)
            {
                // add group
                caption = "My Group";
                options = (int)swAddGroupBoxOptions_e.swGroupBoxOptions_Visible +
                    (int)swAddGroupBoxOptions_e.swGroupBoxOptions_Expanded;
                pmp_group = (PropertyManagerPageGroup)pmp.AddGroupBox(GroupID,caption,options);

                // add selection box
                controlType = (int)swPropertyManagerPageControlType_e.swControlType_Selectionbox;
                caption = "";
                alignment = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_Indent;
                options = (int)swAddControlOptions_e.swControlOptions_Visible +
                    (int)swAddControlOptions_e.swControlOptions_Enabled;
                tip = "Select sketches";
                pmp_sb = (PropertyManagerPageSelectionbox)pmp_group.AddControl(SelectionboxID, (short)controlType, caption, (short)alignment, options, tip);
                swSelectType_e[] filters = new swSelectType_e[1];
                filters[0] = swSelectType_e.swSelSKETCHES;
                object filterObj = null;
                filterObj = filters;
                pmp_sb.SingleEntityOnly = false;
                pmp_sb.AllowMultipleSelectOfSameEntity = false;
                pmp_sb.Height = 50;
                pmp_sb.SetSelectionFilters(filters);

                // set selection
                if(state == 1)
                {
                    mFeatData = (MacroFeatureData)mFeat.GetDefinition();
                    mFeatData.AccessSelections(swModel, null);
                    object sels;
                    object types;
                    object marks;
                    mFeatData.GetSelections(out sels, out types, out marks);
                    if (sels != null)
                    {
                        object[] selArr = (object[])sels;
                        foreach (object sel in selArr)
                        {
                            Feature swFeat = (Feature)sel;
                            swFeat.Select2(true, 0);
                        }
                    }
                }

            }
            else
            {
                MessageBox.Show("Failed to create page");
            }
        }

        // Method to show PMP
        public void Show()
        {
            pmp.Show();
        }

        public void AfterActivation()
        {
            throw new System.NotImplementedException();
        }

        public void OnClose(int Reason)
        {
            if(Reason == (int)swPropertyManagerPageCloseReasons_e.swPropertyManagerPageClose_Okay)
            {
                if(State == 0) // new
                {
                    FeatureManager featMgr = swApp.IActiveDoc2.FeatureManager;
                    Feature mFeat = featMgr.InsertMacroFeature3("MacroFeature", "mStructural.MacroFeatures.MacroFeature", null, null, null, null, null, null, null, null, (int)swMacroFeatureOptions_e.swMacroFeatureByDefault);
                    MacroFeatureData mFeatData = (MacroFeatureData)mFeat.GetDefinition();

                    // set selected object
                    SelectionMgr swSelMgr = swModel.ISelectionManager;
                    int mark = -1;
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
                else if(State == 1) // editing
                {
                    // set selection for feature
                    SelectionMgr swSelMgr = swModel.ISelectionManager;
                    int mark = -1;
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
                    bool bREt = mFeat.ModifyDefinition(mFeatData, swModel, null);
                }
                else
                {
                    MessageBox.Show("Invalid state");
                }
            }
            else
            {
                mFeatData.ReleaseSelectionAccess();
            }
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

        public void OnCheckboxCheck(int Id, bool Checked)
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

        public void OnNumberboxChanged(int Id, double Value)
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

        public void OnSelectionboxListChanged(int Id, int Count)
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
    }
}
