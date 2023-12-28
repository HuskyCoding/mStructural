using mStructural.Classes;
using mStructural.Function;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;

namespace mStructural.WPF
{
    /// <summary>
    /// Interaction logic for CreateBodyViewWPF.xaml
    /// </summary>
    public partial class CreateBodyViewWPF : Window
    {
        #region Public variables
        public bool err = false;
        public static CreateBodyViewWPF Current { get; private set; }
        #endregion

        #region Private variables
        private SldWorks swApp;
        private Message msg;
        private ModelDoc2 curModelDoc;
        private View parentView;
        private View topView;
        private View backView;
        private View btmView;
        private View rightView;
        private View leftView;
        private string curSheetName;
        private Note parentNote;
        private List<Note> swNotes;
        private ObservableCollection<CutlistItem> currentCl;
        #endregion
        
        // constructor
        public CreateBodyViewWPF(SWIntegration swintegration)
        {
            InitializeComponent();
            swApp = swintegration.SwApp;
            msg = new Message(swApp);
            ViewControlPanel.Visibility = Visibility.Hidden;

            // get all weldment cutlist
            bool bRet = PopulateCutlist();
            if (!bRet)
            {
                return;
            }

            // populate sheet name
            PopulateSheetNames();

            // populate available scale
            PopulateScales();
        }

        // Method to generate data
        private ObservableCollection<CutlistItem> GetClItem(int startSearchIndex, int endSearchIndex)
        {
            ObservableCollection<CutlistItem> clitem = new ObservableCollection<CutlistItem>();
            bool bRet = false;
            mainViewCb.Items.Clear();

            // get active doc
            ModelDoc2 swModel = swApp.IActiveDoc2;
            curModelDoc = swModel;
            DrawingDoc swDraw = (DrawingDoc)swModel;
            TableAnnotation swTableAnn = default(TableAnnotation);
            WeldmentCutListAnnotation wclTableAnn;
            View swView = default(View);
            int iterationStatus = 0;
            int currentSheetIndex = 0;

            // get the weldment table
            string[] sheetNames = (string[])swDraw.GetSheetNames();
            for(int i =0;i<sheetNames.Length;i++)
            {
                string sheetName = sheetNames[i];
                swDraw.ActivateSheet(sheetName);
                swView = swDraw.IGetFirstView();
                TableAnnotation tempTableAnn = swView.GetFirstTableAnnotation();
                while (tempTableAnn != null)
                {
                    if(tempTableAnn.Type == (int)swTableAnnotationType_e.swTableAnnotation_WeldmentCutList)
                    {
                        wclTableAnn = (WeldmentCutListAnnotation)tempTableAnn;
                        if(wclTableAnn.WeldmentCutListFeature.GetFeature().Name == wclCb.SelectedValue.ToString())
                        {
                            curSheetName = sheetName;
                            swTableAnn = tempTableAnn;
                            currentSheetIndex = i;
                            iterationStatus += 1;
                            break;
                        }
                    }
                    tempTableAnn = tempTableAnn.GetNext();
                }
                if (iterationStatus == 1)
                    break;
            }

            // add column for cutlist name
            bRet = swTableAnn.InsertColumn2((int)swTableItemInsertPosition_e.swTableItemInsertPosition_Last, 0, "Name", (int)swTableColumnTypes_e.swWeldTableColumnType_CutListName);
            bRet = swTableAnn.SetColumnType(swTableAnn.ColumnCount - 1, (int)swTableColumnTypes_e.swWeldTableColumnType_CutListName);

            // process each row
            for(int i = 1; i < swTableAnn.RowCount; i++)
            {
                CutlistItem cli = new CutlistItem();
                cli.ItemNo = Convert.ToInt32(swTableAnn.DisplayedText2[i, 0, false]);
                cli.Name = swTableAnn.DisplayedText2[i, swTableAnn.ColumnCount - 1, false];
                cli.IsCreated = isBallooned(swDraw, sheetNames, startSearchIndex, endSearchIndex, cli.ItemNo);
                clitem.Add(cli);
            }

            // delete name column after done
            bRet = swTableAnn.DeleteColumn2(swTableAnn.ColumnCount - 1, false);
            if (!bRet)
            {
                MessageBox.Show("Name column deletion failed. Please delete it manually");
            }

            // go to the table sheet
            swDraw.ActivateSheet(sheetNames[currentSheetIndex]);

            // populate views in main view combobox
            swView = swDraw.IGetFirstView();
            swView = swView.IGetNextView();
            string firstViewName = swView.GetName2();
            while (swView != null)
            {
                mainViewCb.Items.Add(swView.GetName2());
                swView = swView.IGetNextView();
            }
            mainViewCb.SelectedValue = firstViewName;

            return clitem;
        }

        // Method to populate all cutlist to the combobox
        private bool PopulateCutlist()
        {
            ModelDoc2 swModel = swApp.IActiveDoc2;
            
            // check if any document is open
            if (swModel == null)
            {
                MessageBox.Show(msg.ModelDocNull);
                return false;
            }

            // check if opened doc is drawing
            if (swModel.GetType() != (int)swDocumentTypes_e.swDocDRAWING)
            {
                MessageBox.Show(msg.NotDrawingDoc);
                return false;
            }

            Feature swFeat = swModel.IFirstFeature();
            while (swFeat != null)
            {
                if(swFeat.GetTypeName2() == "WeldmentTableFeat")
                {
                    wclCb.Items.Add(swFeat.Name);
                }
                swFeat = swFeat.IGetNextFeature();
            }

            return true;
        }

        // Method to populate all sheets
        private void PopulateSheetNames()
        {
            ModelDoc2 swModel = swApp.IActiveDoc2;
            DrawingDoc swDraw = (DrawingDoc)swModel;

            // get the weldment table
            string[] sheetNames = (string[])swDraw.GetSheetNames();
            foreach (string sheetName in sheetNames)
            {
                searchStartCb.Items.Add(sheetName);
                searchEndCb.Items.Add(sheetName);
            }
        }

        // Method to populate predefined scales
        private void PopulateScales()
        {
            // get user preferences
            string settingFileLoc = swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swFileLocationsDrawingScaleStandard);
            settingFileLoc = settingFileLoc + "/" + "drawingscales.txt";
            foreach(string line in
            File.ReadLines(settingFileLoc)
                .SkipWhile(line => !line.Contains("#ISO"))
                .Skip(1)
                .TakeWhile(line => !line.Contains(";;")))
            {
                scaleCb.Items.Add(line);
            }
        }

        // Method to check if balloon added in drawing
        private bool isBallooned(DrawingDoc swdraw, string[] sheetnames, int startIndex, int endIndex, int balloontxt)
        {
            bool bRet = false;
            string sheetName = "";

            for (int i = startIndex; i<endIndex+1; i++)
            {
                sheetName = sheetnames[i];
                swdraw.ActivateSheet(sheetName);
                View swView = swdraw.IGetFirstView();
                while (swView != null)
                {
                    Note swNote = swView.IGetFirstNote();
                    while(swNote != null)
                    {
                        if (swNote.IsBomBalloon())
                        {
                            if(swNote.GetBomBalloonText(true) == balloontxt.ToString())
                            {
                                return true;
                            }
                        }
                        swNote = swNote.IGetNext();
                    }
                    swView = swView.IGetNextView();
                }
            }

            return bRet;
        }

        // Method to select view body
        private bool selectViewBody(ModelDoc2 swmodel, View swview, string clname)
        {
            bool bRet = false;
            Feature clFeat = default(Feature);
            BodyFolder swBodyFolder = default(BodyFolder);
            SelectionMgr swSelMgr = default(SelectionMgr);
            object[] arrBody = null;
            object[] bodies = new object[1];
            DispatchWrapper[] arrBodiesIn = new DispatchWrapper[1];

            swSelMgr = swmodel.ISelectionManager;

            // select the feature with cutlist name column
            bRet = swmodel.Extension.SelectByID2(clname, "BDYFOLDER", 0, 0, 0, false, 0, null, 0);

            // if failed to select
            if (!bRet)
            {
                return false;
            }

            // cast selected object to feature object
            clFeat = (Feature)swSelMgr.GetSelectedObject6(1, -1);

            // get body folder from feature
            swBodyFolder = (BodyFolder)clFeat.GetSpecificFeature2();

            // get all body from body folder
            arrBody = (object[])swBodyFolder.GetBodies();

            // use the first body in the list
            bodies[0] = arrBody[0];

            // create new dispatch wrapper for view body selection
            arrBodiesIn[0] = new DispatchWrapper(bodies[0]);

            // select the body for the view
            swview.Bodies = (arrBodiesIn);

            // clear selection after done
            swmodel.ClearSelection2(true);
            swSelMgr = null;

            return bRet;
        }

        // override windows load event
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            if (Current != null)
            {
                Close();
                if(!Current.IsActive)
                    Current.Activate();
                if(!Current.IsFocused)
                    Current.Focus();
                MessageBox.Show("Another window is already opened!");
            }
            else
            {
                Current = this;
            }
        }

        // override windows close event
        private void Window_Closed(object sender, EventArgs e)
        {
            if ((Current == this) && (Current != null))
                Current = null;
        }

        // event handler for wclCb selection change
        private void wclCb_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            DG1.ItemsSource = GetClItem(0,0);
            ViewControlPanel.Visibility = Visibility.Hidden;
        }

        // event handler to set auto generated column props
        private void DG1_AutoGeneratingColumn(object sender, System.Windows.Controls.DataGridAutoGeneratingColumnEventArgs e)
        {
            e.Column.IsReadOnly = true;
        }

        // handle search button click
        private void searchBtn_Click(object sender, RoutedEventArgs e)
        {
            // check if combobox has any item selected
            if(searchStartCb.SelectedIndex == -1)
            {
                MessageBox.Show("Nothing is selected a start range");
                return;
            }

            if (searchEndCb.SelectedIndex == -1)
            {
                MessageBox.Show("Nothing is selected a end range");
                return;
            }

            if (searchEndCb.SelectedIndex < searchStartCb.SelectedIndex)
            {
                MessageBox.Show("Invalid range, end range cannot be less than start range");
                return;
            }

            DG1.ItemsSource = GetClItem(searchStartCb.SelectedIndex, searchEndCb.SelectedIndex);
        }

        // handle scaleCb changed
        private void scaleCb_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            try
            {
                double[] dScale = new double[2];
                string selectedScale = scaleCb.SelectedValue.ToString();
                dScale[0] = Convert.ToDouble(selectedScale.Substring(0, selectedScale.IndexOf(":")));
                dScale[1] = Convert.ToDouble(selectedScale.Substring(selectedScale.IndexOf(":")+1));
                parentView.ScaleDecimal = dScale[0] / dScale[1];

                // change note text
                DrawingDoc swDraw = (DrawingDoc)curModelDoc;
                Sheet swSheet = swDraw.IGetCurrentSheet();
                double[] sheetProps = (double[])swSheet.GetProperties2();

                if (dScale[0] == sheetProps[2] && dScale[1] == sheetProps[3])
                {
                    parentNote.SetText("$PRPWLD:\"DESCRIPTION\"");
                }
                else
                {
                    parentNote.SetText("$PRPWLD:\"DESCRIPTION\"\r\nSCALE $PRP:\"SW-View Scale(View Scale)\"");
                }

                updateNotePos(parentView, parentNote);

                curModelDoc.GraphicsRedraw2();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        // View Creation
        #region View Creation
        private void DG1_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            // some event handler need to be remove first
            resetControl();
            scaleCb.SelectionChanged -= scaleCb_SelectionChanged;

            // Create a view for the selected body
            DrawingDoc swDraw = (DrawingDoc)curModelDoc;
            swDraw.ActivateSheet(curSheetName);

            // get sheet size
            Sheet swSheet = swDraw.IGetCurrentSheet();
            double[] sheetProps = (double[])swSheet.GetProperties2();

            // main model reference
            SelectionMgr swSelMgr = curModelDoc.ISelectionManager;
            bool bRet = curModelDoc.Extension.SelectByID2(mainViewCb.SelectedValue.ToString(), "DRAWINGVIEW", 0, 0, 0, false, 1, null, 0);
            View swView = (View)swSelMgr.GetSelectedObject6(1,1);
            ModelDoc2 mainModel = swView.ReferencedDocument;

            // create parent view
            parentView = swDraw.CreateDrawViewFromModelView3(mainModel.GetPathName(), "*Front", sheetProps[5] + 0.05, sheetProps[6] - 0.05, 0);
            parentView.SetDisplayTangentEdges2((int)swDisplayTangentEdges_e.swTangentEdgesVisibleAndFonted);

            // Select current scale
            double[] scaleRatio = (double[])parentView.ScaleRatio;
            string scaleStr = scaleRatio[0].ToString() + ":" + scaleRatio[1].ToString();
            scaleCb.SelectedValue = scaleStr;

            // check if the auto add label note enabled
            bool upAddLabel = swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDetailingOrthoView_AddViewLabelOnViewCreation);

            // set body for this view
            CutlistItem selectedRow = (CutlistItem)DG1.SelectedItem;
            bRet = selectViewBody(mainModel, parentView, selectedRow.Name);

            // add note according to toggle setting
            string noteString = "";
            if (upAddLabel)
            {
                noteString = "$PRPWLD:\"DESCRIPTION\"";
            }
            else
            {
                if (scaleRatio[0] == sheetProps[2] && scaleRatio[1] == sheetProps[3])
                {
                    noteString = "$PRPWLD:\"DESCRIPTION\"";
                }
                else
                {
                    noteString = "$PRPWLD:\"DESCRIPTION\"\r\nSCALE $PRP:\"SW-View Scale(View Scale)\"";
                }
            }
            bRet = curModelDoc.Extension.SelectByID2(parentView.Name, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
            parentNote = curModelDoc.IInsertNote(noteString);
            parentNote.SetTextJustification((int)swTextJustification_e.swTextJustificationCenter);
            updateNotePos(parentView, parentNote);

            // show control panel
            ViewControlPanel.Visibility = Visibility.Visible;
            scaleCb.SelectionChanged += scaleCb_SelectionChanged;
            DG1.IsEnabled = false;
            wclCb.IsEnabled = false;
        }

        private void topCh_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                ModelDoc2 swModel = swApp.IActiveDoc2;
                DrawingDoc swDraw = (DrawingDoc)swModel;
                swModel.Extension.SelectByID2(parentView.Name, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
                double[] outline = (double[])parentView.GetOutline();
                topView = swDraw.CreateUnfoldedViewAt3((outline[2] + outline[0])/2, outline[3] + 0.01, 0, true);
                topView.AlignWithView((int)swAlignViewTypes_e.swAlignViewVerticalCenter, parentView);
                backCh.IsEnabled = true;
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void topCh_Unchecked(object sender, RoutedEventArgs e)
        {
            try 
            {
                ModelDoc2 swModel = swApp.IActiveDoc2;
                swModel.Extension.SelectByID2(topView.GetName2(), "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
                swModel.DeleteSelection(false);
                backCh.IsEnabled = false;
                topView = null;

                allCh.Unchecked -= allCh_Unchecked;
                allCh.IsChecked = false;
                allCh.Unchecked += allCh_Unchecked;
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void rightCh_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                ModelDoc2 swModel = swApp.IActiveDoc2;
                DrawingDoc swDraw = (DrawingDoc)swModel;
                swModel.Extension.SelectByID2(parentView.Name, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
                double[] outline = (double[])parentView.GetOutline();
                rightView = swDraw.CreateUnfoldedViewAt3(outline[2] + 0.01, (outline[1] + outline[3]) / 2, 0, true);
                rightView.AlignWithView((int)swAlignViewTypes_e.swAlignViewHorizontalCenter, parentView);

                if (addBalloonRight.IsChecked == true)
                {
                    Macros macros = new Macros(swApp);
                    View[] view = new View[1];
                    view[0] = rightView;
                    macros.BalloonCutlist(view, out swNotes);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void rightCh_Unchecked(object sender, RoutedEventArgs e)
        {
            try
            {
                ModelDoc2 swModel = swApp.IActiveDoc2;
                swModel.Extension.SelectByID2(rightView.GetName2(), "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
                swModel.DeleteSelection(false);
                rightView = null;

                allCh.Unchecked -= allCh_Unchecked;
                allCh.IsChecked = false;
                allCh.Unchecked += allCh_Unchecked;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void backCh_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                ModelDoc2 swModel = swApp.IActiveDoc2;
                DrawingDoc swDraw = (DrawingDoc)swModel;
                swModel.Extension.SelectByID2(topView.Name, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
                double[] outline = (double[])topView.GetOutline();
                backView = swDraw.CreateUnfoldedViewAt3((outline[2] + outline[0]) / 2, outline[3] + 0.01, 0, true);
                backView.AlignWithView((int)swAlignViewTypes_e.swAlignViewVerticalCenter, topView);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void backCh_Unchecked(object sender, RoutedEventArgs e)
        {
            try
            {
                ModelDoc2 swModel = swApp.IActiveDoc2;
                swModel.Extension.SelectByID2(backView.GetName2(), "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
                swModel.DeleteSelection(false);
                backView = null;

                allCh.Unchecked -= allCh_Unchecked;
                allCh.IsChecked = false;
                allCh.Unchecked += allCh_Unchecked;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void leftCh_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                ModelDoc2 swModel = swApp.IActiveDoc2;
                DrawingDoc swDraw = (DrawingDoc)swModel;
                swModel.Extension.SelectByID2(parentView.Name, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
                double[] outline = (double[])parentView.GetOutline();
                leftView = swDraw.CreateUnfoldedViewAt3(outline[0] - 0.01, (outline[1] + outline[3]) / 2, 0, true);
                leftView.AlignWithView((int)swAlignViewTypes_e.swAlignViewHorizontalCenter, parentView);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void leftCh_Unchecked(object sender, RoutedEventArgs e)
        {
            try
            {
                ModelDoc2 swModel = swApp.IActiveDoc2;
                swModel.Extension.SelectByID2(leftView.GetName2(), "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
                swModel.DeleteSelection(false);
                leftView = null;

                allCh.Unchecked -= allCh_Unchecked;
                allCh.IsChecked = false;
                allCh.Unchecked += allCh_Unchecked;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void btmCh_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                ModelDoc2 swModel = swApp.IActiveDoc2;
                DrawingDoc swDraw = (DrawingDoc)swModel;
                swModel.Extension.SelectByID2(parentView.Name, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
                double[] outline = (double[])parentView.GetOutline();
                btmView = swDraw.CreateUnfoldedViewAt3((outline[2] + outline[0]) / 2, outline[1] - 0.01, 0, true);
                btmView.AlignWithView((int)swAlignViewTypes_e.swAlignViewVerticalCenter, parentView);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void btmCh_Unchecked(object sender, RoutedEventArgs e)
        {
            try
            {
                ModelDoc2 swModel = swApp.IActiveDoc2;
                swModel.Extension.SelectByID2(btmView.GetName2(), "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
                swModel.DeleteSelection(false);
                btmView = null;

                allCh.Unchecked -= allCh_Unchecked;
                allCh.IsChecked = false;
                allCh.Unchecked += allCh_Unchecked;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void allCh_Checked(object sender, RoutedEventArgs e)
        {
            topCh.IsChecked = true;
            backCh.IsChecked = true;
            leftCh.IsChecked = true;
            rightCh.IsChecked = true;
            btmCh.IsChecked = true;
        }

        private void allCh_Unchecked(object sender, RoutedEventArgs e)
        {
            backCh.IsChecked = false;
            topCh.IsChecked = false;
            leftCh.IsChecked = false;
            rightCh.IsChecked = false;
            btmCh.IsChecked = false;
        }
        #endregion

        // Handle Check add balloon check box
        private void addBalloonRight_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                if(rightView != null)
                {
                    Macros macros = new Macros(swApp);
                    View[] view = new View[1];
                    view[0] = rightView;
                    macros.BalloonCutlist(view, out swNotes);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void addBalloonRight_Unchecked(object sender, RoutedEventArgs e)
        {
            try
            {
                foreach(Note note in swNotes)
                {
                    Annotation swAnn = note.IGetAnnotation();
                    swAnn.Select3(false, null);
                    ModelDoc2 swModel = swApp.IActiveDoc2;
                    swModel.Extension.DeleteSelection2((int)swDeleteSelectionOptions_e.swDelete_Absorbed);
                    swModel.ClearSelection2(true);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        // Method to rotate parent view
        private void alignLongestEdgeBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Macros macros = new Macros(swApp);
                MathUtility swMath = swApp.IGetMathUtility();
                ModelDoc2 swModel = swApp.IActiveDoc2;

                // check if any document is opened
                if (swModel == null)
                {
                    MessageBox.Show("No active document found.");
                    return;
                }

                if (swModel.GetType() != (int)swDocumentTypes_e.swDocDRAWING)
                {
                    MessageBox.Show("Only support drawing doc.");
                    return;
                }

                DrawingDoc swDraw = (DrawingDoc)swModel;

                Edge longestEdge = default(Edge);
                double angle;

                longestEdge = macros.GetLongestEdge(parentView);
                angle = macros.GetEdgeAngle(swMath, parentView, longestEdge);

                // Must select the view first before change orientation
                swModel.Extension.SelectByID2(parentView.Name, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);

                swDraw.DrawingViewRotate(angle);

                swModel.ClearSelection2(true);

                // update note position
                updateNotePos(parentView, parentNote);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        // Method to dimension the profile
        private void dimProfileBtn_Click(object sender, RoutedEventArgs e)
        {
            bool inputDimDefVal = swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate);
            bool dimSnappingVal = swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference);
            swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);
            swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, false);

            try
            {
                Macros macros = new Macros(swApp);
                View[] vView = new View[1]; 
                View processView = default(View);
                switch (processViewCb.Text)
                {
                    case "PARENT":
                        {
                            processView = parentView;
                            break;
                        }
                    case "TOP":
                        {
                            processView = topView;
                            break;
                        }         
                    case "BTM":
                        {
                            processView = btmView;
                            break;
                        }         
                    case "LEFT":
                        {
                            processView = leftView;
                            break;
                        }      
                    case "RIGHT":
                        {
                            processView = rightView;
                            break;
                        }        
                    case "BACK":
                        {
                            processView = backView;
                            break;
                        }
                    default:
                        {
                            break;
                        }
                }

                if(processView != null)
                {
                    vView[0] = processView;
                    macros.DimensionView(vView);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.StackTrace.ToString());
            }

            swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, inputDimDefVal);
            swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, dimSnappingVal);
        }

        // method to handle confirm Btn
        private void confirmBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // get selected row and set the iscreated column to true
                CutlistItem selectedRow = (CutlistItem)DG1.SelectedItem;
                selectedRow.IsCreated = true;
                DG1.Items.Refresh();

                // reset control
                resetControl();
                
                // hide the control panel
                ViewControlPanel.Visibility = Visibility.Hidden;

                // enable the datagrid
                DG1.IsEnabled = true;

                // enable wcl combobox
                wclCb.IsEnabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        // Method to handle cancel Btn
        private void cancelBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                ModelDoc2 swModel = swApp.IActiveDoc2;
                if (parentView != null)
                {
                    swModel.Extension.SelectByID2(parentView.GetName2(), "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
                    swModel.DeleteSelection(false);
                    parentView = null;
                }
                if (topView != null)
                {
                    swModel.Extension.SelectByID2(topView.GetName2(), "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
                    swModel.DeleteSelection(false);
                    topView = null;
                }
                if (btmView != null)
                {
                    swModel.Extension.SelectByID2(btmView.GetName2(), "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
                    swModel.DeleteSelection(false);
                    btmView = null;
                }
                if (leftView != null)
                {
                    swModel.Extension.SelectByID2(leftView.GetName2(), "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
                    swModel.DeleteSelection(false);
                    leftView = null;
                }
                if (rightView != null)
                {
                    swModel.Extension.SelectByID2(rightView.GetName2(), "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
                    swModel.DeleteSelection(false);
                    rightView = null;
                }
                if (backView != null)
                {
                    swModel.Extension.SelectByID2(backView.GetName2(), "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
                    swModel.DeleteSelection(false);
                    backView = null;
                }
                resetControl();
                ViewControlPanel.Visibility = Visibility.Hidden;
                DG1.IsEnabled = true;
                wclCb.IsEnabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }

        private void resetControl()
        {
            #region uncheck all check box
            topCh.Unchecked -= topCh_Unchecked;
            topCh.IsChecked = false;
            topCh.Unchecked += topCh_Unchecked;

            btmCh.Unchecked -= btmCh_Unchecked;
            btmCh.IsChecked = false;
            btmCh.Unchecked += btmCh_Unchecked;

            leftCh.Unchecked -= leftCh_Unchecked;
            leftCh.IsChecked = false;
            leftCh.Unchecked += leftCh_Unchecked;

            rightCh.Unchecked -= rightCh_Unchecked;
            rightCh.IsChecked = false;
            rightCh.Unchecked += rightCh_Unchecked;

            backCh.Unchecked -= backCh_Unchecked;
            backCh.IsChecked = false;
            backCh.Unchecked += backCh_Unchecked;

            allCh.Unchecked -= allCh_Unchecked;
            allCh.IsChecked = false;
            allCh.Unchecked += allCh_Unchecked;

            addBalloonRight.IsChecked = true;
            #endregion

            #region Check view as null
            parentView = null;
            topView = null;
            btmView = null;
            leftView = null;
            rightView = null;
            backView = null;
            #endregion
        }

        private void updateNotePos(View view, Note note)
        {
            Annotation swAnn = note.IGetAnnotation();
            double[] outline = (double[])view.GetOutline();
            swAnn.SetPosition((outline[2] + outline[0])/2, outline[1], 0);
        }
    }
}
