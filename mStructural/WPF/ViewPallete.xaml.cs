using mStructural.Classes;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;

namespace mStructural.WPF
{
    /// <summary>
    /// Interaction logic for ViewPallete.xaml
    /// </summary>
    public partial class ViewPallete : Window
    {
        #region Private variables
        private SldWorks swApp;
        private ModelDoc2 swModel;
        private ModelDoc2 viewModel;
        private DrawingDoc swDraw;
        private ObservableCollection<CutlistItem2> cutListItems;
        private double[] viewPos = new double[2];
        private string configName;
        private double[] sheetProps = new double[7];

        private View swView;
        private View rightView;
        private View leftView;
        private View topView;
        private View btmView;
        private View backView;

        private int curCutlistIndex = 0;
        #endregion

        // constructor
        public ViewPallete(ObservableCollection<CutlistItem2> cutlistItems, SldWorks swapp, ModelDoc2 swmodel, ModelDoc2 viewmodel, string configname)
        {
            InitializeComponent();

            cutListItems = cutlistItems;
            swApp = swapp;
            swModel = swmodel;
            configName = configname;
            viewModel = viewmodel;

            // populate scale
            PopulateScales();

            // create first view
            swDraw = (DrawingDoc)swModel;
            Sheet swSheet = swDraw.IGetCurrentSheet();

            // get properties for current sheet
            sheetProps = (double[])swSheet.GetProperties();
            viewPos[1] = sheetProps[6];

            createView(cutListItems[curCutlistIndex].Name);
            positionView();
        }

        // Method to populate predefined scales
        private void PopulateScales()
        {
            // get document standard
            int drawingStandard = swModel.Extension.GetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swDetailingDimensionStandard,
                (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified);

            string standardString = "";

            switch (drawingStandard)
            {
                case 1:
                    standardString = "#ANSI";
                    break;
                case 2:
                    standardString = "#ISO";
                    break;
                case 3:
                    standardString = "#DIN";
                    break;
                case 4:
                    standardString = "#JIS";
                    break;
                case 5:
                    standardString = "#BSI";
                    break;
                case 6:
                    standardString = "#GOST";
                    break;
                case 7:
                    standardString = "#GB";
                    break;
                case 8:
                    standardString = "";
                    break;
            }

            // get user preferences
            string settingFileLoc = swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swFileLocationsDrawingScaleStandard);
            settingFileLoc = settingFileLoc + "/" + "drawingscales.txt";

            // check if setting file exist
            if(File.Exists(settingFileLoc))
            {
                foreach (string line in
                File.ReadLines(settingFileLoc)
                    .SkipWhile(line => !line.Contains(standardString))
                    .Skip(1)
                    .TakeWhile(line => !line.Contains(";;")))
                {
                    scaleCb.Items.Add(line);
                }
            }
            else
            {
                scaleCb.Items.Add("1:1");
                scaleCb.Items.Add("1:2");
                scaleCb.Items.Add("1:5");
                scaleCb.Items.Add("1:10");
                scaleCb.Items.Add("1:20");
                scaleCb.Items.Add("1:50");
                scaleCb.Items.Add("1:100");
                scaleCb.Items.Add("2:1");
                scaleCb.Items.Add("5:1");
                scaleCb.Items.Add("10:1");
                scaleCb.Items.Add("20:1");
                scaleCb.Items.Add("50:1");
                scaleCb.Items.Add("100:1");
            }

            if (scaleCb.Items.Count > 0)
            {
                scaleCb.SelectedIndex = 0;
            }
        }

        public bool selectViewBody(ModelDoc2 swModel, View swview, string clname)
        {
            bool bRet = false;
            Feature clFeat = default(Feature);
            BodyFolder swBodyFolder = default(BodyFolder);
            object[] arrBody = null;
            object[] bodies = new object[1];
            DispatchWrapper[] arrBodiesIn = new DispatchWrapper[1];

            // switch to appropriate configuration
            swModel.ShowConfiguration2(configName);

            // select the feature with cutlist name column
            bRet = swModel.Extension.SelectByID2(clname, "BDYFOLDER", 0, 0, 0, false, 0, null, 0);

            // if failed to select
            if (!bRet)
            {
                return false;
            }

            // cast selected object to feature object
            SelectionMgr swSelMgr = swModel.ISelectionManager;
            clFeat = (Feature)swSelMgr.GetSelectedObject6(1, -1);

            if (clFeat.GetTypeName2() == "CutListFolder")
            {
                // get body folder from feature
                swBodyFolder = (BodyFolder)clFeat.GetSpecificFeature2();

                // get all body from body folder
                arrBody = (object[])swBodyFolder.GetBodies();

                if (arrBody != null)
                {
                    // get first body body
                    Body2 swBody = (Body2)arrBody[0];

                    // check if body is hidden
                    if (!swBody.Visible)
                    {
                        for (int i = 1; i < arrBody.Length; i++)
                        {
                            swBody = (Body2)arrBody[i];
                            if (swBody.Visible)
                            {
                                // use the first body in the list
                                bodies[0] = arrBody[0];

                                // create new dispatch wrapper for view body selection
                                arrBodiesIn[0] = new DispatchWrapper(bodies[0]);

                                // select the body for the view
                                swview.Bodies = (arrBodiesIn);

                                break;
                            }
                        }
                    }
                    else
                    {
                        // use the first body in the list
                        bodies[0] = arrBody[0];

                        // create new dispatch wrapper for view body selection
                        arrBodiesIn[0] = new DispatchWrapper(bodies[0]);

                        // select the body for the view
                        swview.Bodies = (arrBodiesIn);
                    }
                }
            }

            // clear selection after done
            swModel.ClearSelection2(true);

            return bRet;
        }

        private void newSheetBtn_Click(object sender, RoutedEventArgs e)
        {
            // select the view
            swModel.Extension.SelectByID2(swView.Name, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);

            // delete view
            swModel.Extension.DeleteSelection2((int)swDeleteSelectionOptions_e.swDelete_Absorbed);

            // get the selected scale
            string selectedScale = scaleCb.Text;
            double[] dScale = new double[2];
            dScale[0] = Convert.ToDouble(selectedScale.Substring(0, selectedScale.IndexOf(":")));
            dScale[1] = Convert.ToDouble(selectedScale.Substring(selectedScale.IndexOf(":") + 1));

            // create new sheet
            swDraw.NewSheet4("", Convert.ToInt32(sheetProps[0]), Convert.ToInt32(sheetProps[1]),
                        dScale[0], dScale[1], Convert.ToBoolean(sheetProps[4]), swDraw.IGetCurrentSheet().GetTemplateName(), sheetProps[5], sheetProps[6], "",
                        0, 0, 0, 0, 0, 0);

            // re-create view
            createView(cutListItems[curCutlistIndex].Name);

            // activate sheet
            Sheet swSheet = swDraw.IGetCurrentSheet();
            swSheet.SetScale(dScale[0], dScale[1], true, true);

            // get new sheet data
            sheetProps = (double[])swSheet.GetProperties();

            // position view
            positionView();
        }

        private void createView(string clname)
        {
            swView = swDraw.CreateDrawViewFromModelView3(viewModel.GetPathName(), "*Front", 0, 0, 0);
            swView.SetDisplayTangentEdges2((int)swDisplayTangentEdges_e.swTangentEdgesVisibleAndFonted);
            swView.UseSheetScale = 1;
            swView.ReferencedConfiguration = configName;

            if (cutListItems.Count > 0)
            {
                bool bRet = selectViewBody(viewModel, swView, clname);
                if (!bRet)
                {
                    return;
                }
            }
        }

        private void positionView()
        {
            // position view
            double[] viewOutline = (double[])swView.GetOutline();
            double[] oriViewPos = (double[])swView.Position;
            double offsetX = oriViewPos[0] - (viewOutline[2] + viewOutline[0]) / 2;
            double offsetY = oriViewPos[1] - (viewOutline[3] + viewOutline[1]) / 2;
            viewPos[0] = viewPos[0] + offsetX;
            viewPos[1] = viewPos[1] + offsetY;

            swView.Position = viewPos;

            viewPos[0] = 0;
            viewPos[1] = sheetProps[6];
        }

        private void AllCh_Checked(object sender, RoutedEventArgs e)
        {
            RightCh.IsChecked = true;
            LeftCh.IsChecked = true;
            TopCh.IsChecked = true;
            BtmCh.IsChecked = true;
            BackCh.IsChecked = true;
        }

        private void AllCh_Unchecked(object sender, RoutedEventArgs e)
        {
            RightCh.IsChecked = false;
            LeftCh.IsChecked = false;
            TopCh.IsChecked = false;
            BtmCh.IsChecked = false;
            BackCh.IsChecked = false;
        }

        private void BtmCh_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                swModel.Extension.SelectByID2(swView.Name, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
                double[] outline = (double[])swView.GetOutline();
                btmView = swDraw.CreateUnfoldedViewAt3((outline[2] + outline[0]) / 2, outline[1] - 0.01, 0, true);
                btmView.AlignWithView((int)swAlignViewTypes_e.swAlignViewVerticalCenter, swView);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void BtmCh_Unchecked(object sender, RoutedEventArgs e)
        {
            try
            {
                swModel.Extension.SelectByID2(btmView.GetName2(), "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
                swModel.DeleteSelection(false);
                btmView = null;

                AllCh.Unchecked -= AllCh_Unchecked;
                AllCh.IsChecked = false;
                AllCh.Unchecked += AllCh_Unchecked;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void RightCh_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                swModel.Extension.SelectByID2(swView.Name, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
                double[] outline = (double[])swView.GetOutline();
                rightView = swDraw.CreateUnfoldedViewAt3(outline[2] + 0.01, (outline[1] + outline[3]) / 2, 0, true);
                rightView.AlignWithView((int)swAlignViewTypes_e.swAlignViewHorizontalCenter, swView);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void RightCh_Unchecked(object sender, RoutedEventArgs e)
        {
            try
            {
                swModel.Extension.SelectByID2(rightView.GetName2(), "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
                swModel.DeleteSelection(false);
                rightView = null;

                AllCh.Unchecked -= AllCh_Unchecked;
                AllCh.IsChecked = false;
                AllCh.Unchecked += AllCh_Unchecked;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void LeftCh_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                swModel.Extension.SelectByID2(swView.Name, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
                double[] outline = (double[])swView.GetOutline();
                leftView = swDraw.CreateUnfoldedViewAt3(outline[0] - 0.01, (outline[1] + outline[3]) / 2, 0, true);
                leftView.AlignWithView((int)swAlignViewTypes_e.swAlignViewHorizontalCenter, swView);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void LeftCh_Unchecked(object sender, RoutedEventArgs e)
        {
            try
            {
                swModel.Extension.SelectByID2(leftView.GetName2(), "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
                swModel.DeleteSelection(false);
                leftView = null;

                AllCh.Unchecked -= AllCh_Unchecked;
                AllCh.IsChecked = false;
                AllCh.Unchecked += AllCh_Unchecked;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void TopCh_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                swModel.Extension.SelectByID2(swView.Name, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
                double[] outline = (double[])swView.GetOutline();
                topView = swDraw.CreateUnfoldedViewAt3((outline[2] + outline[0]) / 2, outline[3] + 0.01, 0, true);
                topView.AlignWithView((int)swAlignViewTypes_e.swAlignViewVerticalCenter, swView);
                BackCh.IsEnabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void TopCh_Unchecked(object sender, RoutedEventArgs e)
        {
            try
            {
                swModel.Extension.SelectByID2(topView.GetName2(), "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
                swModel.DeleteSelection(false);
                BackCh.IsEnabled = false;
                topView = null;

                AllCh.Unchecked -= AllCh_Unchecked;
                AllCh.IsChecked = false;
                AllCh.Unchecked += AllCh_Unchecked;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void BackCh_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
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

        private void BackCh_Unchecked(object sender, RoutedEventArgs e)
        {
            try
            {
                swModel.Extension.SelectByID2(backView.GetName2(), "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
                swModel.DeleteSelection(false);
                backView = null;

                AllCh.Unchecked -= AllCh_Unchecked;
                AllCh.IsChecked = false;
                AllCh.Unchecked += AllCh_Unchecked;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void NextItemBtn_Click(object sender, RoutedEventArgs e)
        {
            // reset all view
            swView = null;
            rightView = null;
            leftView = null;
            topView = null;
            btmView = null;
            backView = null;

            // reset control
            AllCh.Unchecked -= AllCh_Unchecked;
            AllCh.IsChecked = false;
            AllCh.Unchecked += AllCh_Unchecked;
            RightCh.Unchecked -= RightCh_Unchecked;
            RightCh.IsChecked = false;
            RightCh.Unchecked += RightCh_Unchecked;
            LeftCh.Unchecked -= LeftCh_Unchecked;
            LeftCh.IsChecked = false;
            LeftCh.Unchecked += LeftCh_Unchecked;
            TopCh.Unchecked -= TopCh_Unchecked;
            TopCh.IsChecked = false;
            TopCh.Unchecked += TopCh_Unchecked;
            BtmCh.Unchecked -= BtmCh_Unchecked;
            BtmCh.IsChecked = false;
            BtmCh.Unchecked += BtmCh_Unchecked;
            BackCh.Unchecked -= BackCh_Unchecked;
            BackCh.IsChecked = false;
            BackCh.Unchecked += BackCh_Unchecked;

            curCutlistIndex += 1;
            if(curCutlistIndex > cutListItems.Count - 1)
            {
                MessageBox.Show("No more view to create");
                return;
            }

            createView(cutListItems[curCutlistIndex].Name);
            positionView();
        }

        private void NextSheetBtn_Click(object sender, RoutedEventArgs e)
        {
            // select the view
            swModel.Extension.SelectByID2(swView.Name, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);

            // delete view
            swModel.Extension.DeleteSelection2((int)swDeleteSelectionOptions_e.swDelete_Absorbed);

            // go to next sheet
            swDraw.SheetNext();

            // recreate view
            createView(cutListItems[curCutlistIndex].Name);
            positionView();
        }
    }
}
