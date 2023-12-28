using mStructural.Classes;
using mStructural.Function;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace mStructural.WPF
{
    /// <summary>
    /// Interaction logic for ExportDxfWPF.xaml
    /// </summary>
    public partial class ExportDxfWPF : Window
    {
        #region Private Variables
        private SldWorks swApp;
        private ModelDoc2 swModel;
        private DrawingDoc swDraw;
        private ObservableCollection<ExportDxfClass> myClassOc;
        private string firstSheetName;
        private int progressPercentage = 0;
        #endregion

        // Constructor
        public ExportDxfWPF(SldWorks swapp, ModelDoc2 swmodel)
        {
            InitializeComponent();
            swApp = swapp;
            swModel = swmodel;
            myClassOc = new ObservableCollection<ExportDxfClass>();

            PopulateSheetNames();
            DG1.ItemsSource = myClassOc;
        }

        // Method to populate the sheet name in datagrid
        private void PopulateSheetNames()
        {
            // cast model doc to drawing doc
            swDraw = (DrawingDoc)swModel;

            // get all sheetnames
            string[] vSheetNames = (string[])swDraw.GetSheetNames();

            // store first sheet
            firstSheetName = vSheetNames[0];

            // create new class item
            foreach (string vSheetName in vSheetNames)
            {
                if (vSheetName.ToUpper().Contains("DXF"))
                {
                    ExportDxfClass exportDxfClass = new ExportDxfClass();
                    exportDxfClass.SheetName = vSheetName;
                    exportDxfClass.Include = true;
                    exportDxfClass.AmendQty = true;
                    myClassOc.Add(exportDxfClass);
                }
            }
        }

        // magical codes to set single click to check or uncheck
        private void DataGridCell_PreviewMouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            DataGridCell cell = (DataGridCell)sender;
            if (cell != null && !cell.IsEditing && !cell.IsReadOnly)
            {
                if (!cell.IsFocused)
                {
                    cell.Focus();
                }
                if (DG1.SelectionUnit != DataGridSelectionUnit.FullRow)
                {
                    if (!cell.IsSelected)
                    {
                        cell.IsSelected = true;
                    }
                }
                else
                {
                    DataGridRow row = FindVisualParent<DataGridRow>(cell);
                    if (row != null && !row.IsSelected)
                    {
                        row.IsSelected = true;
                    }
                }

            }
        }

        static T FindVisualParent<T>(UIElement element) where T : UIElement
        {
            UIElement parent = element;
            while (parent != null)
            {
                T correctlyTyped = parent as T;
                if (correctlyTyped != null)
                {
                    return correctlyTyped;
                }
                parent = VisualTreeHelper.GetParent(parent) as UIElement;
            }
            return null;
        }

        // Method when export button is pressed
        private void ExportBtn_Click(object sender, RoutedEventArgs e)
        {
            // get to production variable
            bool? toProduction = Export2ProCh.IsChecked;

            // get revision custom property
            CustomPropertyManager swCusPropMgr = swModel.Extension.CustomPropertyManager[""];
            string val = "";
            string valout = "";
            bool bRet = swCusPropMgr.Get4("Revision", false, out val, out valout);

            // initialize string builder
            StringBuilder sb = new StringBuilder();

            // set dxf export option to export active sheet only
            swApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfMultiSheetOption, (int)swDxfMultisheet_e.swDxfActiveSheetOnly);

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
                ppiPathName = ppiPathName.Substring(0, index) + "\\Production\\Pressed Plates - inhouse fab";

                // add production and pressed plates - external fab in the path
                ppePathName = ppePathName.Substring(0, index) + "\\Production\\Pressed Plates - external fab";
            }

            // get unique dependencies of this drawing
            List<string> drawDependencises = new List<string>();
            Hashtable drawDependsTable = getDependency(swModel.GetPathName());

            // Create a list and store the unique dependencies into it
            foreach (DictionaryEntry depend in drawDependsTable)
            {
                drawDependencises.Add(depend.Value.ToString());
            }

            // Progress bar variable
            double dProgressIncrement = 100 / myClassOc.Count;
            int iPercentageIncrement = Convert.ToInt32(dProgressIncrement);
            double dProgressPercentage = 0;

            // process each row if included
            foreach (ExportDxfClass myClass in myClassOc)
            {
                // update progress bar
                 dProgressPercentage += dProgressIncrement;
                progressPercentage = Convert.ToInt32(dProgressPercentage);

                BackgroundWorker worker = new BackgroundWorker();
                worker.WorkerReportsProgress = true;
                worker.DoWork += worker_DoWork;
                worker.ProgressChanged += worker_ProgressChanged;

                worker.RunWorkerAsync();

                if (myClass.Include)
                {
                    string sheetName = myClass.SheetName;
                    int err = -1;
                    int warn = -1;

                    // activate the sheet
                    swDraw.ActivateSheet(sheetName);

                    #region Check if note are tally with balloon
                    // get first view
                    View checkView = swDraw.IGetFirstView();

                    // get next view which is the actual view
                    checkView = checkView.IGetNextView();

                    // if there is view
                    while (checkView != null)
                    {
                        // get first note
                        Note checkNote = checkView.IGetFirstNote();
                        Note appendNote = checkNote;
                        string bomPNText = "";
                        string notePNText = "";
                        string noteBNText = "";
                        string originalNoteText = "";

                        // if there is note
                        while (checkNote != null)
                        {
                            string text = checkNote.GetText();

                            // check if it is balloon
                            if (checkNote.IsBomBalloon())
                            {
                                // get balloon text
                                bomPNText = text;
                            }
                            else if (text.Contains("PN:"))
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
                                noteBNText = originalNoteText.Substring(quantityIndex + 5, (originalNoteText.Length - 4) - (quantityIndex + 5));
                            }

                            // get next note
                            checkNote = checkNote.IGetNext();
                        }

                        // compare
                        if (bomPNText != notePNText)
                        {
                            // change text if necessary
                            originalNoteText = originalNoteText.Replace("PN: " + notePNText, "PN: " + bomPNText);
                            appendNote.SetText(originalNoteText);

                            // log
                            sb.AppendLine("Amended part number item: " + bomPNText + " in View: " + checkView.Name + " in Sheet: " + sheetName);
                        }

                        if (myClass.AmendQty == true && appendNote!=null)
                        {
                            // calculate quantity
                            CountTotalQtyFromDrawView countFunction = new CountTotalQtyFromDrawView(swApp);
                            int finalQty = countFunction.GetViewQuantity(checkView, drawDependencises);

                            // log 0 quantity
                            if(finalQty == 0)
                            {
                                sb.AppendLine("0 quantity is calculated in View: " + checkView.Name + " in Sheet: " + sheetName);
                            }

                            // change text
                            originalNoteText = originalNoteText.Replace("QTY: " + noteBNText, "QTY: " + finalQty);
                            appendNote.SetText(originalNoteText);
                        }

                        // get next view
                        checkView = checkView.IGetNextView();
                    }
                    #endregion

                    if (sheetName.ToLower().EndsWith("p"))
                    {
                        // get file dxf save name
                        // v1.7 change to external
                        string dxfSaveName = ppePathName + "\\" + sFileName + "_" + valout.ToUpper() + "_" + sheetName.ToLower() + ".dxf";

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
                                    bRet = swModel.Extension.SelectByID2(featBendline.Name + "@" + viewDrawComp.Name + "@" + swView.Name, "SKETCH", 0, 0, 0, false, 0, null, 0);

                                    // hide it
                                    swModel.BlankSketch();
                                }
                            }
                            #endregion

                            // get the reference model name of the view
                            string modelName = swView.GetReferencedModelName();

                            // open the model in the back
                            ModelDoc2 viewModel = swApp.OpenDoc6(modelName, (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", err, warn);

                            if (viewModel != null)
                            {
                                // if opened the model without any issue
                                // save the model
                                viewModel.Extension.SaveAs3(stpSaveName, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, null, ref err, ref warn);

                                // close the model 
                                swApp.CloseDoc(modelName);
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
                        string saveName = fpPathName + "\\" + sFileName + "_" + valout.ToUpper() + "_" + sheetName + ".dxf";

                        // make sure the directory exist before save
                        Directory.CreateDirectory(fpPathName);

                        swModel.Extension.SaveAs3(saveName, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, null, ref err, ref warn);
                    }
                }
            }

            // switch back to first sheet
            swDraw.ActivateSheet(firstSheetName);

            // close this form
            Close();

            // Complete
            MessageBox.Show("Save DXF completed!\r\n" + sb.ToString());
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

        // method to update progress bar
        private void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            ExportPb.Value = e.ProgressPercentage;
        }

        private void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            worker.ReportProgress(progressPercentage);
        }
    }
}
