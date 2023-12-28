using SolidWorks.Interop.sldworks;
using System.Windows;
using Ookii.Dialogs.Wpf;
using SolidWorks.Interop.swconst;
using System.Windows.Controls;
using System.Collections.ObjectModel;
using mStructural.Classes;
using System.IO;
using System.Text;
using System;
using mStructural.Function;
using System.Diagnostics;
using System.Windows.Media;
using System.Collections;

namespace mStructural.WPF
{
    /// <summary>
    /// Interaction logic for ExportProjectWPF.xaml
    /// </summary>
    public partial class ExportProjectWPF : Window
    {
        #region Private Variables
        private SldWorks swApp;
        private Message msg;
        private string Revision;
        private ObservableCollection<ExportBomClass> bomCol;
        private ObservableCollection<ExportModelClass> eDrawCol;
        private ObservableCollection<ExportDxfClass> dxfCol;
        private ObservableCollection<ExportPdfClass> pdfCol;
        private string BomCurrentDrawPath;
        private string ModelCurrentDrawPath;
        private string DxfCurrentDrawPath;
        private string PdfCurrentDrawPath;
        private DependancyUtil dependancyUtil;
        #endregion

        public ExportProjectWPF(SldWorks swapp)
        {
            swApp = swapp;
            msg = new Message(swApp);
            dependancyUtil = new DependancyUtil();
            BomCurrentDrawPath = "";
            ModelCurrentDrawPath = "";
            DxfCurrentDrawPath = "";
            PdfCurrentDrawPath = "";
            Revision = "-";
            InitializeComponent();
        }

        private void UpdateBomDatagrid()
        {
            // if the drawing in drawing path does not exist, then do nothing
            if (!File.Exists(drawingPathTb.Text))
            {
                BomTubeLaserTabWarnTxt.Text = "File not found, please go back to main tab and check your drawing path.";
                bomCol = null;
                BomDg.ItemsSource = null;
                return;
            }

            // check if the path is drawing doc
            int err = -1;
            int warn = -1;
            ModelDoc2 swModel = swApp.OpenDoc6(drawingPathTb.Text, (int)swDocumentTypes_e.swDocDRAWING, (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", ref err, ref warn);
            if(swModel == null)
            {
                BomTubeLaserTabWarnTxt.Text = "Only support drawing file. Check your drawing path.";
                bomCol = null;
                BomDg.ItemsSource = null;
                return;
            }

            // reset warning text
            BomTubeLaserTabWarnTxt.Text = "";

            // check if drawing path has changed
            if(BomCurrentDrawPath != drawingPathTb.Text)
            {
                BomCurrentDrawPath = drawingPathTb.Text;

                // clear current source
                bomCol = new ObservableCollection<ExportBomClass>();

                // initialize custom prop manager
                CustomPropertyManager swCusPropMgr = swModel.Extension.CustomPropertyManager[""];

                // get revision custom prop
                string val = "";
                string valout = "";
                bool bRet = swCusPropMgr.Get4("Revision", false, out val, out valout);
                Revision = valout.ToUpper();

                // cast modeldoc to drawing
                DrawingDoc swDraw = swModel as DrawingDoc;
                object[] arrSheets = swDraw.GetViews();
                foreach( object sheet in arrSheets )
                {
                    object[] arrViews = (object[])sheet;
                    View swView = (View)arrViews[0];
                    if (swView.GetTableAnnotationCount() > 0)
                    {
                        object[] arrTableAnns = swView.GetTableAnnotations();
                        foreach( object arrTableAnn in arrTableAnns)
                        {
                            TableAnnotation swTableAnn = (TableAnnotation)arrTableAnn;
                            if(swTableAnn.Type == (int)swTableAnnotationType_e.swTableAnnotation_WeldmentCutList || 
                                swTableAnn.Type == (int)swTableAnnotationType_e.swTableAnnotation_BillOfMaterials)
                            {
                                string balloonRef = swTableAnn.DisplayedText2[1, 0, false];
                                ObservableCollection<string> configCol = new ObservableCollection<string>();
                                bRet = false;
                                ModelDoc2 viewModel = null;
                                string viewConfig = "";

                                // check each view to get the model
                                for(int i = 1;i<arrViews.Length;i++)
                                {
                                    View checkView = (View)arrViews[i];
                                    Note swNote = checkView.IGetFirstNote();
                                    while(swNote != null)
                                    {
                                        if(swNote.GetBomBalloonText(true) == balloonRef)
                                        {
                                            viewConfig = checkView.ReferencedConfiguration;
                                            viewModel = checkView.ReferencedDocument;
                                            string[] configNames = viewModel.GetConfigurationNames();
                                            foreach(string configName in configNames)
                                            {
                                                configCol.Add(configName);
                                                if (configName.Contains("laser", StringComparison.OrdinalIgnoreCase))
                                                {
                                                    viewConfig = configName;
                                                }
                                            }
                                            bRet = true;
                                            break;
                                        }
                                        swNote = swNote.IGetNext();
                                    }

                                    if(bRet)
                                    {
                                        break;
                                    }
                                }

                                ExportBomClass epc = new ExportBomClass();

                                string featName = "";
                                string bomConfigName = "";
                                bool bExportTubeLaser = false;
                                switch (swTableAnn.Type)
                                {
                                    case (int)swTableAnnotationType_e.swTableAnnotation_WeldmentCutList:
                                        {
                                            WeldmentCutListAnnotation clAnn = swTableAnn as WeldmentCutListAnnotation;
                                            WeldmentCutListFeature clFeat = clAnn.WeldmentCutListFeature;
                                            featName = clFeat.GetFeature().Name;
                                            bomConfigName = clFeat.Configuration;
                                            bExportTubeLaser = true;
                                            break;
                                        }
                                    case (int)swTableAnnotationType_e.swTableAnnotation_BillOfMaterials:
                                        {
                                            BomTableAnnotation bomAnn = swTableAnn as BomTableAnnotation;
                                            BomFeature bomFeat = bomAnn.BomFeature;
                                            switch (bomFeat.TableType)
                                            {
                                                case (int)swBomType_e.swBomType_Indented:
                                                case (int)swBomType_e.swBomType_PartsOnly:
                                                    {
                                                        bomConfigName = bomFeat.Configuration;
                                                        break;
                                                    }
                                                case (int)swBomType_e.swBomType_TopLevelOnly:
                                                    {
                                                        object bomVisibleRef = null;
                                                        bomConfigName = bomFeat.GetConfigurations(true, bomVisibleRef)[0];
                                                        break;
                                                    }
                                            }
                                            featName = bomFeat.GetFeature().Name;
                                            break;
                                        }
                                }

                                epc.TableName = featName;
                                epc.SwTableAnn = swTableAnn;
                                epc.SwModel = viewModel;
                                epc.BomConfiguration = bomConfigName;
                                epc.BalloonRef = balloonRef;
                                epc.ExportTubeLaser = bExportTubeLaser;
                                epc.TubeLaserConfiguration = viewConfig;
                                epc.TubeLaserConfigurationList = configCol;
                                bomCol.Add(epc);
                            }
                        }
                    }
                }

                // update datasource
                BomDg.ItemsSource = bomCol;
            }
        }

        private void UpdateModelTab()
        {
            int err = -1;
            int warn = -1;
            
            // if the drawing in drawing path does not exist, then do nothing
            if (!File.Exists(drawingPathTb.Text))
            {
                ModelTabWarnTxt.Text = "File not found, please go back to main tab and check your drawing path.";
                eDrawCol = null;
                edrawingLb.ItemsSource = null;
                TubeLaserCb.Items.Clear();
                return;
            }

            // check if the path is drawing doc
            ModelDoc2 swModel = swApp.OpenDoc6(drawingPathTb.Text, (int)swDocumentTypes_e.swDocDRAWING, (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", ref err, ref warn);
            if (swModel == null)
            {
                ModelTabWarnTxt.Text = "Only support drawing file. Check your drawing path.";
                eDrawCol = null;
                edrawingLb.ItemsSource = null;
                TubeLaserCb.Items.Clear();
                return;
            }

            // check if drawing path has changed
            if (ModelCurrentDrawPath != drawingPathTb.Text)
            {
                ModelCurrentDrawPath = drawingPathTb.Text;

                // clear current selection
                eDrawCol = new ObservableCollection<ExportModelClass>();
                edrawingLb.ItemsSource = null;
                TubeLaserCb.Items.Clear();
                ModelTabWarnTxt.Text = "'";
                
                // initialize custom prop manager
                CustomPropertyManager swCusPropMgr = swModel.Extension.CustomPropertyManager[""];

                // get revision custom prop
                string val = "";
                string valout = "";
                bool bRet = swCusPropMgr.Get4("Revision", false, out val, out valout);
                Revision = valout.ToUpper();

                // populate list for the combobox
                Hashtable drawDep = dependancyUtil.getDependancy(swApp, drawingPathTb.Text);
                if (drawDep.Count > 0)
                {
                    foreach(DictionaryEntry entry in drawDep)
                    {
                        TopLevelModelCb.Items.Add(entry.Value.ToString());
                    }
                }

                // get top level assembly
                string topLevelPathName = dependancyUtil.GetTopModelStrValue(swApp, drawingPathTb.Text);
                if(topLevelPathName != "")
                {
                    // update status text
                    TopLevelModelCb.SelectedValue = topLevelPathName;
                
                    // check doc type
                    string ext = topLevelPathName.Substring(topLevelPathName.Length - 6).ToUpper();
                    int emDocType = -1;
                    if(ext == "SLDASM")
                    {
                        emDocType = (int)swDocumentTypes_e.swDocASSEMBLY;
                    }
                    else if(ext == "SLDPRT")
                    {
                        emDocType = (int)swDocumentTypes_e.swDocPART;
                    }

                    // open doc
                    ModelDoc2 emModel = swApp.OpenDoc6(topLevelPathName, emDocType,(int)swOpenDocOptions_e.swOpenDocOptions_Silent,"",ref err, ref warn);

                    // get all configuration
                    string[] configNames = emModel.GetConfigurationNames();
                    string tubeLaserCbSelectedValue = configNames[0];

                    // add to list box
                    foreach(string configName in configNames)
                    {
                        ExportModelClass emc = new ExportModelClass();
                        emc.Configuration = configName;
                        emc.IsChecked = false;

                        eDrawCol.Add(emc);
                        TubeLaserCb.Items.Add(configName);
                        if (configName.Contains("laser", StringComparison.OrdinalIgnoreCase))
                        {
                            tubeLaserCbSelectedValue = configName;
                        }
                    }

                    // check on first config
                    if (eDrawCol.Count > 0)
                    {
                        eDrawCol[0].IsChecked = true;
                    }

                    TubeLaserCb.SelectedValue = tubeLaserCbSelectedValue;
                    edrawingLb.ItemsSource = eDrawCol;
                }
                else
                {
                    ModelTabWarnTxt.Text = "Enable to find top level model. Check your drawing path.";
                }
            }
        }

        private void UpdateDxfTab()
        {
            // if the drawing in drawing path does not exist, then do nothing
            if (!File.Exists(drawingPathTb.Text))
            {
                DxfTabWarnTxt.Text = "File not found, please go back to main tab and check your drawing path.";
                dxfCol = null;
                DxfDg.ItemsSource = null;
                return;
            }

            // check if the path is drawing doc
            int err = -1;
            int warn = -1;
            ModelDoc2 swModel = swApp.OpenDoc6(drawingPathTb.Text, (int)swDocumentTypes_e.swDocDRAWING, (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", ref err, ref warn);
            if (swModel == null)
            {
                DxfTabWarnTxt.Text = "Only support drawing file. Check your drawing path.";
                dxfCol = null;
                DxfDg.ItemsSource = null;
                return;
            }

            // reset warning text
            DxfTabWarnTxt.Text = "";

            // check if drawing path has changed
            if(DxfCurrentDrawPath != drawingPathTb.Text)
            {
                DxfCurrentDrawPath = drawingPathTb.Text;

                // clear current source
                dxfCol = new ObservableCollection<ExportDxfClass>();

                // initialize custom prop manager
                CustomPropertyManager swCusPropMgr = swModel.Extension.CustomPropertyManager[""];

                // get revision custom prop
                string val = "";
                string valout = "";
                bool bRet = swCusPropMgr.Get4("Revision", false, out val, out valout);
                Revision = valout.ToUpper();

                // cast model doc to drawing doc
                DrawingDoc swDraw = swModel as DrawingDoc;

                // get all sheetnames
                string[] vSheetNames = (string[])swDraw.GetSheetNames();

                // create new class item
                foreach (string vSheetName in vSheetNames)
                {
                    if (vSheetName.ToUpper().Contains("DXF"))
                    {
                        ExportDxfClass edc = new ExportDxfClass();
                        edc.SheetName = vSheetName;
                        edc.Include = true;
                        edc.AmendQty = false;
                        dxfCol.Add(edc);
                    }
                }

                // update datasource
                DxfDg.ItemsSource = dxfCol;
            }

        }

        private void UpdatePdfTab()
        {
            // if the drawing in drawing path does not exist, then do nothing
            if (!File.Exists(drawingPathTb.Text))
            {
                PdfTabWarnTxt.Text = "File not found, please go back to main tab and check your drawing path.";
                pdfCol = null;
                PdfDg.ItemsSource = null;
                return;
            }

            // check if the path is drawing doc
            int err = -1;
            int warn = -1;
            ModelDoc2 swModel = swApp.OpenDoc6(drawingPathTb.Text, (int)swDocumentTypes_e.swDocDRAWING, (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", ref err, ref warn);
            if (swModel == null)
            {
                PdfTabWarnTxt.Text = "Only support drawing file. Check your drawing path.";
                pdfCol = null;
                PdfDg.ItemsSource = null;
                return;
            }

            // reset warning text
            PdfTabWarnTxt.Text = "";

            // check if drawing path has changed
            if (PdfCurrentDrawPath != drawingPathTb.Text)
            {
                PdfCurrentDrawPath = drawingPathTb.Text;

                // clear current source
                pdfCol = new ObservableCollection<ExportPdfClass>();

                // initialize custom prop manager
                CustomPropertyManager swCusPropMgr = swModel.Extension.CustomPropertyManager[""];

                // get revision custom prop
                string val = "";
                string valout = "";
                bool bRet = swCusPropMgr.Get4("Revision", false, out val, out valout);
                Revision = valout.ToUpper();

                // cast model doc to drawing doc
                DrawingDoc swDraw = swModel as DrawingDoc;

                // get all sheetnames
                string[] vSheetNames = (string[])swDraw.GetSheetNames();

                foreach (string vSheetName in vSheetNames)
                {
                    ExportPdfClass epc = new ExportPdfClass();
                    epc.SheetName = vSheetName;
                    if (vSheetName.ToUpper().Contains("DXF"))
                    {
                        epc.Include = false;
                    }
                    else
                    {
                        epc.Include = true;
                    }
                    pdfCol.Add(epc);
                }

                // update datasource
                PdfDg.ItemsSource = pdfCol;
            }
        }

        private void BrowseBtn_Click(object sender, RoutedEventArgs e)
        {
            // instantiate ookii open file dialog
            VistaOpenFileDialog ofd = new VistaOpenFileDialog();

            // set filter to solidworksw drawing only
            ofd.Filter = "SolidWorks Drawings (*.slddrw)|*.slddrw";

            // if user clicked ok, set the drawing path textbox to the selected drawing path
            if (ofd.ShowDialog() == true)
            {
                drawingPathTb.Text = ofd.FileName;
            }
        }

        private void GetActivePathBtn_Click(object sender, RoutedEventArgs e)
        {
            // set modeldoc to active doc
            ModelDoc2 swModel = swApp.IActiveDoc2;

            // check if there is any active doc
            if(swModel != null)
            {
                // check doc type, only support drawing
                if(swModel.GetType() == (int)swDocumentTypes_e.swDocDRAWING)
                {
                    drawingPathTb.Text = swModel.GetPathName();
                }
                else
                {
                    MessageBox.Show("Only support drawing document.");
                    return;
                }
            }
            else
            {
                MessageBox.Show("No active document.");
                return;
            }
        }

        private void BomCh_Checked(object sender, RoutedEventArgs e)
        {
            // enable bom and tube laser tab
            BomTab.IsEnabled = true;
            UpdateBomDatagrid();
        }

        private void BomCh_Unchecked(object sender, RoutedEventArgs e)
        {
            // disable bom and tube laser tab and clear check all check box
            BomTab.IsEnabled = false;
            CheckAllCh.Unchecked -= CheckAllCh_Unchecked;
            CheckAllCh.IsChecked = false;
            CheckAllCh.Unchecked += CheckAllCh_Unchecked;
        }

        private void CheckAllCh_Checked(object sender, RoutedEventArgs e)
        {
            // check on all checkbox
            BomCh.IsChecked = true;
            ModelCh.IsChecked = true;
            DxfCh.IsChecked = true;
            PdfCh.IsChecked = true;
        }

        private void CheckAllCh_Unchecked(object sender, RoutedEventArgs e)
        {
            // uncheck on all checkbox
            BomCh.IsChecked = false;
            ModelCh.IsChecked = false;
            DxfCh.IsChecked = false;
            PdfCh.IsChecked = false;
        }

        private void ConfirmBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // create log
                StringBuilder sb = new StringBuilder();
                Log(sb, "Initializing...");

                // if the drawing in drawing path does not exist, then do nothing
                if (!File.Exists(drawingPathTb.Text))
                {
                    Log(sb, "File does not exist.");
                }
                else
                {
                    // check if the path is drawing doc
                    int err = -1;
                    int warn = -1;
                    ModelDoc2 swModel = swApp.OpenDoc6(drawingPathTb.Text, (int)swDocumentTypes_e.swDocDRAWING, (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", ref err, ref warn);
                    if (swModel == null)
                    {
                        Log(sb, "Only support drawing file.");
                    }
                    else
                    {
                        if (BomCh.IsChecked == true)
                        {
                            ExportBom2 eb = new ExportBom2();
                            StringBuilder sb2 = eb.Run(swApp, drawingPathTb.Text, IsProductionCh.IsChecked ?? false, bomCol, Revision);
                            sb.Append(sb2);
                        }

                        if (ModelCh.IsChecked == true)
                        {
                            ExportModel2 em = new ExportModel2();
                            StringBuilder sb2 = em.Run(swApp, TopLevelModelCb.Text, IsProductionCh.IsChecked ?? false, EDrawingCh.IsChecked ?? false, TubeLaserCh.IsChecked ?? false, eDrawCol, TubeLaserCb.Text, Revision);
                            sb.Append(sb2);
                        }

                        if (DxfCh.IsChecked == true)
                        {
                            ExportDxf2 ed = new ExportDxf2();
                            StringBuilder sb2 = ed.Run(swApp, drawingPathTb.Text, IsProductionCh.IsChecked ?? false, dxfCol, Revision);
                            sb.Append(sb2);
                        }

                        if (PdfCh.IsChecked == true)
                        {
                            ExportPdf2 ep = new ExportPdf2();
                            StringBuilder sb2 = ep.Run(swApp, drawingPathTb.Text, IsProductionCh?.IsChecked ?? false, pdfCol, Revision);
                            sb.Append(sb2);
                        }
                    }
                }

                Log(sb, "End...");

                // get my document location
                string logPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments) + "\\mStructural Download\\Log";

                // create backup directory
                Directory.CreateDirectory(logPath);

                using (StreamWriter file = new StreamWriter(logPath + "\\ExportProjectLog.txt"))
                {
                    file.WriteLine(sb.ToString());
                }

                Close();
                Process.Start("notepad.exe", logPath + "\\ExportProjectLog.txt");
            }
            catch (Exception ex)
            {
                msg.ErrorMsg(ex.ToString());
            }
        }

        private void MainTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if(e.Source is TabControl)
            {
                TabControl tabControl = (TabControl)e.Source;

                // update information on selected tab
                if(tabControl.SelectedIndex == 1)
                {
                    UpdateBomDatagrid();
                }
                else if(tabControl.SelectedIndex == 2)
                {
                    UpdateModelTab();
                }
                else if(tabControl.SelectedIndex == 3)
                {
                    UpdateDxfTab();
                }
                else if(tabControl.SelectedIndex == 4)
                {
                    UpdatePdfTab();
                }
            }
        }

        private void ModelCh_Checked(object sender, RoutedEventArgs e)
        {
            ModelTab.IsEnabled = true;
            UpdateModelTab();
        }

        private void ModelCh_Unchecked(object sender, RoutedEventArgs e)
        {
            ModelTab.IsEnabled= false;
            CheckAllCh.Unchecked -= CheckAllCh_Unchecked;
            CheckAllCh.IsChecked = false;
            CheckAllCh.Unchecked += CheckAllCh_Unchecked;
        }

        private void DxfCh_Checked(object sender, RoutedEventArgs e)
        {
            DxfTab.IsEnabled = true;
            UpdateDxfTab();
        }

        private void DxfCh_Unchecked(object sender, RoutedEventArgs e)
        {
            DxfTab.IsEnabled= false;
            CheckAllCh.Unchecked -= CheckAllCh_Unchecked;
            CheckAllCh.IsChecked = false;
            CheckAllCh.Unchecked += CheckAllCh_Unchecked;
        }

        private void PdfCh_Checked(object sender, RoutedEventArgs e)
        {
            PdfTab.IsEnabled = true;
            UpdatePdfTab();
        }

        private void PdfCh_Unchecked(object sender, RoutedEventArgs e)
        {
            PdfTab.IsEnabled= false;
            CheckAllCh.Unchecked -= CheckAllCh_Unchecked;
            CheckAllCh.IsChecked = false;
            CheckAllCh.Unchecked += CheckAllCh_Unchecked;
        }

        private void Log(StringBuilder sb, string text)
        {
            sb.AppendLine(DateTime.Now.ToString() + " : " +text);
        }

        // magical codes to set single click to check or uncheck
        private void DxfDgCell_PreviewMouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            DataGridCell cell = (DataGridCell)sender;
            if (cell != null && !cell.IsEditing && !cell.IsReadOnly)
            {
                if (!cell.IsFocused)
                {
                    cell.Focus();
                }
                if (DxfDg.SelectionUnit != DataGridSelectionUnit.FullRow)
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

        private void PdfDgCell_PreviewMouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            DataGridCell cell = (DataGridCell)sender;
            if (cell != null && !cell.IsEditing && !cell.IsReadOnly)
            {
                if (!cell.IsFocused)
                {
                    cell.Focus();
                }
                if (PdfDg.SelectionUnit != DataGridSelectionUnit.FullRow)
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

        private void PdfCheckBtn_Click(object sender, RoutedEventArgs e)
        {
            // get selected row in datagrid
            IList rows = PdfDg.SelectedItems;
            foreach (var row in rows)
            {
                ExportPdfClass epc = (ExportPdfClass)row;
                epc.Include = true;
            }
            PdfDg.Items.Refresh();
        }

        private void PdfUncheckBtn_Click(object sender, RoutedEventArgs e)
        {
            // get selected row in datagrid
            IList rows = PdfDg.SelectedItems;
            foreach (var row in rows)
            {
                ExportPdfClass epc = (ExportPdfClass)row;
                epc.Include = false;
            }
            PdfDg.Items.Refresh();
        }

        private void DxfCheckBtn_Click(object sender, RoutedEventArgs e)
        {
            // get selected row in datagrid
            IList rows = DxfDg.SelectedItems;
            foreach (var row in rows)
            {
                ExportDxfClass edc = (ExportDxfClass)row;
                edc.Include = true;
            }
            DxfDg.Items.Refresh();
        }

        private void DxfUncheckBtn_Click(object sender, RoutedEventArgs e)
        {
            // get selected row in datagrid
            IList rows = DxfDg.SelectedItems;
            foreach (var row in rows)
            {
                ExportDxfClass edc = (ExportDxfClass)row;
                edc.Include = false;
            }
            DxfDg.Items.Refresh();
        }

        private void TopLevelModelCb_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBox cmb = (ComboBox)sender;
            if (!cmb.IsLoaded) return;

            string text = (sender as ComboBox).SelectedItem as string;

            // clear listbox and combobox for model tab
            eDrawCol = new ObservableCollection<ExportModelClass>();
            TubeLaserCb.Items.Clear();

            // check doc type
            string ext = text.Substring(text.Length - 6).ToUpper();
            int emDocType = -1;
            if (ext == "SLDASM")
            {
                emDocType = (int)swDocumentTypes_e.swDocASSEMBLY;
            }
            else if (ext == "SLDPRT")
            {
                emDocType = (int)swDocumentTypes_e.swDocPART;
            }

            // open doc
            int err = -1;
            int warn = -1;
            ModelDoc2 emModel = swApp.OpenDoc6(text, emDocType, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref err, ref warn);

            // get all configuration
            string[] configNames = emModel.GetConfigurationNames();

            // add to list box
            foreach (string configName in configNames)
            {
                ExportModelClass emc = new ExportModelClass();
                emc.Configuration = configName;
                emc.IsChecked = false;

                eDrawCol.Add(emc);
                TubeLaserCb.Items.Add(configName);
            }

            // check on first config
            if (eDrawCol.Count > 0)
            {
                eDrawCol[0].IsChecked = true;
            }

            TubeLaserCb.SelectedIndex = 0;
            edrawingLb.ItemsSource = eDrawCol;

            swApp.CloseDoc(emModel.GetPathName());
            swApp.ActivateDoc2(drawingPathTb.Text, false, ref err);
        }
    }
}
