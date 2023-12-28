using mStructural.Classes;
using mStructural.Function;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace mStructural.WPF
{
    /// <summary>
    /// Interaction logic for ExportPdfWPF.xaml
    /// </summary>
    public partial class ExportPdfWPF : Window
    {
        #region Private Variables
        private SldWorks swApp;
        private ModelDoc2 swModel;
        private DrawingDoc swDraw;
        private Message msg;
        private SwErrors swErr;
        private ObservableCollection<ExportPdfClass> exportCol;
        #endregion

        public ExportPdfWPF(SldWorks swapp, ModelDoc2 swmodel)
        {
            InitializeComponent();
            swApp = swapp;
            swModel = swmodel;
            msg = new Message(swApp);
            swErr = new SwErrors();

            // generate data grid
            exportCol = new ObservableCollection<ExportPdfClass>();

            // get all sheets of the drawing
            swDraw = (DrawingDoc)swModel;
            string[] sheetNames = (string[])swDraw.GetSheetNames();
            foreach (string sheetName in sheetNames)
            {
                ExportPdfClass exportPdfClass = new ExportPdfClass();
                exportPdfClass.SheetName = sheetName;
                if (sheetName.ToUpper().Contains("DXF"))
                {
                    exportPdfClass.Include = false;
                }
                else
                {
                    exportPdfClass.Include = true;
                }
                exportCol.Add(exportPdfClass);
            }

            DG1.ItemsSource = exportCol;
        }

        private void ExportBtn_Click(object sender, RoutedEventArgs e)
        {
            // get to production value
            bool toProduction = ProductionCh.IsChecked ?? false;

            // initialize custom prop manager
            CustomPropertyManager swCusPropMgr = swModel.Extension.CustomPropertyManager[""];

            // get revision custom prop
            string val = "";
            string valout = "";
            bool bRet = swCusPropMgr.Get4("Revision", false, out val, out valout);

            // get current model name
            string sFileName = Path.GetFileNameWithoutExtension(swModel.GetPathName());

            // get current path name
            string sPathName = Path.GetDirectoryName(swModel.GetPathName());

            // find the \
            int index = sPathName.LastIndexOf("\\");

            // add production in path
            if (toProduction == true)
            {
                sPathName = sPathName.Substring(0, index) + "\\Production";
            }

            // create directory
            Directory.CreateDirectory(sPathName);

            // get file save name
            string saveName = sPathName + "\\" + sFileName + "_Rev" + valout.ToUpper() + ".PDF";

            // get export data
            ExportPdfData swEpd = (ExportPdfData)swApp.GetExportFileData((int)swExportDataFileType_e.swExportPdfData);

            // initialize dispatch wrapper and intermediate object
            List<DispatchWrapper> dwList = new List<DispatchWrapper>();

            foreach(ExportPdfClass exportPdfClass in exportCol)
            {
                if (exportPdfClass.Include)
                {
                    Sheet swSheet = swDraw.Sheet[exportPdfClass.SheetName];
                    DispatchWrapper dw = new DispatchWrapper(swSheet);
                    dwList.Add(dw);
                }
            }

            DispatchWrapper[] dwArr = dwList.ToArray();

            // save the drawings to pdf
            int err = -1, warn = -1;
            bRet = swEpd.SetSheets((int)swExportDataSheetsToExport_e.swExportData_ExportSpecifiedSheets, (dwArr));
            swEpd.ViewPdfAfterSaving = true;
            bRet = swModel.Extension.SaveAs(saveName, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, swEpd, ref err, ref warn);

            // if error
            if (!bRet)
            {
                msg.ErrorMsg(swErr.GetFileSaveError(err));
            }
            else
            {
                msg.InfoMsg("Export PDF completed.");
            }

            Close();
        }

        private void CheckUncheckCh_Checked(object sender, RoutedEventArgs e)
        {
            // get selected row in datagrid
            IList rows = DG1.SelectedItems;
            foreach (var row in rows)
            {
                ExportPdfClass epc = (ExportPdfClass)row;
                epc.Include = true;
            }
            DG1.Items.Refresh();
        }

        private void CheckUncheckCh_Unchecked(object sender, RoutedEventArgs e)
        {
            // get selected row in datagrid
            IList rows = DG1.SelectedItems;
            foreach (var row in rows)
            {
                ExportPdfClass epc = (ExportPdfClass)row;
                epc.Include = false;
            }
            DG1.Items.Refresh();
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
    }
}
