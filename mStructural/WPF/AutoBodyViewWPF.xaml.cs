using mStructural.Classes;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace mStructural.WPF
{
    /// <summary>
    /// Interaction logic for AutoBodyViewWPF.xaml
    /// </summary>
    public partial class AutoBodyViewWPF : Window
    {
        #region Private Variables
        private SldWorks swApp;
        private ModelDoc2 swModel;
        private DrawingDoc swDraw;
        private TableAnnotation clTable;
        private ObservableCollection<CutlistItem2> cutList;
        private string configName;
        #endregion

        // Constructor
        public AutoBodyViewWPF(SldWorks swapp, ModelDoc2 swmodel, TableAnnotation cltable )
        {
            // grab the sldworks object
            swApp = swapp;

            // grab the modeldoc2 object
            swModel = swmodel;

            // grab the annotation table object
            clTable = cltable;
            
            InitializeComponent();

            // populate views in combo box and preselect select the first view
            swDraw = (DrawingDoc)swModel;
            View swView = swDraw.IGetFirstView(); // first view is sheet
            swView = swView.IGetNextView();
            while(swView != null )
            {
                RefViewCb.Items.Add( swView.Name);
                swView = swView.IGetNextView();
            }

            // select first combobox
            if(RefViewCb.Items.Count > 0)
            {
                RefViewCb.SelectedIndex = 0;
            }

            cutList = new ObservableCollection<CutlistItem2>();
            // lock row height
            for (int i = 1; i < clTable.RowCount; i++)
            {
                clTable.SetLockRowHeight(i, true);
            }

            // insert name row
            clTable.InsertColumn2((int)swTableItemInsertPosition_e.swTableItemInsertPosition_Last, 0, "Name", (int)swTableColumnTypes_e.swWeldTableColumnType_CutListName);
            clTable.SetColumnType(clTable.ColumnCount - 1, (int)swTableColumnTypes_e.swWeldTableColumnType_CutListName);

            // extract info
            for (int i = 1; i < clTable.RowCount; i++)
            {
                CutlistItem2 cli = new CutlistItem2();
                cli.ItemNo = Convert.ToInt32(clTable.DisplayedText2[i, 0, false]);
                cli.Name = clTable.DisplayedText2[i, clTable.ColumnCount - 1, false];
                cli.Include = true;
                cutList.Add(cli);
            }

            // delete name row
            clTable.DeleteColumn2(clTable.ColumnCount - 1, false);

            // unlock row height
            for (int i = 1; i < clTable.RowCount; i++)
            {
                clTable.SetLockRowHeight(i, false);
            }

            // get configuration
            WeldmentCutListAnnotation wclAnn = (WeldmentCutListAnnotation)cltable;
            WeldmentCutListFeature wclFeat = wclAnn.WeldmentCutListFeature;
            configName = wclFeat.Configuration;

            // get reference configuration
            DG1.ItemsSource = cutList;
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

        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            // get selected row in datagrid
            IList rows = DG1.SelectedItems;
            foreach(var row in rows)
            {
                CutlistItem2 cli = (CutlistItem2)row;
                cli.Include = true;
            }
            DG1.Items.Refresh();
        }

        private void CheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            // get selected row in datagrid
            IList rows = DG1.SelectedItems;
            foreach (var row in rows)
            {
                CutlistItem2 cli = (CutlistItem2)row;
                cli.Include = false;
            }
            DG1.Items.Refresh();
        }

        private void GenerateBtn_Click(object sender, RoutedEventArgs e)
        {
            #region Checking
            // populate confirmed list
            ObservableCollection<CutlistItem2> confirmedCl = new ObservableCollection<CutlistItem2>();
            foreach(CutlistItem2 cl in cutList)
            {
                if (cl.Include)
                {
                    confirmedCl.Add(cl);
                }
            }

            // check if data grid is empty
            if (cutList.Count == 0)
            {
                MessageBox.Show("Empty cutlist.");
                return;
            }

            // get view model
            swModel.Extension.SelectByID2(RefViewCb.Text, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
            SelectionMgr swSelMgr = swModel.ISelectionManager;
            if (swSelMgr.GetSelectedObjectCount2(-1) == 0)
            {
                MessageBox.Show("Invalid view.");
                return;
            }

            View swView = (View)swSelMgr.GetSelectedObject6(1,-1);
            ModelDoc2 viewModel = swView.ReferencedDocument;

            // check if view model is correct
            foreach(CutlistItem2 cli in confirmedCl)
            {
                if(!viewModel.Extension.SelectByID2(cli.Name,"BDYFOLDER", 0, 0, 0, false, 0, null, 0))
                {
                    MessageBox.Show("Either drawing view selected is wrong or there is a body hidden in the cutlist.");
                    return;
                };
            }

            // check if confirmed cut list is empty
            if(confirmedCl.Count == 0)
            {
                MessageBox.Show("No item in the cut list is included");
                return;
            }
            #endregion

            ViewPallete vp = new ViewPallete(confirmedCl, swApp, swModel, viewModel, configName);
            vp.Show();

            Close();
        }
    }
}
