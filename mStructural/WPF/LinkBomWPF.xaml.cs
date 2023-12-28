using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Windows;

namespace mStructural.WPF
{
    /// <summary>
    /// Interaction logic for LinkBomWPF.xaml
    /// </summary>
    public partial class LinkBomWPF : Window
    {
        #region Private Variables
        private ModelDoc2 swModel;
        #endregion

        // constructor
        public LinkBomWPF(ModelDoc2 swmodel)
        {
            InitializeComponent();
            swModel = swmodel;
            getAllBom();
        }

        private void LinkBtn_Click(object sender, RoutedEventArgs e)
        {
            // Instantiate selection manager
            SelectionMgr swSelMgr = swModel.ISelectionManager;

            // check if any item selected
            int selCount = swSelMgr.GetSelectedObjectCount2(-1);
            if(selCount > 0)
            {
                for(int i=1; i<selCount+1; i++)
                {
                    // check selected object type, only support drawing view
                    if(swSelMgr.GetSelectedObjectType2(i) == (int)swSelectType_e.swSelDRAWINGVIEWS)
                    {
                        // cast selected object as View
                        View swView = (View)swSelMgr.GetSelectedObject6(i, -1);

                        // set the balloon text to link to selected bom in combo box
                        swView.SetKeepLinkedToBOM(true, bomCb.Text);
                    }
                }

                MessageBox.Show("Link BOM completed!");
            }
            else
            {
                MessageBox.Show("No item selected.");
            }
        }

        // method to get all bom
        private void getAllBom()
        {
            // get first feature
            Feature swFeat = swModel.IFirstFeature();

            // traverse feature
            while(swFeat != null)
            {
                // add to combobox if the feature type is either bom feature or weldment table feature
                if(swFeat.GetTypeName2()=="BomFeat" || swFeat.GetTypeName2() == "WeldmentTableFeat")
                {
                    bomCb.Items.Add(swFeat.Name);
                }
                swFeat = swFeat.IGetNextFeature();
            }
        }
    }
}
