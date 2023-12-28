using mStructural.WPF;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Windows;

namespace mStructural.Function
{
    public class AutoBodyView
    {
        #region Private Variables
        private SldWorks swApp;
        private ModelDoc2 swModel;
        private Macros macros;
        #endregion

        // constructor
        public AutoBodyView(SldWorks swapp)
        {
            // grab sldworks 
            swApp = swapp;

            // instantiate macro class
            macros = new Macros(swApp);
        }

        // main mthod
        public void Run()
        {
            // check if a weldment cutlist is preselected
            swModel = swApp.IActiveDoc2;

            // only support drawing file
            if (!macros.checkDocNullandType(swModel, swDocumentTypes_e.swDocDRAWING)) return;

            // instantiate selection manager
            SelectionMgr swSelMgr = swModel.ISelectionManager;

            // check if a cutlist is selected
            if (swSelMgr.GetSelectedObjectCount2(-1) != 1)
            {
                MessageBox.Show("Please select a cut list, you might selected multiple items or did not select any.");
            }
            else
            {
                // if slection is annotation table
                if (swSelMgr.GetSelectedObjectType3(1, -1) == (int)swSelectType_e.swSelANNOTATIONTABLES)
                {
                    // cast selected object as annotation table
                    TableAnnotation clTable = (TableAnnotation)swSelMgr.GetSelectedObject6(1,-1);

                    // check if the annotation table type is cutlist
                    if(clTable.Type == (int)swTableAnnotationType_e.swTableAnnotation_WeldmentCutList)
                    {
                        // show the user interface
                        AutoBodyViewWPF autoBodyView = new AutoBodyViewWPF(swApp, swModel, clTable);
                        autoBodyView.Show();
                    }
                    else
                    {
                        MessageBox.Show("Selected item is not a weldment cutlist");
                    }
                }
                else
                {
                    MessageBox.Show("Selected item is not an annotation table");
                }
            }
        }
    }
}
