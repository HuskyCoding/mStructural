using mStructural.WPF;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace mStructural.Function
{
    public class ExportModel
    {
        #region Private Variables
        private SldWorks swApp;
        private Macros macros;
        #endregion

        // Constructor
        public ExportModel(SldWorks swapp)
        {
            swApp = swapp;
            macros = new Macros(swApp);
        }

        // Main Method
        public void Run()
        {
            ModelDoc2 swModel = swApp.IActiveDoc2;
            if (!macros.checkDocNullandType(swModel, swDocumentTypes_e.swDocPART, swDocumentTypes_e.swDocASSEMBLY)) return;
            ExportModelWPF wpf = new ExportModelWPF(swApp, swModel);
            wpf.Show();
        }
    }
}
