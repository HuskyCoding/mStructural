using mStructural.WPF;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace mStructural.Function
{
    public class ExportDxfEntry
    {
        #region Private Variables
        private SldWorks swApp;
        private Macros macros;
        #endregion

        // Constructor
        public ExportDxfEntry(SldWorks swapp)
        {
            swApp = swapp;
            macros = new Macros(swApp);
        }

        // Main Method
        public void Run()
        {
            // get active document
            ModelDoc2 swModel = swApp.IActiveDoc2;

            // support drawing file only, check if it is drawing file
            if (!macros.checkDocNullandType(swModel, swDocumentTypes_e.swDocDRAWING)) return;

            // show interface
            ExportDxfWPF exportDxfWPF = new ExportDxfWPF(swApp, swModel);
            exportDxfWPF.Show();
        }
    }
}
