using mStructural.WPF;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace mStructural.Function
{
    public class ExportPdfEntry
    {
        #region Private Variables
        private SldWorks swApp;
        private Macros macros;
        #endregion

        // constructor
        public ExportPdfEntry(SldWorks swapp)
        {
            swApp = swapp;
            macros = new Macros(swApp);
        }

        // main method
        public void Run()
        {
            ModelDoc2 swModel = swApp.IActiveDoc2;
            if (!macros.checkDocNullandType(swModel, swDocumentTypes_e.swDocDRAWING)) return;
            ExportPdfWPF exportPdfWPF = new ExportPdfWPF(swApp, swModel);
            exportPdfWPF.Show();
        }
    }
}
