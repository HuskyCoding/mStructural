using mStructural.WPF;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace mStructural.Function
{
    public class PrepDxfEntry
    {
        #region Private Variables
        private SldWorks swApp;
        private Macros macros;
        #endregion

        // constructor
        public PrepDxfEntry(SldWorks swapp)
        {
            swApp = swapp;
            macros = new Macros(swApp);
        }

        // main method
        public void Run()
        {
            // get active document
            ModelDoc2 swModel = swApp.IActiveDoc2;

            // support drawing file only, check if it is drawing file
            if (!macros.checkDocNullandType(swModel, swDocumentTypes_e.swDocDRAWING)) return;

            // show interface
            PrepDxfWPF prepDxfWPF = new PrepDxfWPF(swApp,swModel);
            prepDxfWPF.Show();
        }
    }
}
