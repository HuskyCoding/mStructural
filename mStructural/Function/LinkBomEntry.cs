using mStructural.WPF;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace mStructural.Function
{
    public class LinkBomEntry
    {
        #region Private Variables
        private SldWorks swApp;
        private Macros macros;
        #endregion

        // constructor
        public LinkBomEntry(SldWorks swapp)
        {
            swApp = swapp;
            macros = new Macros(swApp);
        }

        // main method
        public void Run()
        {
            // get active model doc
            ModelDoc2 swModel = swApp.IActiveDoc2;

            // check doc type
            if (!macros.checkDocNullandType(swModel, swDocumentTypes_e.swDocDRAWING)) return;

            LinkBomWPF linkBomWPF = new LinkBomWPF(swModel);
            linkBomWPF.Show();
        }
    }
}
