using mStructural.WPF;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace mStructural.Function
{
    public class NamingProject
    {
        #region Private Variables
        private SldWorks swApp;
        private Macros macros;
        #endregion

        // constructor
        public NamingProject(SldWorks swapp)
        {
            swApp = swapp;
            macros = new Macros(swApp);
        }

        // main method
        public void Run()
        {
            ModelDoc2 swModel = swApp.IActiveDoc2;
            if (!macros.checkDocNullandType(swModel, swDocumentTypes_e.swDocASSEMBLY)) return;
            NamingProjectWPF namingProjectWPF = new NamingProjectWPF(swModel);
            namingProjectWPF.Show();
        }
    }
}
