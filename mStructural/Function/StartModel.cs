using mStructural.WPF;
using SolidWorks.Interop.sldworks;

namespace mStructural.Function
{
    public class StartModel
    {
        #region Private Variables
        private SldWorks swApp;
        private Macros macros;
        private appsetting appset;
        #endregion

        // Constructor
        public StartModel(SWIntegration swIntegration)
        {
            swApp = swIntegration.SwApp;
            macros = new Macros(swApp);
            appset = new appsetting();
        }

        // Main method
        public void Run()
        {
            // get pdm root folder
            string pdmRootFolder = appset.StartModelPath;

            // show form
            StartModelMenu menu = new StartModelMenu(swApp);
            menu.ShowDialog();
        }
    }
}
