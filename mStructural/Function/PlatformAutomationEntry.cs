using mStructural.WPF;
using SolidWorks.Interop.sldworks;
using System.Windows;

namespace mStructural.Function
{
    public class PlatformAutomationEntry
    {
        #region Private Variables
        private SldWorks swApp;
        #endregion

        // constructor
        public PlatformAutomationEntry(SldWorks swapp)
        {
            swApp = swapp;
        }

        // main method
        public void Run()
        {
            // check if any model is open
            ModelDoc2 swModel = swApp.IActiveDoc2;
            if (swModel != null)
            {
                MessageBox.Show("Close all open documents before running this function.");
                return;
            }
            else
            {
                PlatformAutomation3 pa = new PlatformAutomation3(swApp);
                pa.Show();
            }
        }
    }
}
