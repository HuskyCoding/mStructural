using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Drawing;

namespace mStructural.Function
{
    public class EtchView
    {
        #region Private Variables
        private SldWorks swApp;
        Macros macros;
        Message msg;
        #endregion

        // constructor
        public EtchView(SldWorks swapp)
        {
            swApp = swapp;
            macros = new Macros(swApp);
            msg = new Message(swApp);
        }

        // main method
        public void Run()
        {
            // get active model doc
            ModelDoc2 swModel = swApp.IActiveDoc2;

            // cast to drawing doc to access to drawing doc method
            DrawingDoc swDraw = swModel as DrawingDoc;

            // use macro to check doc for null and doc type
            if (!macros.checkDocNullandType(swModel, swDocumentTypes_e.swDocDRAWING)) return;
            
            // init selection manager
            SelectionMgr swSelMgr = swModel.SelectionManager;
            
            // if there are selection
            if (swSelMgr.GetSelectedObjectCount2(-1) > 0)
            {
                // only process drawing view object
                for(int i =1; i<swSelMgr.GetSelectedObjectCount2(-1); i++)
                {
                    // check selected object type
                    if (swSelMgr.GetSelectedObjectType3(i, -1) == (int)swSelectType_e.swSelDRAWINGVIEWS)
                    {
                        // cast it to view object
                        View swView = swSelMgr.GetSelectedObject6(i, -1);

                        // get visible component
                        object[] vComps = (object[])swView.GetVisibleComponents();
                        
                        // loop component
                        foreach (object vComp in vComps)
                        {
                            // cast it to component object
                            Component2 swComp = (Component2)vComp;

                            // get all visible edge entity
                            object[] vEdges = (object[])swView.GetVisibleEntities2(swComp, (int)swViewEntityType_e.swViewEntityType_Edge);
                            
                            // loop entity
                            foreach (object vEdge in vEdges)
                            {
                                // cast to entity class
                                Entity swEnt = (Entity)vEdge;

                                // check if the model name contian etch
                                if (swEnt.ModelName.Contains("etch"))
                                {
                                    // select the entity if yes
                                    swEnt.Select4(true, null);
                                }
                            }
                        }
                    }
                }

                // color selected entity red
                Color myColor = Color.Red;
                int winColor = ColorTranslator.ToWin32(myColor);
                swDraw.SetLineColor(winColor);

                // clear selection
                swModel.ClearSelection2(true);
            }
            else
            {
                msg.ErrorMsg("Please select view(s) before using this function.");
                return;
            }

            msg.InfoMsg("Etch View completed.");
        }
    }
}
