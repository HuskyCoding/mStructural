using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Drawing;

namespace mStructural.Function
{
    public class ColourView
    {
        #region Private Variables
        private SldWorks swApp;
        private Macros macros;
        private Message msg;
        #endregion

        // constructor
        public ColourView(SldWorks swapp)
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

            // cast to drawing doc
            DrawingDoc swDraw = swModel as DrawingDoc;

            // check document 
            if (!macros.checkDocNullandType(swModel, swDocumentTypes_e.swDocDRAWING)) return;

            // instantiate selection manager
            SelectionMgr swSelMgr = swModel.SelectionManager;

            // check if there is any selection
            if (swSelMgr.GetSelectedObjectCount2(-1)>0)
            {
                // loop through each selection
                for(int i = 1; i<swSelMgr.GetSelectedObjectCount(); i++)
                {
                    // only process drawing view
                    if(swSelMgr.GetSelectedObjectType3(i,-1) == (int)swSelectType_e.swSelDRAWINGVIEWS)
                    {
                        // cast selected object to view object
                        View swView = swSelMgr.GetSelectedObject6(i,-1);

                        // get all visible component in the selected view
                        object[] vComps = swView.GetVisibleComponents();

                        // loop components for edge
                        foreach(object vComp in vComps)
                        {
                            // cast the object to component2 object
                            Component2 swComp = vComp as Component2;

                            // get all visible edge from the component
                            object[] vEdges = swView.GetVisibleEntities2(swComp, (int)swViewEntityType_e.swViewEntityType_Edge);

                            // loop edges to select entity
                            foreach(object vEdge in vEdges)
                            {
                                // cast object to entity for selection
                                Entity swEnt = vEdge as Entity;

                                // select entity and append to current selection 
                                swEnt.Select4(true, null);
                            }
                        }
                    }
                }
            }

            // Colour selected edges to black
            Color myColor = Color.Black;
            int winColor = ColorTranslator.ToWin32(myColor);
            swDraw.SetLineColor(winColor);

            // clear selection
            swModel.ClearSelection2(true);

            msg.InfoMsg("Completed!");
        }
    }
}
