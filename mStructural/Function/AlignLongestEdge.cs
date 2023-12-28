using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace mStructural.Function
{
    public class AlignLongestEdge
    {
        #region Private Variables
        private SldWorks swApp;
        private Macros macros;
        private Message msg;
        #endregion

        // constructor
        public AlignLongestEdge(SldWorks swapp)
        {
            swApp = swapp;
            macros = new Macros(swApp);
            msg = new Message(swApp);
        }

        // Main method
        public void Run()
        {
            // make sure view is selected
            // instantiate model doc 2
            ModelDoc2 swModel = swApp.IActiveDoc2;

            // cast modeldoc2 as drawingdoc
            DrawingDoc swDraw = (DrawingDoc)swModel;

            // check 
            if (!macros.checkDocNullandType(swModel, swDocumentTypes_e.swDocDRAWING)) return;

            // instantiate selection manager
            SelectionMgr swSelMgr = swModel.ISelectionManager;

            // check selection count
            if (swSelMgr.GetSelectedObjectCount2(-1) != 1)
            {
                msg.ErrorMsg("Please select a single view to align.");
                return;
            }

            // check selection type
            if(swSelMgr.GetSelectedObjectType3(1,-1) != (int)swSelectType_e.swSelDRAWINGVIEWS)
            {
                msg.ErrorMsg("Selected object is not a drawing view.");
                return;
            }

            // cast selected object as view object
            View swView = (View)swSelMgr.GetSelectedObject6(1, -1);

            // get the longest edge
            Edge longestEdge = macros.GetLongestEdge(swView);

            // instantiate math object
            MathUtility swMath = swApp.IGetMathUtility();

            // get the angle of the longest edge to horizontal
            double angle = macros.GetEdgeAngle(swMath, swView, longestEdge);

            // rotate view
            swDraw.DrawingViewRotate(angle);
        }
    }
}
