using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;

namespace mStructural.Function
{
    public class BalloonView
    {
        #region Privte Variables
        SldWorks swApp;
        Message msg;
        Macros macros;
        #endregion

        public BalloonView(SWIntegration swIntegration)
        {
            swApp = swIntegration.SwApp;
            macros = new Macros(swApp);
            msg = new Message(swApp);
        }

        public void Run()
        {
            List<View> viewList = new List<View> ();
            List<Note> noteList = new List<Note> ();
            ModelDoc2 swModel = swApp.IActiveDoc2;

            bool bRet = macros.checkDocNullandType(swModel, swDocumentTypes_e.swDocDRAWING); if (!bRet) return; // check
            SelectionMgr swSelMgr = swModel.ISelectionManager;
            int selectedCount = swSelMgr.GetSelectedObjectCount2(-1);

            if(selectedCount > 0)
            {
                try
                {
                    for(int i = 1; i< selectedCount + 1; i++)
                    {
                        if(swSelMgr.GetSelectedObjectType3(i,-1) == (int)swSelectType_e.swSelDRAWINGVIEWS)
                        {
                            View swView = (View)swSelMgr.GetSelectedObject6(i, -1);
                            viewList.Add(swView);
                        }
                    }

                    View[] swViews = viewList.ToArray();
                    macros.BalloonCutlist(swViews, out noteList);
                    swModel.ClearSelection2(true);
                }
                catch(Exception ex)
                {
                    msg.ErrorMsg(ex.ToString());
                }
            }
            else
            {
                msg.ErrorMsg("Please select at least a view to proceed.");
                return;
            }
        }
    }
}
