using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;

namespace mStructural.Function
{
    public class DimProfile
    {
        #region Private Members
        SldWorks swApp;
        Message msg;
        Macros macros;
        #endregion

        public DimProfile(SWIntegration swIntegration)
        {
            swApp = swIntegration.SwApp;
            msg = new Message(swApp);
            macros = new Macros(swApp);
        }

        public void Run()
        {
            // get active model doc
            ModelDoc2 swModel = swApp.IActiveDoc2;

            bool bRet = macros.checkDocNullandType(swModel, swDocumentTypes_e.swDocDRAWING); if (!bRet) return; // check

            SelectionMgr swSelMgr = swModel.ISelectionManager;
            List<View> viewList = new List<View>();
            int selectedCount = swSelMgr.GetSelectedObjectCount2(-1);

            if (selectedCount > 0)
            {
                try
                {
                    for(int i =1; i < selectedCount+1; i++)
                    {
                        if(swSelMgr.GetSelectedObjectType3(i,-1) == (int)swSelectType_e.swSelDRAWINGVIEWS)
                        {
                            View swView = (View)swSelMgr.GetSelectedObject6(i, -1);
                            viewList.Add(swView);
                        }
                    }

                    View[] swViews = viewList.ToArray();
                    macros.DimensionView(swViews);
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
