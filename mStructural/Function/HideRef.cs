using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Collections.Generic;

namespace mStructural.Function
{
    public class HideRef
    {
        #region Private Variables
        private ModelDoc2 swModel;
        private List<string> hideTypeList;
        private Message msg;
        #endregion

        public void Run(SldWorks swApp, List<string> hidetypelist)
        {
            hideTypeList = hidetypelist;

            // instantiate message macro
            msg = new Message(swApp);

            // get active document
            swModel = swApp.IActiveDoc2;

            // check document type
            if(swModel != null)
            {
                switch (swModel.GetType())
                {
                    case (int)swDocumentTypes_e.swDocPART:
                        {
                            HideRefPart();
                            break;
                        }
                    case (int)swDocumentTypes_e.swDocASSEMBLY:
                        {
                            HideRefAsm();
                            break;
                        }
                    case (int)swDocumentTypes_e.swDocDRAWING:
                        {
                            HideRefDraw();
                            break;
                        }
                }
            }
        }

        // method to hide ref for part
        private void HideRefPart()
        {
            TraverseModelFeatures(swModel);

            msg.InfoMsg("Completed!");
        }

        // method to hide ref for assembly
        private void HideRefAsm()
        {
            Configuration swConf = swModel.GetActiveConfiguration();
            Component2 swRootComp = swConf.GetRootComponent();
            TraverseModelFeatures(swModel);
            TraverseComponent(swRootComp);

            msg.InfoMsg("Completed!");
        }

        // method to hide ref for drawing
        private void HideRefDraw()
        {
            // cast to drawing doc for method
            DrawingDoc swDraw = swModel as DrawingDoc;
            
            // get all sheet name
            string[] vSheetNames = swDraw.GetSheetNames();
            
            // loop sheet name
            foreach(string vSheetName in vSheetNames)
            {
                // activate sheet
                bool bRet = swDraw.ActivateSheet(vSheetName);
                
                // if successfully activated
                if (bRet)
                {
                    // get the first view from the view
                    View swView = swDraw.GetFirstView();

                    // loop for view
                    while(swView != null)
                    {
                        // get the drawing component
                        DrawingComponent drawComp = swView.RootDrawingComponent;
                        
                        // if there is component
                        if( drawComp != null)
                        {
                            // traverse for feature, then traverse for child component
                            TraverseDrawingComponent(drawComp, swView);
                        }

                        // get next view
                        swView = swView.GetNextView();
                    } 
                }
            }

            msg.InfoMsg("Completed!");
        }

        // method to traverse model for feature
        private void TraverseModelFeatures(ModelDoc2 swmodel)
        {
            Feature swFeat = swmodel.FirstFeature();
            TraverseFeatureFeatures(swFeat);
        }

        // method to traverse for features
        public void TraverseFeatureFeatures(Feature swFeat)
        {
            Feature swSubFeat;
            Feature swSubSubFeat;
            Feature swSubSubSubFeat;
            List<Feature> featList = new List<Feature>();

            while ((swFeat != null))
            {
                if(TestFeatureType(swFeat)) featList.Add(swFeat);
                swSubFeat = (Feature)swFeat.GetFirstSubFeature();
                while ((swSubFeat != null))
                {
                    if (TestFeatureType(swSubFeat)) featList.Add(swSubFeat);
                    swSubSubFeat = (Feature)swSubFeat.GetFirstSubFeature();
                    while ((swSubSubFeat != null))
                    {
                        if (!TestFeatureType(swSubSubFeat)) featList.Add(swSubSubFeat);
                        swSubSubSubFeat = (Feature)swSubSubFeat.GetFirstSubFeature();
                        while ((swSubSubSubFeat != null))
                        {
                            if(!TestFeatureType(swSubSubSubFeat)) featList.Add(swSubSubSubFeat);
                            swSubSubSubFeat = (Feature)swSubSubSubFeat.GetNextSubFeature();
                        }
                        swSubSubFeat = (Feature)swSubSubFeat.GetNextSubFeature();
                    }
                    swSubFeat = (Feature)swSubFeat.GetNextSubFeature();
                }
                swFeat = (Feature)swFeat.GetNextFeature();
            }

            BlankFeatureInList(featList);
        }

        // method to traverse for component
        public void TraverseComponent(Component2 swComp)
        {
            object[] vChildComp;
            Component2 swChildComp;

            vChildComp = (object[])swComp.GetChildren();
            for (int i = 0; i < vChildComp.Length; i++)
            {
                swChildComp = (Component2)vChildComp[i];

                TraverseComponentFeatures(swChildComp);
                TraverseComponent(swChildComp);
            }
        }

        // method to traverse for component feature
        public void TraverseComponentFeatures(Component2 swComp)
        {
            Feature swFeat;

            swFeat = swComp.FirstFeature();
            TraverseFeatureFeatures(swFeat);
        }

        // method to test feature for type
        private bool TestFeatureType(Feature swFeat)
        {
            bool result = false;
            string featType = swFeat.GetTypeName2();
            if (hideTypeList.Contains(featType))
            {
                if(swFeat.Visible == (int)swVisibilityState_e.swVisibilityStateShown || swFeat.Visible == (int)swVisibilityState_e.swVisibilityStateUnknown)
                {
                    return true;
                }
            }
            return result;
        }

        // method to blank feature
        private void BlankFeatureInList(List<Feature> features)
        {
            foreach(Feature feature in features)
            {
                feature.Select2(true, 0);
            }
            swModel.BlankRefGeom();
            swModel.BlankSketch();

            swModel.ClearSelection2(true);
        }

        // method to traverse for drawing component
        private void TraverseDrawingComponent(DrawingComponent drawComp, View swView)
        {
            TraverseDrawingComponentFeature(drawComp, swView);
            if (drawComp.GetChildrenCount() > 0)
            {
                object[] vChilds  = drawComp.GetChildren();
                foreach(object vChild in vChilds)
                {
                    DrawingComponent drawChildComp = vChild as DrawingComponent;
                    TraverseDrawingComponent(drawChildComp, swView);
                }
            }
        }

        // method to traverse for drawing component feature
        private void TraverseDrawingComponentFeature(DrawingComponent drawComp, View swView)
        {
            Component2 swComp = drawComp.Component;
            Feature swFeat = swComp.FirstFeature();
            while ((swFeat != null))
            {
                SelectFeatureOnDrawing(drawComp, swFeat, swView);
                Feature swSubFeat = (Feature)swFeat.GetFirstSubFeature();
                while ((swSubFeat != null))
                {
                    SelectFeatureOnDrawing(drawComp, swSubFeat, swView);
                    Feature  swSubSubFeat = (Feature)swSubFeat.GetFirstSubFeature();
                    while ((swSubSubFeat != null))
                    {
                        SelectFeatureOnDrawing(drawComp, swSubSubFeat, swView);
                        Feature swSubSubSubFeat = (Feature)swSubSubFeat.GetFirstSubFeature();
                        while ((swSubSubSubFeat != null))
                        {
                            SelectFeatureOnDrawing(drawComp, swSubSubSubFeat, swView);
                            swSubSubSubFeat = (Feature)swSubSubSubFeat.GetNextSubFeature();
                        }
                        swSubSubFeat = (Feature)swSubSubFeat.GetNextSubFeature();
                    }
                    swSubFeat = (Feature)swSubFeat.GetNextSubFeature();
                }
                swFeat = (Feature)swFeat.GetNextFeature();
            }

            swModel.BlankRefGeom();
            swModel.BlankSketch();

            swModel.ClearSelection2(true);
        }

        // method to select feature on drawing
        private void SelectFeatureOnDrawing(DrawingComponent drawComp, Feature swFeat, View swView)
        {
            if(TestFeatureType(swFeat))
            {
                string featFullName = getFeatureFullName(swView.Name, drawComp.Name, swFeat.Name);
                
                string selectionString = "";
                switch (swFeat.GetTypeName2())
                {
                    case "OriginProfileFeature":
                        selectionString = "EXTSKETCHPOINT";
                        featFullName = "Point1@" + featFullName;
                        break;
                    case "RefPlane":
                        selectionString = "PLANE";
                        break;
                    case "RefAxis":
                        selectionString = "AXIS";
                        break;
                    case "RefPoint":
                        selectionString = "DATUMPOINT";
                        break;
                    case "CoordSys":
                        selectionString = "COORDSYS";
                        break;
                    case "ProfileFeature":
                        selectionString = "SKETCH";
                        break;
                    case "3DProfileFeature":
                        selectionString = "SKETCH";
                        break;
                    case "3DSplineCurve":
                        selectionString = "REFERENCECURVES";
                        break;
                    case "CompositeCurve":
                        selectionString = "REFCURVE";
                        break;
                    case "Helix":
                        selectionString = "HELIX";
                        break;
                }
                swModel.Extension.SelectByID2(featFullName, selectionString, 0, 0, 0, true, 0, null, 0);
            }
        }

        // method to get full name of the feature for selection
        private string getFeatureFullName(string viewName, string drawCompName, string featName)
        {
            string fullFeatName = "";

            string[] vStr = drawCompName.Split('/');
            string tempStr = featName + "@" + vStr[0] + "@" + viewName;
            for(int i = 0; i < vStr.Length - 1; i++)
            {
                string strParentName = vStr[i].Substring(0, vStr[i].LastIndexOf("-") );
                tempStr = tempStr + "/" + vStr[i + 1] + "@" + strParentName;
            }

            fullFeatName = tempStr;
            return fullFeatName;
        }
    }
}
