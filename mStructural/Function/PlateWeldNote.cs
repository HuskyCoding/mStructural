using SolidWorks.Interop.sldworks;

namespace mStructural.Function
{
    public class PlateWeldNote
    {
        #region Private Variables
        private SldWorks swApp;
        #endregion

        // constructor
        public PlateWeldNote(SldWorks swapp)
        {
            swApp = swapp;
        }

        // Main method
        public void Run()
        {
            // get active doc
            ModelDoc2 swModel = swApp.IActiveDoc2;

            // Get model extension
            ModelDocExtension swModelExt = swModel.Extension;

            // get first feature
            Feature swFeat = swModel.IFirstFeature();

            // loop feature
            while (swFeat != null)
            {
                // if it is cutlist folder
                if(swFeat.GetTypeName2() == "CutListFolder")
                {
                    bool bRet;

                    // cast the feature specific to body folder
                    BodyFolder swBodyFolder = (BodyFolder)swFeat.GetSpecificFeature2();

                    // update cutlist first
                    swBodyFolder.UpdateCutList();

                    // condition
                    if (swFeat.Name.ToLower().Contains("sheet"))
                    {
                        // if contain sheet
                        bRet = swModelExt.SelectByID2(swFeat.Name, "SUBWELDFOLDER", 0, 0, 0, false, 0, null, 0);

                        // create bounding box
                        swModelExt.Create3DBoundingBox();

                        // get custom prop, width, length and thickness
                        string[] strValue = new string[6];
                        swFeat.CustomPropertyManager.Get4("Bounding Box Width", false, out strValue[0], out strValue[1]);
                        swFeat.CustomPropertyManager.Get4("Bounding Box Length", false, out strValue[2], out strValue[3]);
                        swFeat.CustomPropertyManager.Get4("Sheet Metal Thickness", false, out strValue[4], out strValue[5]);
                        
                        // get custom prop mgr
                        CustomPropertyManager swCusPropMgr = swFeat.CustomPropertyManager;

                        // add description
                        swCusPropMgr.Add3("Description", 30, "PLATE, " + strValue[4] + " x " + strValue[0] + " x " + strValue[2], 1);
                    }
                    else if (swFeat.Name.ToLower().Contains("cut-list-item"))
                    {
                        // select the folder
                        bRet = swModelExt.SelectByID2(swFeat.Name, "SUBWELDFOLDER", 0, 0, 0, false, 0, null, 0);

                        // create bounding box
                        swModelExt.Create3DBoundingBox();
                    }
                }

                // get next feature
                swFeat = swFeat.IGetNextFeature();
            }
        }

    }
}
