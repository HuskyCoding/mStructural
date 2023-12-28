using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using mStructural.Classes;
using System.IO;
using System.Collections.Generic;

namespace mStructural.Function
{
    public class DeleteTubeLaserBody
    {
        #region Private Variables
        private SldWorks swApp;
        private ModelDoc2 swModel;
        private Macros macros;
        private Message msg;
        private appsetting appset;
        private string[] bodyNameStrArr = null;
        private List<string> bodyNameStrList = null;
        #endregion

        // constructor
        public DeleteTubeLaserBody(SldWorks swapp)
        {
            swApp = swapp;
            swModel = swApp.IActiveDoc2;
            macros = new Macros(swApp);
            msg = new Message(swApp);
            appset= new appsetting();
        }

        // main method
        public void Run()
        {
            // check doc type
            if (!macros.checkDocNullandType(swModel, swDocumentTypes_e.swDocPART)) return;

            // clear selection first
            swModel.ClearSelection2(true);

            // read string from file
            string txtPath = appset.DeleteTubeLaserBodyLoc;
            string bodyNameString = File.ReadAllText(txtPath);

            // split the string with ;
            bodyNameStrArr = bodyNameString.Split(';');

            // remove any member that is white space or null
            bodyNameStrList = new List<string>();
            foreach(string str in bodyNameStrArr)
            {
                if (!string.IsNullOrEmpty(str))
                {
                    string cleaned = str.Replace("\n", "").Replace("\r", "");
                    bodyNameStrList.Add(cleaned);
                }
            }

            // Traverse feature
            Feature swFeat = swModel.IFirstFeature();
            while( swFeat != null )
            {
                // do the work
                SelectBodies(swFeat);

                // Traverse sub feature
                Feature subFeat = swFeat.IGetFirstSubFeature();
                while(subFeat != null )
                {
                    // do the work
                    SelectBodies(subFeat);
                    subFeat = subFeat.IGetNextSubFeature();
                }
                swFeat = swFeat.IGetNextFeature();
            }

            // check if any body is selected
            SelectionMgr swSelMgr = swModel.ISelectionManager;
            if (swSelMgr.GetSelectedObjectCount2(-1) > 0)
            {
                // delete select bodies
                Feature delBodyFeat = swModel.FeatureManager.InsertDeleteBody2(false);
                DeleteBodyFeatureData delBodyFeatData = (DeleteBodyFeatureData)delBodyFeat.GetDefinition();
                msg.InfoMsg("Feature inserted, deleted " + delBodyFeatData.GetBodiesCount().ToString() +" body(s)");
            }
            else
            {
                msg.ErrorMsg("No body is selected.");
            }

        }

        // method to add select bodies according to list of string
        private void SelectBodies(Feature swfeat)
        {
            // check if it is cut list folder
            if(swfeat.GetTypeName() == "CutListFolder")
            {
                // check if the name contain predefined character
                if (CheckName(swfeat.Name))
                {
                    // get the body folder
                    BodyFolder bodyFolder = (BodyFolder)swfeat.GetSpecificFeature2();

                    // get all bodies
                    object[] vBodies = null;
                    vBodies = (object[])bodyFolder.GetBodies();
                    if (vBodies != null)
                    {
                        foreach(object vBody in vBodies)
                        {
                            // cast to Body2 type
                            Body2 swBody = (Body2)vBody;

                            // select the body
                            swBody.Select2(true, null);
                        }
                    }
                }
            }
        }

        // method to check if the bodies contains the name included in the setting
        private bool CheckName(string name)
        {
            foreach(string s in bodyNameStrList)
            {
                if (name.Contains(s, StringComparison.OrdinalIgnoreCase)) return true;
            }

            return false;
        }
    }
}
