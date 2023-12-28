using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using System;

namespace mStructural.MacroFeatures
{
    public class EtchFeatureMF : SwComFeature
    {
        // Method when edit feature
        public object Edit(object app, object modelDoc, object feature)
        {
            // create new pmp when edit
            EtchFeaturePMP etchFeatPMP = new EtchFeaturePMP((SldWorks)app, 1, (Feature)feature);
            etchFeatPMP.Show();
            return null;
        }
        
        public object Regenerate(object app, object modelDoc, object feature)
        {
            ModelDoc2 swModel = (ModelDoc2)modelDoc;
            PartDoc swPart = (PartDoc)modelDoc;
            Feature swFeat = (Feature)feature;

            // get text parameter
            MacroFeatureData swFeatData = (MacroFeatureData)swFeat.GetDefinition();
            string Text;
            swFeatData.GetStringByName("Text", out Text);

            // get state parameter
            int State;
            swFeatData.GetIntegerByName("State", out State);

            if(State == 1 || State == 0)
            {
                int counter = 0; // counter for etch naming
                DateTime dateTime = DateTime.Now;
                string DateTimeStr = dateTime.ToString("yyMMddHHmmss");

                // get feature colour
                double[] featColours = null;
                featColours = (double[])swModel.MaterialPropertyValues;
                featColours[0] = 1;
                featColours[1] = 0;
                featColours[2] = 0;
            
                // get selections
                object sels;
                object types;
                object marks;
                swFeatData.GetSelections(out sels, out types, out marks);
            
                if (sels != null)
                {
                    object[] selArr = (object[])sels;
                    foreach (object sel in selArr)
                    {
                        // get selected features
                        Feature etchFeat = (Feature)sel;
                                            
                        // get all faces related to this feature
                        object[] vFaces = (object[])etchFeat.GetFaces();
                        if (vFaces != null)
                        {
                            foreach(object vFace in vFaces)
                            {
                                Face2 swFace = (Face2)vFace;
                        
                                // change selected face colour
                                swFace.MaterialPropertyValues = featColours;
                            
                                // get all edges for this face
                                object[] vEdges = (object[])swFace.GetEdges();
                                foreach(object vEdge in vEdges)
                                {
                                    // cast to edge
                                    Edge swEdge = (Edge)vEdge;

                                    // cast to entity
                                    Entity swEnt = (Entity)swEdge;

                                    // delete existing name
                                    // swPart.DeleteEntityName(swEnt);
                            
                                    // get current entity name
                                    string entName = null;
                                    entName = swPart.GetEntityName(swEnt);
                                                                    
                                    if(entName == "")
                                    {
                                        // set name accordign to counter
                                        bool bRet = swPart.SetEntityName(swEnt, Text + "_"  +DateTimeStr + counter);

                                        // counter increment
                                        counter++;
                                    }
                                }
                            
                            }
                        
                        }
                    }
                }

                swFeatData.SetIntegerByName("State", 2);
            }

            return true;
        }

        public object Security(object app, object modelDoc, object feature)
        {
            return swMacroFeatureSecurityOptions_e.swMacroFeatureSecurityByDefault;
        }
    }
}
