using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace mStructural.MacroFeatures
{
    [Guid("3B5CF0AA-24F3-4A2F-95B5-B1E9EF9421A7")]
    public class ExperimentalMF : SwComFeature
    {
        public object Edit(object app, object modelDoc, object feature)
        {
            ExperimentalPMP pmp = new ExperimentalPMP((SldWorks)app, 1, (Feature)feature);
            pmp.Show();

            return null;
        }

        public object Regenerate(object app, object modelDoc, object feature)
        {
            object functionReturnValue = null;

            SldWorks swApp = (SldWorks)app;
            Feature swFeat = (Feature)feature;
            ModelDoc2 swModel = (ModelDoc2)modelDoc;
            int mark = 1;

            // get parameter
            MacroFeatureData swFeatData = (MacroFeatureData)swFeat.GetDefinition();
            object pNames;
            object pTypes;
            object pValues;
            swFeatData.GetParameters(out pNames, out pTypes, out pValues);
            object[] valueArr = (object[])pValues;
            double depth = double.Parse(valueArr[0].ToString());
            int cutDirection = int.Parse(valueArr[1].ToString());

            // select sketches
            object sels;
            object types;
            object marks;
            swFeatData.GetSelections(out sels, out types, out marks);
            if (sels != null)
            {
                object[] selArr = (object[])sels;
                foreach (object sel in selArr)
                {
                    Feature skFeat = (Feature)sel;
                    skFeat.Select2(true, mark);
                }
            }

            // create preview bodies
            Experimental etchFeature = new Experimental(swApp);
            List<Body2> bodyList = new List<Body2>();
            bodyList = etchFeature.UpdatePreview(swModel, depth, mark, cutDirection);

            // get all bodies
            object[] vbodies = (object[])swFeatData.EditBodies;
            
            // create temp bodies
            List<Body2> tempBodyList = new List<Body2>();
            if (vbodies.Length > 0)
            {
                foreach(object body in vbodies)
                {
                    Body2 tempBody= ((Body2)body).ICopy();
                    tempBodyList.Add(tempBody);
                }
            }

            // combine tool body
            Body2 resToolBody;

            

            List<Body2> outBodyList = new List<Body2>();
            foreach(Body2 body in tempBodyList)
            {
                foreach(Body2 toolBody in bodyList)
                {
                    int err;
                    // check intersection first
                    object[] arrEdges = (object[])body.GetIntersectionEdges(toolBody);
                    if(arrEdges.Length > 0)
                    {
                        object[] outputBodies = (object[])body.Operations2((int)swBodyOperationType_e.SWBODYCUT, toolBody, out err);
                        if (outputBodies != null)
                        {
                            foreach(object outbody in outputBodies)
                            {
                                Body2 outBody = (Body2)outbody;
                                outBodyList.Add(outBody);
                            }
                        }
                    }
                }
            }

            foreach(Body2 body in outBodyList)
            {
                //Body2 swbody = (Body2)body;
                AssignUserIds(body, swFeatData);
            }

            functionReturnValue = outBodyList.ToArray();
            
            return functionReturnValue;
        }

        public object Security(object app, object modelDoc, object feature)
        {
            return swMacroFeatureSecurityOptions_e.swMacroFeatureSecurityByDefault;
        }

        // method to assign user id
        private void AssignUserIds(Body2 body, MacroFeatureData featdata)
        {
            object vFaces = null;
            object vEdges = null;
            object[] faceArr = null;
            object[] edgeArr = null;
            int i = 0;
            int j = 0;

            featdata.GetEntitiesNeedUserId(body, out vFaces, out vEdges);
            faceArr = (object[])vFaces;
            edgeArr = (object[])vEdges;

            if(faceArr != null)
            {
                foreach(object face in faceArr)
                {
                    Face2 swFace = (Face2)face;
                    featdata.SetFaceUserId(swFace, 0, i);
                    i++;
                }
            }

            if(edgeArr != null)
            {
                foreach (object edge in edgeArr)
                {
                    Edge swEdge = (Edge)edge;
                    featdata.SetEdgeUserId(swEdge, 0, j);
                    j++;
                }
            }

        }
    }
}
