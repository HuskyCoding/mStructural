using SolidWorks.Interop.sldworks;
using System.Collections.Generic;

namespace mStructural.MacroFeatures
{
    public class Experimental
    {
        #region Private Variables
        private SldWorks swApp;
        #endregion

        // constructor
        public Experimental(SldWorks swapp)
        {
            swApp = swapp;
        }

        // Method to update preview
        public List<Body2> UpdatePreview(ModelDoc2 swModel, double depth, int mark, int cutDirection)
        {
            MathUtility swMath = swApp.IGetMathUtility();
            List<Body2> bodyList = new List<Body2>();
            List<Sketch> skList = new List<Sketch>();
            List<Curve> curveList = new List<Curve>();

            // get all selected sketch
            SelectionMgr swSelMgr = swModel.ISelectionManager;
            int selObjCount = swSelMgr.GetSelectedObjectCount2(mark);
            for(int i = 1; i < selObjCount+1; i++)
            {
                if (swSelMgr.GetSelectedObjectType3(i, mark) == 9)
                {
                    object selObj = swSelMgr.GetSelectedObject6(i, mark);
                    Feature swFeat=(Feature)selObj;
                    Sketch swSketch = (Sketch)swFeat.GetSpecificFeature2();
                    skList.Add(swSketch);
                }
            }

            /*
            // get sketch segment for each sketch
            foreach(Sketch sketch in skList)
            {
                curveList = new List<Curve>();
                object[] vContours = (object[])sketch.GetSketchContours();
                foreach(object vContour in vContours)
                {
                    SketchContour skSkCon = (SketchContour)vContour;
                    object[] vEdges = (object[])skSkCon.GetEdges();
                    foreach(object vEdge in vEdges)
                    {
                        Edge swEdge = (Edge)vEdge;
                        CurveParamData swCurveData = swEdge.GetCurveParams3();
                        Curve swCurve = swEdge.IGetCurve();
                        double[] stPt = (double[])swCurveData.StartPoint;
                        double[] enPt = (double[])swCurveData.EndPoint;
                        swCurve = swCurve.CreateTrimmedCurve2(stPt[0], stPt[1], stPt[2], enPt[0], enPt[1], enPt[2]);
                        curveList.Add(swCurve);
                    }
                }

                #region Create Planar surface
                int entType = 0;
                Entity swEnt = (Entity)sketch.GetReferenceEntity(ref entType);

                double[] dPt = new double[3];
                double[] dVec = new double[3];

                if(entType == (int)swSelectType_e.swSelFACES)
                {
                    Face2 swRefFace = (Face2)swEnt;
                    dVec = (double[])swRefFace.Normal;
                }
                else if(entType == (int)swSelectType_e.swSelDATUMPLANES)
                {
                    RefPlane swRefPlane = (RefPlane)swEnt;
                    MathTransform RefPlaneTransform = swRefPlane.Transform;
                    double[] vXfm = (double[])RefPlaneTransform.ArrayData;
                    dVec[0] = vXfm[9];
                    dVec[1] = vXfm[10];
                    dVec[2] = vXfm[11];
                }

                dPt[0] = 0;
                dPt[1] = 0;
                dPt[2] = 0;
                MathPoint swRootPt = (MathPoint)swMath.CreatePoint(dPt);

                dVec[0] = 0;
                dVec[1] = 0;
                dVec[2] = 1;
                MathVector swNormVec = (MathVector)swMath.CreateVector(dVec);

                MathTransform swSkTransform = sketch.ModelToSketchTransform.IInverse();

                swRootPt = swRootPt.IMultiplyTransform(swSkTransform);
                swNormVec = swNormVec.IMultiplyTransform(swSkTransform);

                Modeler swModeler = swApp.IGetModeler();
                Surface swSurf = (Surface)swModeler.CreatePlanarSurface(swRootPt.ArrayData, swNormVec.ArrayData);
                #endregion

                #region Create Body from region
                Body2 swBody = (Body2)swSurf.CreateTrimmedSheet5(curveList.ToArray(), true, 0.00001);

                #endregion

                #region Extrude region
                if (swBody != null)
                {
                    object[] vFaces = (object[])swBody.GetFaces();
                    Face2 swFace = (Face2)vFaces[0];
                    MathVector swDir = (MathVector)swMath.CreateVector(swFace.Normal);
                    Body2 regionBody = swModeler.CreateExtrudedBody(swBody, swDir, depth);
                    if (regionBody != null)
                    {
                        bodyList.Add(regionBody);
                    }
                }
                #endregion
            }
            */
            
            foreach (Sketch swSketch in skList)
            {
                #region Get Curves From Region
                object[] vSkRegions = (object[])swSketch.GetSketchRegions();
                int i;
                for(i = vSkRegions.GetLowerBound(0); i<=vSkRegions.GetUpperBound(0); i++)
                {
                    SketchRegion swSkRegion = (SketchRegion)vSkRegions[i];
                    if (swSkRegion != null)
                    {
                        curveList = new List<Curve>();
                        Loop2 swLoop = swSkRegion.GetFirstLoop();
                        while (swLoop != null)
                        {
                            object[] vLoopEdges = (object[])swLoop.GetEdges();
                            int k;
                            for (k=vLoopEdges.GetLowerBound(0);k<=vLoopEdges.GetUpperBound(0); k++)
                            {
                                Edge swLoopEdge = (Edge)vLoopEdges[k];
                                Curve swCurve = swLoopEdge.IGetCurve().ICopy();
                                curveList.Add(swCurve);
                            }
                            swLoop = swLoop.IGetNext();
                        }

                        #region Create Planar surface
                        double[] dPt = new double[3];
                        double[] dVec = new double[3];

                        dPt[0] = 0;
                        dPt[1] = 0;
                        dPt[2] = 0;
                        MathPoint swRootPt = (MathPoint)swMath.CreatePoint(dPt);

                        dVec[0] = 0;
                        dVec[1] = 0;
                        dVec[2] = cutDirection;
                        MathVector swNormVec = (MathVector)swMath.CreateVector(dVec);

                        dVec[0] = 1;
                        dVec[1] = 0;
                        dVec[2] = 0;
                        MathVector swRefVec = (MathVector)swMath.CreateVector(dVec);

                        MathTransform swSkTransform = swSketch.ModelToSketchTransform.IInverse();

                        swRootPt = swRootPt.IMultiplyTransform(swSkTransform);
                        swNormVec = swNormVec.IMultiplyTransform(swSkTransform);
                        swRefVec = swRefVec.IMultiplyTransform(swSkTransform);

                        Modeler swModeler = swApp.IGetModeler();
                        Surface swSurf = (Surface)swModeler.CreatePlanarSurface2(swRootPt.ArrayData, swNormVec.ArrayData, swRefVec.ArrayData);
                        #endregion

                        #region Extrude region
                        // Create Body from region
                        Body2 swBody = (Body2)swSurf.CreateTrimmedSheet5(curveList.ToArray(), true, 0.00001);
                        if(swBody != null)
                        {
                            object[] vFaces = (object[])swBody.GetFaces();
                            Face2 swFace = (Face2)vFaces[0];
                            MathVector swDir = (MathVector)swMath.CreateVector(swFace.Normal);
                            Body2 regionBody = swModeler.CreateExtrudedBody(swBody, swDir, depth);
                            if(regionBody != null)
                            {
                                bodyList.Add(regionBody);
                            }
                        }
                        #endregion
                    }
                }
                #endregion
            }
            
            return bodyList;
        }
    }
}
