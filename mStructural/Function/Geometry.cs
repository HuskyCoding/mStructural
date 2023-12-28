using SolidWorks.Interop.sldworks;
using System.Collections.Generic;

namespace mStructural.Function
{
    public class Geometry
    {
        #region Private Variables
        private SldWorks swApp;
        private MathUtility swMath;
        #endregion

        // constructor
        public Geometry(SldWorks swapp)
        {
            swApp = swapp;
            swMath = swApp.IGetMathUtility();
        }

        public Body2 ExtrudeRegion(SketchRegion region, double depth)
        {
            MathUtility swMath = swApp.IGetMathUtility();
            Body2 swBody = CreateBodyFromRegion(region);
            Modeler swModeler = swApp.IGetModeler();
            object[] vFaces =(object[]) swBody.GetFaces();
            Face2 swFace = (Face2)vFaces[0];
            MathVector swDir = (MathVector)swMath.CreateVector(swFace.Normal);
            Body2 extrudeBody = swModeler.CreateExtrudedBody(swBody,swDir,depth);
            return extrudeBody;
        }

        private Body2 CreateBodyFromRegion(SketchRegion region)
        {
            List<Curve> curveList = GetCurvesFromRegion(region);
            Sketch swSketch = region.Sketch;
            Surface swSurf = CreatePlanarSurfaceFromSketch(swSketch);
            Curve[] swCurves = curveList.ToArray();
            Body2 swBody = (Body2)swSurf.CreateTrimmedSheet5(swCurves, true, 0.00001);
            return swBody;
        }

        private List<Curve> GetCurvesFromRegion(SketchRegion region)
        {
            List<Curve> curveList = new List<Curve>();
            Loop2 swLoop = region.GetFirstLoop();
            while (swLoop != null)
            {
                object[] vLoopEdges = (object[])swLoop.GetEdges();
                foreach(object vLoopEdge in vLoopEdges)
                {
                    Edge swEdge = (Edge)vLoopEdge;
                    Curve swCurve = swEdge.IGetCurve().ICopy();
                    curveList.Add(swCurve);
                }
                swLoop = swLoop.IGetNext();
            }
            return curveList;
        }

        private Surface CreatePlanarSurfaceFromSketch(Sketch sketch)
        {
            double[] dPt = new double[3];
            dPt[0] = 0;
            dPt[1] = 0;
            dPt[2] = 0;

            double[] dVec = new double[3];
            dVec[0] = 0;
            dVec[1] = 0;
            dVec[2] = 1;

            double[] dVecR = new double[3];
            dVecR[0] = 1;
            dVecR[1] = 0;
            dVecR[2] = 0;

            Modeler swModeler = swApp.IGetModeler();
            Surface swSurf = (Surface)swModeler.CreatePlanarSurface2((dPt), (dVec), (dVecR));
            return swSurf;
        }
    }
}
