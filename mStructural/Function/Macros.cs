using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Linq;
using static mStructural.Classes.Structs;

namespace mStructural.Function
{
    public class Macros
    {
        #region Private variables
        SldWorks swApp;
        Message msg;
        #endregion

        // Contructor
        public Macros(SldWorks swapp)
        {
            swApp = swapp;
            msg = new Message(swApp);
        }

        // Method to add balloon
        public void BalloonCutlist(View[] swViews, out List<Note> swNotes)
        {
            ModelDoc2 swModel = swApp.IActiveDoc2;
            MathUtility swMath = swApp.IGetMathUtility();
            List<Note> swnotes = new List<Note>(); 

            // check if any document is opened
            bool bRet = checkDocNullandType(swModel, swDocumentTypes_e.swDocDRAWING);
            if (!bRet)
            {
                swNotes = swnotes;
            }

            foreach(View swView in swViews)
            {
                object[] vComps = (object[])swView.GetVisibleComponents();
                Component2 swComp = (Component2)vComps[0];
                object[] vVerts = (object[])swView.GetVisibleEntities2(swComp, (int)swViewEntityType_e.swViewEntityType_Vertex);
                MathTransform modelToViewXF = swView.ModelToViewTransform;
                BalloonOptions swBalOp = swModel.Extension.CreateBalloonOptions();
                swBalOp.ShowQuantity = true;
                swBalOp.QuantityPlacement = (int)swBalloonQuantityPlacement_e.swBalloonQuantityPlacement_Right;

                if (swView.GetVisibleEntityCount2(swComp, (int)swViewEntityType_e.swViewEntityType_Vertex) > 0)
                {
                    MathTransform drawToSkXF = DrawToViewTransform(swMath, swView).IInverse();

                    // get view outline as starting
                    double[] outline = new double[4];
                    outline = (double[])swView.GetOutline();
                    double[] topRightPt = new double[3];
                    topRightPt[0] = outline[0];
                    topRightPt[0] = outline[1];
                    double[] topRightPtRounded = RoundDoubleArray(topRightPt, 6);

                    foreach(object vVert in vVerts)
                    {
                        Vertex swVert = (Vertex)vVert;
                        double[] dPt = new double[3];
                        dPt = (double[])swVert.GetPoint();
                        MathPoint swPt = (MathPoint)swMath.CreatePoint(dPt);
                        swPt = swPt.IMultiplyTransform(drawToSkXF);
                        swPt = swPt.IMultiplyTransform(modelToViewXF);
                        dPt = (double[])swPt.ArrayData;
                        double[] dPtRounded = RoundDoubleArray(dPt,6);

                        if (dPtRounded[1] > topRightPtRounded[1])
                        {
                            topRightPt[0] = dPt[0];
                            topRightPt[1] = dPt[1];
                        }
                        else if (dPtRounded[1] == topRightPtRounded[1])
                        {
                            if (dPtRounded[0] > topRightPtRounded[0])
                            {
                                topRightPt[0] = dPt[0];
                                topRightPt[1] = dPt[1];
                            }
                        }
                    }

                    bRet = swModel.Extension.SelectByID2("", "VERTEX", topRightPt[0], topRightPt[1], 0, false, 0, null, 0);
                    Note swNote = swModel.Extension.InsertBOMBalloon2(swBalOp);
                    if (swNote != null)
                    {
                        Annotation swAnn = swNote.IGetAnnotation();
                        bRet = swAnn.SetLeaderAttachmentPointAtIndex(1, topRightPt[0], topRightPt[1], 0);
                        bRet = swAnn.SetPosition2(topRightPt[0] + 0.007, topRightPt[1] + 0.01, 0);
                        swnotes.Add(swNote);
                    }

                    swModel.ClearSelection2(true);
                }
                else
                {
                    object[] vEdges = (object[])swView.GetVisibleEntities2(swComp, (int)swViewEntityType_e.swViewEntityType_Edge);
                    double[] outerCircleParam = new double[7];
                    Entity outerCircleEnt = default(Entity);
                    double largestR = 0;
                    foreach(object vEdge in vEdges)
                    {
                        Edge swEdge = (Edge)vEdge;
                        Curve swCurve = swEdge.IGetCurve();
                        double[] circleParam = new double[7];
                        circleParam = (double[])swCurve.CircleParams;
                        if (circleParam[6] > largestR)
                        {
                            largestR = circleParam[6];
                            outerCircleParam = circleParam;
                            outerCircleEnt = (Entity)swEdge;
                        }
                    }

                    double cirEdgeOffset = Math.Sqrt(Math.Pow(outerCircleParam[6], 2) / 2);
                    double[] dCenter = new double[3];
                    dCenter[0] = outerCircleParam[0];
                    dCenter[1] = outerCircleParam[1];
                    dCenter[2] = outerCircleParam[2];
                    MathPoint swPt = (MathPoint)swMath.CreatePoint(dCenter);
                    swPt = swPt.IMultiplyTransform(modelToViewXF);
                    double[] dPt = new double[3];
                    dPt = (double[])swPt.ArrayData;
                    bRet = outerCircleEnt.Select4(false, null);
                    SelectionMgr swSelMgr = swModel.ISelectionManager;
                    bRet = swSelMgr.SetSelectionPoint2(1, -1, dPt[0] + cirEdgeOffset, dPt[1] + cirEdgeOffset, 0);
                    Note swNote = swModel.Extension.InsertBOMBalloon2(swBalOp);
                    Annotation swAnn = swNote.IGetAnnotation();
                    bRet = swAnn.SetLeaderAttachmentPointAtIndex(1, dPt[0] + cirEdgeOffset, dPt[1] + cirEdgeOffset, 0);
                    bRet = swAnn.SetPosition2(dPt[0] + cirEdgeOffset + 0.007, dPt[1] + cirEdgeOffset + 0.01, 0);
                    bRet = swAnn.SetLeaderAttachmentPointAtIndex(1, dPt[0] + cirEdgeOffset, dPt[1] + cirEdgeOffset, 0);

                    swnotes.Add(swNote);
                    swModel.ClearSelection2(true);
                }
            }

            swNotes = swnotes;
        }

        // Method to get drawing to view transform
        public MathTransform DrawToViewTransform(MathUtility swMath, View swView)
        {
            double[] transformData = new double[16];
            double[] viewPos = new double[16];
            viewPos = (double[])swView.Position;
            transformData[0] = 1;
            transformData[1] = 0;
            transformData[2] = 0;
            transformData[3] = 0;
            transformData[4] = 1;
            transformData[5] = 0;
            transformData[6] = 0;
            transformData[7] = 0;
            transformData[8] = 1;
            transformData[9] = viewPos[0];
            transformData[10] = viewPos[1];
            transformData[11] = 0;
            transformData[12] = 1;
            transformData[13] = 0;
            transformData[14] = 0;
            transformData[15] = 0;
            MathTransform drawToViewTransform = (MathTransform)swMath.CreateTransform(transformData);
            return drawToViewTransform;
        }

        // Method to round double array
        public double[] RoundDoubleArray(double[] da, int roundingInt)
        {
            double[] Da;
            Da = da;
            for (int i = 0; i < Da.Length; i++)
            {
                Da[i] = Math.Round(Da[i], roundingInt);
            }
            return Da;
        }

        // Method to get the longest edge
        public Edge GetLongestEdge(View view)
        {
            object[] edgeObjs = null;
            object[] compObjs = null;
            Component2 swComp = default(Component2);
            Edge longestedge = default(Edge);
            double longestEdgeLength = 0;

            compObjs = (object[])view.GetVisibleComponents();
            swComp = (Component2)compObjs[0];

            edgeObjs = (object[])view.GetVisibleEntities2(swComp, (int)swViewEntityType_e.swViewEntityType_Edge);

            // loop to find the longest edge
            foreach (object edgeobj in edgeObjs)
            {
                Edge curEdge = default(Edge);
                double curEdgeLength;

                curEdge = (Edge)edgeobj;
                curEdgeLength = GetEdgeData(curEdge).length;

                if (curEdgeLength > longestEdgeLength)
                {
                    longestedge = curEdge;
                    longestEdgeLength = curEdgeLength;
                }
            }
            return longestedge;
        }

        // Method to get the angle of edge
        public double GetEdgeAngle(MathUtility swMath, View view, Edge edge)
        {
            EdgeData LongestEdgeData = new EdgeData();
            MathPoint swStartPt = default(MathPoint);
            MathPoint swEndPt = default(MathPoint);
            MathPoint viewStartPt = default(MathPoint);
            MathPoint viewEndPt = default(MathPoint);
            MathTransform viewxform = default(MathTransform);
            double diff;
            // double theta;
            double[] dViewStartPt = null;
            double[] dViewEndPt = null;

            // calculate angle of the longest edge to horizontal
            LongestEdgeData = GetEdgeData(edge);

            // Get transform for this view
            viewxform = view.ModelToViewTransform;

            // Create math point and transform it
            swStartPt = (MathPoint)swMath.CreatePoint(LongestEdgeData.start);
            swEndPt = (MathPoint)swMath.CreatePoint(LongestEdgeData.end);
            viewStartPt = (MathPoint)swStartPt.MultiplyTransform(viewxform);
            viewEndPt = (MathPoint)swEndPt.MultiplyTransform(viewxform);
            dViewStartPt = ((double[])viewStartPt.ArrayData);
            dViewEndPt = ((double[])viewEndPt.ArrayData);

            double deltaX = dViewEndPt[0] - dViewStartPt[0];
            double deltaY = dViewStartPt[1] - dViewEndPt[1];

            if (deltaX == 0)
            {
                diff = Math.PI/2;
            }
            else if (deltaY == 0)
            {
                diff = 0;
            }
            else
            {
                diff = Math.Atan2(deltaY , deltaX);
            }
            return diff;
        }

        // Method to get the data for the edge
        public EdgeData GetEdgeData(Edge edge)
        {
            EdgeData edgeData = new EdgeData();
            edgeData.start = new double[3];
            edgeData.end = new double[3];
            Curve swCurve = default(Curve);
            CurveParamData curveData = default(CurveParamData);
            double[] startpt = new double[3];
            double[] endpt = new double[3];

            swCurve = (Curve)edge.GetCurve();
            curveData = edge.GetCurveParams3();
            edgeData.length = swCurve.GetLength3(curveData.UMinValue, curveData.UMaxValue);
            startpt = (double[])curveData.StartPoint;
            endpt = (double[])curveData.EndPoint;
            edgeData.start[0] = startpt[0];
            edgeData.start[1] = startpt[1];
            edgeData.start[2] = startpt[2];
            edgeData.end[0] = endpt[0];
            edgeData.end[1] = endpt[1];
            edgeData.end[2] = endpt[2];

            return edgeData;
        }

        // Method to dimension view
        public void DimensionView(View[] swViews)
        {
            // v1.8.0 move preferences to outside of macro so that if it failed, the user setting still preserve
            ModelDoc2 swModel = swApp.IActiveDoc2;
            MathUtility swMath = swApp.IGetMathUtility();
            //bool inputDimDefVal = swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate);
            //bool dimSnappingVal = swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference);
            //swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);
            //swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, false);

            // check if any document is opened
            bool bRet = checkDocNullandType(swModel, swDocumentTypes_e.swDocDRAWING); if (!bRet) return;

            DrawingDoc swDraw = (DrawingDoc)swModel;

            foreach(View swView in swViews)
            {
                object[] vComps = (object[])swView.GetVisibleComponents();
                Component2 swComp = (Component2)vComps[0];
                object[] vVerts = (object[])swView.GetVisibleEntities2(swComp, (int)swViewEntityType_e.swViewEntityType_Vertex);
                MathTransform drawToSkXF = DrawToViewTransform(swMath, swView).IInverse();
                MathTransform modelToViewXF = swView.ModelToViewTransform;

                double[] outline = new double[4];
                outline = (double[])swView.GetOutline();

                swModel.ViewZoomTo2(outline[0], outline[1], 0, outline[2], outline[3], 0);

                double outlineMidX = (outline[0] + outline[2]) / 2;
                double outlineMidY = (outline[1] + outline[3]) / 2;

                double[] topLeftV = { outline[2], outline[1], 0 };
                double[] topRightV = { outline[0], outline[1], 0 };
                double[] btmLeftV = { outline[2], outline[3], 0 };
                double[] btmRightV = { outline[0], outline[3], 0 };

                double smallestX = outline[2];
                double largestX = outline[0];

                // find extremes vertices, smallest and largest x value
                foreach(object vVert in vVerts)
                {
                    Vertex swVert = (Vertex)vVert;
                    double[] dPt = new double[3];
                    dPt = (double[])swVert.GetPoint();
                    MathPoint swPt = (MathPoint)swMath.CreatePoint(dPt);
                    swPt = swPt.IMultiplyTransform(drawToSkXF);
                    swPt = swPt.IMultiplyTransform(modelToViewXF);
                    dPt = (double[])swPt.ArrayData;

                    // top left
                    if (Math.Round(dPt[1], 5) > Math.Round(topLeftV[1], 5))
                    {
                        topLeftV[0] = dPt[0];
                        topLeftV[1] = dPt[1];
                    }
                    else if (Math.Round(dPt[1], 5) == Math.Round(topLeftV[1], 5))
                    {
                        if (Math.Round(dPt[0], 5) < Math.Round(topLeftV[0], 5))
                        {
                            topLeftV[0] = dPt[0];
                            topLeftV[1] = dPt[1];
                        }
                    }

                    // top right
                    if (Math.Round(dPt[1], 5) > Math.Round(topRightV[1], 5))
                    {
                        topRightV[0] = dPt[0];
                        topRightV[1] = dPt[1];
                    }
                    else if (Math.Round(dPt[1], 5) == Math.Round(topRightV[1], 5))
                    {
                        if (Math.Round(dPt[0], 5) > Math.Round(topRightV[0], 5))
                        {
                            topRightV[0] = dPt[0];
                            topRightV[1] = dPt[1];
                        }
                    }

                    // btm left
                    if (Math.Round(dPt[1], 5) < Math.Round(btmLeftV[1], 5))
                    {
                        btmLeftV[0] = dPt[0];
                        btmLeftV[1] = dPt[1];
                    }
                    else if (Math.Round(dPt[1], 5) == Math.Round(btmLeftV[1], 5))
                    {
                        if (Math.Round(dPt[0], 5) < Math.Round(btmLeftV[0], 5))
                        {
                            btmLeftV[0] = dPt[0];
                            btmLeftV[1] = dPt[1];
                        }
                    }

                    // btm right
                    if (Math.Round(dPt[1], 5) < Math.Round(btmRightV[1], 5))
                    {
                        btmRightV[0] = dPt[0];
                        btmRightV[1] = dPt[1];
                    }
                    else if (Math.Round(dPt[1], 5) == Math.Round(btmRightV[1], 5))
                    {
                        if (Math.Round(dPt[0], 5) > Math.Round(btmRightV[0], 5))
                        {
                            btmRightV[0] = dPt[0];
                            btmRightV[1] = dPt[1];
                        }
                    }

                    // smallest
                    if (Math.Round(dPt[0], 6) < Math.Round(smallestX, 6))
                    {
                        smallestX = dPt[0];
                    }

                    // largest
                    if (Math.Round(dPt[0], 6) > Math.Round(largestX, 6))
                    {
                        largestX = dPt[0];
                    }
                }

                List<double[]> vSxPt = new List<double[]>();
                List<double[]> vLxPt = new List<double[]>();

                // find the points at smallest x and largest x
                foreach(object vVert in vVerts)
                {
                    Vertex swVert = (Vertex)vVert;
                    double[] dPt = new double[3];
                    dPt = (double[])swVert.GetPoint();
                    MathPoint swPt = (MathPoint)swMath.CreatePoint(dPt);
                    swPt = swPt.IMultiplyTransform(drawToSkXF);
                    swPt = swPt.IMultiplyTransform(modelToViewXF);
                    dPt = (double[])swPt.ArrayData;

                    if (Math.Round(dPt[0], 6) == Math.Round(smallestX, 6))
                    {
                        vSxPt.Add(dPt);
                    }
                    
                    if (Math.Round(dPt[0], 6) == Math.Round(largestX, 6))
                    {
                        vLxPt.Add(dPt);
                    }
                }

                double[] dSxPt = vSxPt.First();
                double[] dLxPt = vLxPt.First();

                Sketch swSketch = swView.IGetSketch();
                MathTransform modelToSkXF = swSketch.ModelToSketchTransform;

                double angleXOffset = 0.012;
                int isTab = 0;

                double[] edgeTopMid = { (topLeftV[0] + topRightV[0]) / 2, topLeftV[1] };
                double[] edgeBtmMid = { (btmLeftV[0] + btmRightV[0]) / 2, btmLeftV[1] };

                bRet = swDraw.ActivateView(swView.GetName2());

                double[] lenDimLoc = { 0, 0 };
                double[] mostLeftPt = new double[2];
                double[] mostRightPt = new double[2];
                double[] rightLenDimLoc = new double[2];
                double[] leftLenDimLoc = new double[2];

                // Left side
                if(vSxPt.Count == 1)
                {
                    if (Math.Round(smallestX, 6) != Math.Round(topLeftV[0], 6) && Math.Round(smallestX, 6) != Math.Round(btmLeftV[0], 6))
                    {
                        // dual angle
                        double[] edgeMid1 = { (dSxPt[0] + topLeftV[0]) / 2, (dSxPt[1] + topLeftV[1]) / 2 };
                        double[] edgeMid2 = { (dSxPt[0] + btmLeftV[0]) / 2, (dSxPt[1] + btmLeftV[1]) / 2 };
                        double[] angDimLoc1 = { topLeftV[0] + 0.02, (topLeftV[1] + 0.01) };
                        double[] angDimLoc2 = { btmLeftV[0] + 0.02, btmLeftV[1] - 0.01 };
                        leftLenDimLoc[0] = dSxPt[0] - 0.01;
                        leftLenDimLoc[1] = topLeftV[1] + 0.01;
                        double[] angLenDimLoc = { dSxPt[0] - 0.01, btmLeftV[1] - 0.01 };
                        mostLeftPt = dSxPt;

                        bRet = DimensionAngle(swModel, swMath, edgeMid1, edgeTopMid, angDimLoc1, topLeftV[0], topLeftV[1], dSxPt[0], dSxPt[1], modelToSkXF);
                        bRet = DimensionAngle(swModel, swMath, edgeMid2, edgeBtmMid, angDimLoc2, dSxPt[0], dSxPt[1], btmLeftV[0], btmLeftV[1], modelToSkXF);
                        bRet = DimensionHorizontal(swModel, topLeftV[0], topLeftV[1], dSxPt[0], dSxPt[1], leftLenDimLoc);
                        bRet = DimensionHorizontal(swModel, dSxPt[0], dSxPt[1], btmLeftV[0], btmLeftV[1], angLenDimLoc);

                        lenDimLoc[1] += 0.01;
                    }
                    else
                    {
                        // single angle
                        double[] edgeMid = { (btmLeftV[0] + topLeftV[0]) / 2, (btmLeftV[1] + topLeftV[1]) / 2 };
                        if (Math.Round(topLeftV[0], 6) > Math.Round(btmLeftV[0], 6))
                        {
                            double[] angDimLoc = { btmLeftV[0] - angleXOffset, btmLeftV[1] };
                            leftLenDimLoc[0] = btmLeftV[0] - 0.01;
                            leftLenDimLoc[1] = topLeftV[1] + 0.01;
                            mostLeftPt = btmLeftV;

                            bRet = DimensionAngle(swModel, swMath, edgeMid, edgeTopMid, angDimLoc, btmLeftV[0], btmLeftV[1], topLeftV[0], topLeftV[1], modelToSkXF);
                            bRet = DimensionHorizontal(swModel, topLeftV[0], topLeftV[1], btmLeftV[0], btmLeftV[1], leftLenDimLoc);
                        }
                        else if (Math.Round(topLeftV[0], 6) < Math.Round(btmLeftV[0], 6))
                        {
                            double[] angDimLoc = { topLeftV[0] - angleXOffset, topLeftV[1] };
                            leftLenDimLoc[0] = topLeftV[0] - 0.01;
                            leftLenDimLoc[1] = btmLeftV[1] - 0.01;
                            mostLeftPt = topLeftV;

                            bRet = DimensionAngle(swModel, swMath, edgeMid, edgeBtmMid, angDimLoc, btmLeftV[0], btmLeftV[1], topLeftV[0], topLeftV[1], modelToSkXF);
                            bRet = DimensionHorizontal(swModel, topLeftV[0], topLeftV[1], btmLeftV[0], btmLeftV[1], leftLenDimLoc);
                        }
                    }
                }
                else if (vSxPt.Count == 2)
                {
                    if (Math.Round(smallestX, 6) != Math.Round(topLeftV[0], 6))
                    {
                        // tab side view
                        isTab = 1;
                        foreach (double[] sxPt in vSxPt)
                        {
                            if (Math.Round(sxPt[1], 6) > Math.Round(dSxPt[1], 6))
                            {
                                dSxPt = sxPt;
                            }
                        }
                        mostLeftPt = topLeftV;
                    }
                    else
                    {
                        mostLeftPt = topLeftV;
                    }
                }
                else if(vSxPt.Count == 4)
                {
                    double[] edgeMid = { (btmLeftV[0] + topLeftV[0]) / 2, (btmLeftV[1] + topLeftV[1]) / 2 };
                    bRet = swModel.Extension.SelectByID2("", "EDGE", edgeMid[0], edgeMid[1], 0, false, 0, null, 0);
                    double tabIteration = 0;
                    if (!bRet)
                    {
                        while (!bRet)
                        {
                            tabIteration += 0.0001;
                            edgeMid[0] += 0.0001;
                            bRet = swModel.Extension.SelectByID2("", "EDGE", edgeMid[0], edgeMid[1], 0, false, 0, null, 0);
                            if (tabIteration > 0.001)
                            {
                                bRet = true;
                            }
                        }

                        if (bRet)
                        {
                            isTab = 2;
                            mostLeftPt = edgeMid;
                            foreach (double[] sxPt in vSxPt)
                            {
                                if (Math.Round(sxPt[1], 6) > Math.Round(dSxPt[1], 6))
                                {
                                    dSxPt = sxPt;
                                }
                            }
                        }
                    }
                    else
                    {
                        mostLeftPt = topLeftV;
                    }
                }
                else
                {
                    mostLeftPt = topLeftV;
                }

                // Right side
                if (vLxPt.Count == 1)
                {
                    if (Math.Round(largestX, 6) != Math.Round(topRightV[0], 6) && Math.Round(largestX, 6) != Math.Round(btmRightV[0], 6))
                    {
                        // dual angle
                        double[] edgeMid1 = { (dLxPt[0] + topRightV[0]) / 2, (dLxPt[1] + topRightV[1]) / 2 };
                        double[] edgeMid2 = { (dLxPt[0] + btmRightV[0]) / 2, (dLxPt[1] + btmRightV[1]) / 2 };
                        double[] angDimLoc1 = { topRightV[0] - 0.02, (topRightV[1] + 0.01) };
                        double[] angDimLoc2 = { btmRightV[0] - 0.02, btmRightV[1] - 0.01 };
                        rightLenDimLoc[0] = dLxPt[0] + 0.01;
                        rightLenDimLoc[1] = topRightV[1] + 0.01;
                        double[] angLenDimLoc = { dLxPt[0] + 0.01, btmRightV[1] - 0.01 };
                        mostRightPt = dLxPt;

                        bRet = DimensionAngle(swModel, swMath, edgeMid1, edgeTopMid, angDimLoc1, topRightV[0], topRightV[1], dLxPt[0], dLxPt[1], modelToSkXF);
                        bRet = DimensionAngle(swModel, swMath, edgeMid2, edgeBtmMid, angDimLoc2, dLxPt[0], dLxPt[1], btmRightV[0], btmRightV[1], modelToSkXF);
                        bRet = DimensionHorizontal(swModel, topRightV[0], topRightV[1], dLxPt[0], dLxPt[1], rightLenDimLoc);
                        bRet = DimensionHorizontal(swModel, dLxPt[0], dLxPt[1], btmRightV[0], btmRightV[1], angLenDimLoc);

                        lenDimLoc[1] += 0.01;
                    }
                    else
                    {
                        // single angle
                        double[] edgeMid = { (btmRightV[0] + topRightV[0]) / 2, (btmRightV[1] + topRightV[1]) / 2 };
                        if (Math.Round(topRightV[0], 6) < Math.Round(btmRightV[0], 6))
                        {
                            double[] angDimLoc = { btmRightV[0] + angleXOffset, btmRightV[1] };
                            rightLenDimLoc[0] = btmRightV[0] + 0.01;
                            rightLenDimLoc[1] = topRightV[1] + 0.01;
                            mostRightPt = btmRightV;

                            bRet = DimensionAngle(swModel, swMath, edgeMid, edgeTopMid, angDimLoc, btmRightV[0], btmRightV[1], topRightV[0], topRightV[1], modelToSkXF);
                            bRet = DimensionHorizontal(swModel, topRightV[0], topRightV[1], btmRightV[0], btmRightV[1], rightLenDimLoc);
                        }
                        else if (Math.Round(topRightV[0], 6) > Math.Round(btmRightV[0], 6))
                        {
                            double[] angDimLoc = { topRightV[0] + angleXOffset, topRightV[1] };
                            rightLenDimLoc[0] = topRightV[0] + 0.01;
                            rightLenDimLoc[1] = btmRightV[1] - 0.01;
                            mostRightPt = topRightV;

                            bRet = DimensionAngle(swModel, swMath, edgeMid, edgeBtmMid, angDimLoc, btmRightV[0], btmRightV[1], topRightV[0], topRightV[1], modelToSkXF);
                            bRet = DimensionHorizontal(swModel, topRightV[0], topRightV[1], btmRightV[0], btmRightV[1], rightLenDimLoc);
                        }
                    }
                }
                else if (vLxPt.Count == 2)
                {
                    if (Math.Round(largestX, 6) != Math.Round(topRightV[0], 6))
                    {
                        // tab side view
                        isTab = 1;
                        foreach (double[] lxPt in vLxPt)
                        {
                            if (Math.Round(lxPt[1], 6) > Math.Round(dLxPt[1], 6))
                            {
                                dLxPt = lxPt;
                            }
                        }
                        mostRightPt = topRightV;
                    }
                    else
                    {
                        mostRightPt = topRightV;
                    }
                }
                else if (vLxPt.Count == 4)
                {
                    double[] edgeMid = { (btmRightV[0] + topRightV[0]) / 2, (btmRightV[1] + topRightV[1]) / 2 };
                    bRet = swModel.Extension.SelectByID2("", "EDGE", edgeMid[0], edgeMid[1], 0, false, 0, null, 0);
                    double tabIteration = 0;
                    if (!bRet)
                    {
                        while (!bRet)
                        {
                            tabIteration += 0.0001;
                            edgeMid[0] -= 0.0001;
                            bRet = swModel.Extension.SelectByID2("", "EDGE", edgeMid[0], edgeMid[1], 0, false, 0, null, 0);
                            if (tabIteration > 0.001)
                            {
                                bRet = true;
                            }
                        }

                        if (bRet)
                        {
                            isTab = 2;
                            mostRightPt = edgeMid;
                            foreach (double[] lxPt in vLxPt)
                            {
                                if (Math.Round(lxPt[1], 6) > Math.Round(dLxPt[1], 6))
                                {
                                    dLxPt = lxPt;
                                }
                            }
                        }
                    }
                    else
                    {
                        mostRightPt = topRightV;
                    }
                }
                else
                {
                    mostRightPt = topRightV;
                }

                lenDimLoc[0] = (mostRightPt[0] + mostLeftPt[0]) / 2;
                if (rightLenDimLoc[1] > topRightV[1])
                {
                    lenDimLoc[1] = lenDimLoc[1] + topRightV[1] + 0.015;
                }
                else if (leftLenDimLoc[1] > topLeftV[1])
                {
                    lenDimLoc[1] = lenDimLoc[1] + topLeftV[1] + 0.015;
                }
                else
                {
                    lenDimLoc[1] = topRightV[1] + 0.01;
                }

                bRet = swModel.Extension.SelectByID2("", "VERTEX", mostRightPt[0], mostRightPt[1], 0, false, 0, null, 0);
                bRet = swModel.Extension.SelectByID2("", "VERTEX", mostLeftPt[0], mostLeftPt[1], 0, true, 0, null, 0);
                DisplayDimension swDispDimLen = swModel.IAddHorizontalDimension2(lenDimLoc[0], lenDimLoc[1], 0);
                if (swDispDimLen != null)
                {
                    swDispDimLen.CenterText = true;
                }
                swModel.ClearSelection2(true);

                if(isTab == 1)
                {
                    swModel.ViewZoomTo2(dSxPt[0] - 0.005, dSxPt[1] - 0.005, 0, dSxPt[0] + 0.005, dSxPt[1] + 0.005, 0);
                    bRet = swModel.Extension.SelectByID2("", "VERTEX", dSxPt[0], dSxPt[1], 0, false, 0, null, 0);
                    swModel.ViewZoomTo2(dLxPt[0] - 0.005, dLxPt[1] - 0.005, 0, dLxPt[0] + 0.005, dLxPt[1] + 0.005, 0);
                    bRet = swModel.Extension.SelectByID2("", "VERTEX", dLxPt[0], dLxPt[1], 0, true, 0, null, 0);
                    swDispDimLen = swModel.IAddHorizontalDimension2(lenDimLoc[0] + 0.005, lenDimLoc[1] + 0.005, 0);
                    if (swDispDimLen != null)
                    {
                        swDispDimLen.CenterText = true;
                        swDispDimLen.ShowParenthesis = true;
                    }
                    swModel.ClearSelection2(true);
                }
                else if(isTab==2)
                {
                    swModel.ViewZoomTo2(dSxPt[0] - 0.005, dSxPt[1] - 0.005, 0, dSxPt[0] + 0.005, dSxPt[1] + 0.005, 0);
                    bRet = swModel.Extension.SelectByID2("", "VERTEX", dSxPt[0], dSxPt[1], 0, false, 0, null, 0);
                    swModel.ViewZoomTo2(dLxPt[0] - 0.005, dLxPt[1] - 0.005, 0, dLxPt[0] + 0.005, dLxPt[1] + 0.005, 0);
                    bRet = swModel.Extension.SelectByID2("", "VERTEX", dLxPt[0], dLxPt[1], 0, true, 0, null, 0);
                    swDispDimLen = swModel.IAddHorizontalDimension2(lenDimLoc[0] + 0.005, lenDimLoc[1] + 0.005, 0);
                    if (swDispDimLen != null)
                    {
                        swDispDimLen.CenterText = true;
                        swDispDimLen.ShowParenthesis = true;
                    }
                    swModel.ClearSelection2(true);
                }
            }

            //swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, inputDimDefVal);
            //swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, dimSnappingVal);
        }

        // Method to dimension angle
        private bool DimensionAngle(ModelDoc2 swModel, MathUtility swMath, double[] D1, double[] D2, double[] D3, double st1, double st2, double end1, double end2, MathTransform modelToSkXF)
        {
            bool bRet = false;
            DisplayDimension swDispDimAng = default(DisplayDimension);

            // dimension angle
            swModel.Extension.SelectByID2("", "EDGE", D1[0], D1[1], 0, false, 0, null, 0);
            swModel.Extension.SelectByID2("", "EDGE", D2[0], D2[1], 0, true, 0, null, 0);
            swDispDimAng = (DisplayDimension)swModel.AddDimension2(D3[0], D3[1], 0);

            // draw segment if failed
            if (swDispDimAng == null)
            {
                double[] dStPt = new double[3];
                double[] dEndPt = new double[3];
                double[] stPt = new double[3];
                double[] endPt = new double[3];
                MathPoint swPt = default(MathPoint);

                dStPt[0] = st1;
                dStPt[1] = st2;
                dStPt[2] = 0;
                dEndPt[0] = end1;
                dEndPt[1] = end2;
                dEndPt[2] = 0;

                swPt = (MathPoint)swMath.CreatePoint(dStPt);
                swPt = (MathPoint)swPt.MultiplyTransform(modelToSkXF);
                stPt = (double[])swPt.ArrayData;

                swPt = (MathPoint)swMath.CreatePoint(dEndPt);
                swPt = (MathPoint)swPt.MultiplyTransform(modelToSkXF);
                endPt = (double[])swPt.ArrayData;

                SketchSegment leftSeg1 = swModel.SketchManager.CreateLine(stPt[0], stPt[1], stPt[2], endPt[0], endPt[1], endPt[2]);
                leftSeg1.ConstructionGeometry = true;
                swModel.Extension.SelectByID2("", "SKETCHSEGMENT", D1[0], D1[1], 0, false, 0, null, 0);
                swModel.Extension.SelectByID2("", "EDGE", D2[0], D2[1], 0, true, 0, null, 0);
                swDispDimAng = (DisplayDimension)swModel.AddDimension2(D3[0], D3[1], 0);
            }

            if (swDispDimAng != null)
            {
                swDispDimAng.CenterText = true;
            }

            swModel.ClearSelection2(true);

            return bRet;
        }

        // Method to dimension horizontal line
        private bool DimensionHorizontal(ModelDoc2 swModel, double D1, double D2, double D3, double D4, double[] D5)
        {
            DisplayDimension swDispDim = default(DisplayDimension);
            bool bRet = false;

            swModel.Extension.SelectByID2("", "VERTEX", D1, D2, 0, false, 0, null, 0);
            swModel.Extension.SelectByID2("", "VERTEX", D3, D4, 0, true, 0, null, 0);
            swDispDim = (DisplayDimension)swModel.AddHorizontalDimension2(D5[0], D5[1], 0);
            if(swDispDim != null)
            {
                swDispDim.ShowParenthesis = true;
            }
            swModel.ClearSelection2(true);

            return bRet;
        }

        // Method to check if any active document is opened and the document type is correct
        public bool checkDocNullandType(ModelDoc2 swModel, swDocumentTypes_e swDocType, swDocumentTypes_e swDocType2 = swDocumentTypes_e.swDocNONE)
        {
            if (swModel == null) 
            {
                msg.ErrorMsg(msg.ModelDocNull);
                return false;
            }

            if(swModel.GetType() != (int)swDocType)
            {
                switch (swDocType)
                {
                    case swDocumentTypes_e.swDocPART:
                        {
                            if(swDocType2 == swDocumentTypes_e.swDocASSEMBLY)
                            {
                                if(swModel.GetType() != (int)swDocType2)
                                {
                                    msg.ErrorMsg(msg.NotPartNorAssemDoc);
                                    return false;
                                }
                                else
                                {
                                    return true;
                                }
                            }
                            else
                            {
                                msg.ErrorMsg(msg.NotPartDoc); 
                                return false;
                            }
                        }
                    case swDocumentTypes_e.swDocASSEMBLY:
                        {
                            msg.ErrorMsg(msg.NotAssemDoc);
                            return false;
                        }
                    case swDocumentTypes_e.swDocDRAWING:
                        {
                            msg.ErrorMsg(msg.NotDrawingDoc);
                            return false;
                        }
                }
            }

            return true;
        }
    }
}
