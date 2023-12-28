using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using static mStructural.Classes.Structs;

namespace mStructural.Function
{
    public class CreateAllBodyView
    {
        #region Private Variables
        private SldWorks swApp;
        private ModelDoc2 swModel;
        private ModelDoc2 viewModel;
        private DrawingDoc swDraw;
        private MathUtility swMath;
        private Message msg;
        private Macros macros;
        private appsetting appset;
        private double gap;
        private double curX = 0;
        private double curY = 0;
        private double sheetWidth = 0;
        private double sheetHeight = 0;
        private string noteString = "";
        #endregion

        // Constructor
        public CreateAllBodyView(SWIntegration swIntegration)
        {
            swApp = swIntegration.SwApp;
            macros = new Macros(swApp);
            msg = new Message(swApp);
            appset = new appsetting();
            gap = appset.ViewGap / 1000;
        }

        // Main Method
        public void Run()
        {
            // Instantiate math utility
            swMath = swApp.IGetMathUtility();

            // Instantiate model doc 2
            swModel = swApp.IActiveDoc2;

            bool bRet = macros.checkDocNullandType(swModel, swDocumentTypes_e.swDocDRAWING); if (!bRet) return; // check

            // cast to drawing doc
            swDraw = (DrawingDoc)swModel;

            // Instantiate selection manager
            SelectionMgr swSelMgr = swModel.ISelectionManager;

            // Instantiate table annotation
            TableAnnotation clTable = default(TableAnnotation);

            // check if any table is selected
            try
            {
                clTable = (TableAnnotation)swSelMgr.GetSelectedObject6(1, -1);
            }
            catch
            {
                msg.ErrorMsg("Please select a weldment cutlist to proceed");
                return;
            }

            // check if anything selected
            if (clTable == null)
            {
                msg.ErrorMsg("No item selected");
                return;
            }

            // check if selected table is weldment cutlist
            if (clTable.Type != (int)swTableAnnotationType_e.swTableAnnotation_WeldmentCutList)
            {
                msg.ErrorMsg("Selected table is not a weldment cutlist");
                return;
            }

            // add cloumn for cutlist name name
            bRet = clTable.InsertColumn2((int)swTableItemInsertPosition_e.swTableItemInsertPosition_Last, 0, "Name", (int)swTableColumnTypes_e.swWeldTableColumnType_CutListName);
            bRet = clTable.SetColumnType(clTable.ColumnCount - 1, (int)swTableColumnTypes_e.swWeldTableColumnType_CutListName);

            // Add all cutlist name into a list
            List<string> listClName = new List<string>();
            for(int i = 1; i < clTable.RowCount; i++)
            {
                string cutlistName = clTable.DisplayedText2[i, clTable.ColumnCount - 1, false];
                listClName.Add(cutlistName);
            }

            // delete column afterwards
            bRet = clTable.DeleteColumn2(clTable.ColumnCount - 1, false);
            if (!bRet)
            {
                msg.WarnMsg("Name column deletion failed. Please delete it manually");
            }

            // Get reference modeldoc
            View swView = swDraw.IGetFirstView();
            swView = swView.IGetNextView();
            viewModel = swView.ReferencedDocument;

            // create new sheet
            Sheet swSheet = swDraw.IGetCurrentSheet();
            double[] sheetProps = (double[])swSheet.GetProperties();
            bRet = swDraw.NewSheet4("", 0, Convert.ToInt32(sheetProps[1]), 1, 10, Convert.ToBoolean(sheetProps[4]),
                swSheet.GetTemplateName(), 0, 0, "", 0, 0, 0, 0, 0, 0);

            // Set sheet scale
            swSheet = swDraw.IGetCurrentSheet();
            swSheet.SetScale(1, 10, true, true);

            // get sheet size
            double[] pos = new double[2];
            swSheet.GetSize(ref sheetWidth, ref sheetHeight);
            curY = sheetHeight;
            pos[0] = gap;
            pos[1] = sheetHeight - gap;

            // get reference configuration
            WeldmentCutListAnnotation clTableAnn = (WeldmentCutListAnnotation)clTable;
            WeldmentCutListFeature clFeat = clTableAnn.WeldmentCutListFeature;
            string clRefConfig = clFeat.Configuration;

            // change mainmodel configuration
            viewModel.ShowConfiguration2(clRefConfig);

            UserProgressBar swPBar = default(UserProgressBar);
            bRet = swApp.GetUserProgressBar(out swPBar);
            swPBar.Start(0, listClName.Count(), "mStructural: Creating View...");
            int pBarPos = 0;

            foreach(string clName in listClName)
            {
                createView(clName, pos, out pos);
                pBarPos += 1;
                swPBar.UpdateProgress(pBarPos);
            }

            swPBar.End();
            msg.InfoMsg("View Creation Completed");
        }

        // Method to create view
        private bool createView(string clName, double[] pos, out double[] outPos)
        {
            View swView = swDraw.CreateDrawViewFromModelView3(viewModel.GetPathName(), "*Front", 0, 0, 0);
            swView.SetDisplayTangentEdges2((int)swDisplayTangentEdges_e.swTangentEdgesVisibleAndFonted);
            Sheet swSheet = swDraw.IGetCurrentSheet();
            swSheet.SetScale(1, 10, true, true);
            swView.UseSheetScale = 1;

            bool bRet = selectViewBody(viewModel, swView, clName);
            if (!bRet)
            {
                outPos = pos;
                return false;
            }

            ChooseBestView(swView);

            RotateAndScaleView(swView);

            PositionView(swView, pos, out outPos);

            CreateProjectedView(swView, clName, outPos, out outPos);

            return true;
        }

        // Method to select view body
        public bool selectViewBody(ModelDoc2 swModel, View swview, string clname)
        {
            bool bRet = false;
            Feature clFeat = default(Feature);
            BodyFolder swBodyFolder = default(BodyFolder);
            object[] arrBody = null;
            object[] bodies = new object[1];
            DispatchWrapper[] arrBodiesIn = new DispatchWrapper[1];

            // select the feature with cutlist name column
            bRet = swModel.Extension.SelectByID2(clname, "BDYFOLDER", 0, 0, 0, false, 0, null, 0);

            // if failed to select
            if (!bRet)
            {
                return false;
            }

            // cast selected object to feature object
            SelectionMgr swSelMgr = swModel.ISelectionManager;
            clFeat = (Feature)swSelMgr.GetSelectedObject6(1, -1);

            if(clFeat.GetTypeName2() == "CutListFolder")
            {
                // get body folder from feature
                swBodyFolder = (BodyFolder)clFeat.GetSpecificFeature2();

                // get all body from body folder
                arrBody = (object[])swBodyFolder.GetBodies();
                
                if(arrBody != null)
                {
                    // get first body body
                    Body2 swBody = (Body2)arrBody[0];

                    // check if body is hidden
                    if (!swBody.Visible)
                    {
                        for(int i = 1; i < arrBody.Length; i++)
                        {
                            swBody = (Body2)arrBody[i];
                            if (swBody.Visible)
                            {
                                // use the first body in the list
                                bodies[0] = arrBody[0];

                                // create new dispatch wrapper for view body selection
                                arrBodiesIn[0] = new DispatchWrapper(bodies[0]);

                                // select the body for the view
                                swview.Bodies = (arrBodiesIn);

                                break;
                            }
                        }
                    }
                    else
                    {
                        // use the first body in the list
                        bodies[0] = arrBody[0];

                        // create new dispatch wrapper for view body selection
                        arrBodiesIn[0] = new DispatchWrapper(bodies[0]);

                        // select the body for the view
                        swview.Bodies = (arrBodiesIn);
                    }
                }
            }

            // clear selection after done
            swModel.ClearSelection2(true);

            return bRet;
        }

        // choose best view by comparing view area for front top and right view, the bigger the better
        private void ChooseBestView(View swView)
        {
            double[] curOutline = null;
            ViewStruct viewStruct = new ViewStruct();

            // Start with front view
            viewStruct.ViewInt = (int)swStandardViews_e.swFrontView;

            // get current outline first
            curOutline = (double[])swView.GetOutline();

            // calculate area
            viewStruct.CurArea = Math.Abs((curOutline[2] - curOutline[0]) * (curOutline[3] - curOutline[1])); // in meter

            // Must select the view first before change orientation
            swModel.Extension.SelectByID2(swView.Name, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);

            // Compare with top and right view
            viewStruct = CompareViewArea(swView, viewStruct.CurArea, viewStruct.ViewInt, (int)swStandardViews_e.swTopView);
            viewStruct = CompareViewArea(swView, viewStruct.CurArea, viewStruct.ViewInt, (int)swStandardViews_e.swRightView);

            // change back to best view
            swModel.ShowNamedView2("", viewStruct.ViewInt);

            // best practice to clear selection
            swModel.ClearSelection2(true);
        }

        // Compare View Area
        private ViewStruct CompareViewArea(View swView, double curarea, int curviewint, int comviewint)
        {
            double[] comOutline = null;
            double comArea = 0;
            ViewStruct viewstruct = new ViewStruct();
            viewstruct.CurArea = curarea;
            viewstruct.ViewInt = curviewint;

            // changed to top view
            swModel.ShowNamedView2("", comviewint);

            // get outline again
            comOutline = (double[])swView.GetOutline();

            // calculate area
            comArea = Math.Abs((comOutline[2] - comOutline[0]) * (comOutline[3] - comOutline[1])); // in meter

            // Compare
            if (comArea > curarea)
            {
                viewstruct.ViewInt = comviewint;
                viewstruct.CurArea = comArea;
            }

            return viewstruct;
        }

        // Rotate the view to align the longest entity to become horizontal
        private void RotateAndScaleView(View swView)
        {
            Edge longestEdge = macros.GetLongestEdge(swView);
            double edgeLength = macros.GetEdgeData(longestEdge).length;
            double angle = macros.GetEdgeAngle(swMath, swView, longestEdge);

            // Must select the view first before change orientation
            swModel.Extension.SelectByID2(swView.Name, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);

            // rotate view
            swDraw.DrawingViewRotate(angle);

            if (edgeLength * 1000 < 200)
            {
                // if less than 200 use 1/5 scale
                swView.ScaleDecimal = 0.2;
                noteString = "$PRPWLD:\"DESCRIPTION\"\r\nSCALE $PRP:\"SW-View Scale(View Scale)\"";
            }
            else if (edgeLength * 1000 > 1000)
            {
                // if more than 1000 use 1/20 scale
                swView.ScaleDecimal = 0.05;
                noteString = "$PRPWLD:\"DESCRIPTION\"\r\nSCALE $PRP:\"SW-View Scale(View Scale)\"";
            }
            else
            {
                noteString = "$PRPWLD:\"DESCRIPTION\"";
            }

            swModel.ClearSelection2(true);
        }

        // Function to position view
        private void PositionView(View swView, double[] curPos, out double[] newPos)
        {
            double[] viewoutline = new double[10];
            double[] extreme = new double[4];
            double[] newpos = new double[2];
            double[] notePos = new double[2];

            viewoutline = GetViewOutline(swView);

            // calc new pos
            newpos[0] = gap + viewoutline[0] / 2 + viewoutline[2] + curPos[0];
            newpos[1] = curPos[1] - (gap + viewoutline[1] / 2 + viewoutline[3]);

            // set new pos
            swView.Position = newpos;

            // get new view outline
            viewoutline = GetViewOutline(swView);

            // set note pos
            notePos[0] = viewoutline[4];
            notePos[1] = viewoutline[7]-gap;

            InsertNote(swView, notePos);

            // calc new loc y for next view
            newpos[0] = viewoutline[8] + gap;
            newpos[1] = curPos[1];

            // return loc y
            newPos = newpos;
        }

        // Function to get view outline for select body view
        private double[] GetViewOutline(View swView)
        {
            double[] ViewOutline = new double[10];
            double[] extreme = new double[4];
            double[] curpos = new double[2];

            // get outline of the view
            extreme = CalcExtreme(swView);
            ViewOutline[0] = Math.Abs(extreme[0] - extreme[2]); // width of extreme
            ViewOutline[1] = Math.Abs(extreme[1] - extreme[3]); // height of extreme
            ViewOutline[4] = (extreme[2] + extreme[0]) / 2; // x midpoint of extreme
            ViewOutline[5] = (extreme[3] + extreme[1]) / 2; // y midpoint of extreme

            // get old pos
            curpos = (double[])swView.Position;
            ViewOutline[2] = curpos[0] - ViewOutline[4]; // x offset of pos from extreme mid
            ViewOutline[3] = ViewOutline[5] - curpos[1]; // y offset of pos from extreme mid

            ViewOutline[6] = extreme[0]; // min x
            ViewOutline[7] = extreme[1]; // min y
            ViewOutline[8] = extreme[2]; // max x
            ViewOutline[9] = extreme[3]; // max y

            return ViewOutline;
        }

        // Insert text to view
        private void InsertNote(View swView, double[] notePos)
        {
            Note swNote = default(Note);
            Annotation swAnn = default(Annotation);
            bool boolstatus;

            // select view to project
            boolstatus = swModel.Extension.SelectByID2(swView.Name, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
            swNote = (Note)swModel.InsertNote(noteString);
            swNote.SetTextJustification((int)swTextJustification_e.swTextJustificationCenter);
            swAnn = (Annotation)swNote.GetAnnotation();

            swAnn.SetPosition(notePos[0], notePos[1], 0);
        }

        // Method to get extreme corners
        private double[] CalcExtreme(View swView)
        {
            Component2 swComp = default(Component2);
            object[] vComp = null;
            object[] vVert = null;
            double[] extremexy = new double[4];

            // get visible components
            vComp = (object[])swView.GetVisibleComponents();

            // cast first comp to sldworks object
            swComp = (Component2)vComp[0];

            // get visible vertices for the first component found
            vVert = (object[])swView.GetVisibleEntities2(swComp, (int)swViewEntityType_e.swViewEntityType_Vertex);

            // get drawing and view transform
            MathTransform draw2viewXF = (MathTransform)macros.DrawToViewTransform(swMath, swView).Inverse();
            MathTransform model2viewXF = swView.ModelToViewTransform;

            // get outline for comparison
            double[] outline = (double[])swView.GetOutline();

            // set baseline
            extremexy[0] = outline[2];
            extremexy[1] = outline[3];
            extremexy[2] = outline[0];
            extremexy[3] = outline[1];
            if (vVert.Count() > 0)
            {
                foreach (object v in vVert)
                {
                    Vertex swV = (Vertex)v;
                    double[] dPt = (double[])swV.GetPoint();
                    MathPoint swPt = (MathPoint)swMath.CreatePoint(dPt);
                    swPt = (MathPoint)swPt.MultiplyTransform(draw2viewXF);
                    swPt = (MathPoint)swPt.MultiplyTransform(model2viewXF);
                    dPt = (double[])swPt.ArrayData;

                    // min x 
                    if (Math.Round(dPt[0],5) < Math.Round(extremexy[0],5))
                    {
                        extremexy[0] = dPt[0];
                    }

                    // min y
                    if (Math.Round(dPt[1],5) < Math.Round(extremexy[1],5))
                    {
                        extremexy[1] = dPt[1];
                    }

                    // max x 
                    if (Math.Round(dPt[0],5) > Math.Round(extremexy[2],5))
                    {
                        extremexy[2] = dPt[0];
                    }

                    // max y 
                    if (Math.Round(dPt[1],5) > Math.Round(extremexy[3],5))
                    {
                        extremexy[3] = dPt[1];
                    }
                }
                return extremexy;
            }
            else
            {
                extremexy[0] = outline[0];
                extremexy[1] = outline[1];
                extremexy[2] = outline[2];
                extremexy[3] = outline[3];
                return extremexy;
            }
        }

        // Create projected view
        private void CreateProjectedView(View swView, string clName, double[] pos, out double[] outPos)
        {
            View rightView = default(View);
            View topView = default(View);
            double[] newPos = new double[2];
            bool boolstatus;
            double[] curPos = new double[2];
            double[] proPos = new double[2];
            double[] viewoutline = new double[10];
            double[] topViewoutline = new double[10];
            double[] rightViewoutline = new double[10];

            // select view to project
            boolstatus = swModel.Extension.SelectByID2(swView.Name, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);

            // get view outline
            viewoutline = GetViewOutline(swView);

            // get selected view pos
            curPos = (double[])swView.Position;

            // Project side
            proPos[0] = viewoutline[4] + viewoutline[0] / 2 + gap;
            proPos[1] = viewoutline[5];
            rightView = swDraw.CreateUnfoldedViewAt3(proPos[0], proPos[1], 0, true);
            rightView.AlignWithView((int)swAlignViewTypes_e.swAlignViewHorizontalCenter, swView);
            rightViewoutline = GetViewOutline(rightView);

            // select view to project
            boolstatus = swModel.Extension.SelectByID2(swView.Name, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);

            // Project to top
            proPos[0] = viewoutline[4];
            proPos[1] = viewoutline[5] + viewoutline[1] / 2 + gap;
            topView = swDraw.CreateUnfoldedViewAt3(proPos[0], proPos[1], 0, true);
            topView.AlignWithView((int)swAlignViewTypes_e.swAlignViewVerticalCenter, swView);
            topViewoutline = GetViewOutline(topView);

            // clear selection after select
            swModel.ClearSelection();

            // reposition parent views, child will follow
            double reposOffset = gap + topViewoutline[1];
            curPos[1] -= reposOffset;
            swView.Position = curPos;

            // Check if out of sheet
            rightViewoutline = GetViewOutline(rightView);
            curX = rightViewoutline[8];

            if (curX > sheetWidth)
            {
                // repos parent view
                curPos[0] = 2 * gap + viewoutline[0] / 2 + viewoutline[2];
                curPos[1] = curY - (gap + viewoutline[1] / 2 + viewoutline[3]) - reposOffset;
                swView.Position = curPos;
                rightViewoutline = GetViewOutline(rightView);

                // check Y 
                if (rightViewoutline[7]<0.046) // block height
                {
                    // delete views
                    SelectionMgr swSelMgr = swModel.ISelectionManager;
                    swModel.Extension.SelectByID2(swView.Name, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
                    swModel.Extension.SelectByID2(topView.Name, "DRAWINGVIEW", 0, 0, 0, true, 0, null, 0);
                    swModel.Extension.SelectByID2(rightView.Name, "DRAWINGVIEW", 0, 0, 0, true, 0, null, 0);
                    swModel.Extension.DeleteSelection2((int)swDeleteSelectionOptions_e.swDelete_Absorbed);
                    swModel.ClearSelection2(true);

                    // create new sheet
                    Sheet swSheet = swDraw.IGetCurrentSheet();
                    double[] sheetProps = (double[])swSheet.GetProperties();
                    bool bRet = swDraw.NewSheet4("", 0, Convert.ToInt32(sheetProps[1]), 1, 10, Convert.ToBoolean(sheetProps[4]),
                        swSheet.GetTemplateName(), 0, 0, "", 0, 0, 0, 0, 0, 0);

                    // Set sheet scale
                    swSheet = swDraw.IGetCurrentSheet();
                    swSheet.SetScale(1, 10, true, true);

                    // set new pos
                    newPos[0] = gap;
                    newPos[1] = sheetHeight - gap;

                    curY = newPos[1];

                    // recreate view at that sheet
                    createView(clName, newPos, out newPos);
                }
                else
                {
                    curX = rightViewoutline[8];
                    newPos[0] = curX + gap;
                    newPos[1] = curY;

                    curY = rightViewoutline[7] - gap;

                    // balloon right view
                    View[] views = new View[1];
                    List<Note> notes = new List<Note>();
                    views[0] = rightView;
                    macros.BalloonCutlist(views, out notes);
                }
            }
            else
            {
                double tempoY = rightViewoutline[7] - gap;
                if (tempoY < curY)
                {
                    curY = tempoY;
                }

                // calc new pos for next view
                newPos[0] = pos[0] + rightViewoutline[0] + gap;
                newPos[1] = pos[1];

                // balloon right view
                View[] views = new View[1];
                List<Note> notes = new List<Note>();
                views[0] = rightView;
                macros.BalloonCutlist(views, out notes);
            }



            outPos = newPos;
        }
    }
}
