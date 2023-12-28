using mStructural.WPF;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Drawing;

namespace mStructural.Function
{
    public class PrepDxf
    {
        #region Private Region
        private SldWorks swApp;
        private ModelDoc2 swModel;
        private DrawingDoc swDraw;
        private Sheet swSheet;
        private Message msg;
        private string fullMatCode = "";
        private string basicMaterial = "";
        private string plateThickness = "";
        private bool bBlack = false;
        private bool isUpdateQty = false;
        private string topAssemblyPath = "";
        private string selectedConfig = "";
        #endregion

        // Constructor
        public PrepDxf(SldWorks swapp, ModelDoc2 swmodel, string fullmatcode, string basicmaterial, string platethickness, bool bblack, bool isupdateqty, string topassemblypath, string selectedconfig)
        {
            swApp = swapp;
            swModel = swmodel;
            fullMatCode = fullmatcode;
            basicMaterial = basicmaterial;
            plateThickness = platethickness;
            bBlack = bblack;
            isUpdateQty = isupdateqty;
            topAssemblyPath = topassemblypath;
            selectedConfig = selectedconfig;
            msg = new Message(swApp);
        }

        // Main Method
        public void Run()
        {
            // cast to swdraw
            swDraw = (DrawingDoc)swModel;

            // Get sheet
            swSheet = swDraw.IGetCurrentSheet();

            // delete all note
            DeleteAllNote();

            // the first view is sheet
            View swView = swDraw.IGetFirstView();

            // text offset size
            double bomTextOffset_x = 0.007;
            double bomTextOffset_y = 0.004;

            // set sheet scale to 1:1
            bool bRet = swSheet.SetScale(1, 1, true, false);

            // next view is the actual first view
            swView = swView.IGetNextView();
            while (swView != null) // loop all views
            {
                // run method to color etch edge
                bRet = ColorEtchEdges(swDraw, swView);

                // get outline of the view
                double[] outline = (double[])swView.GetOutline();

                // if coloured the etch edge, need to add note
                if (bRet)
                {
                    // select view first
                    swModel.Extension.SelectByID2(swView.Name, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);

                    // insert note
                    Note swNote = swModel.IInsertNote("*LINES MARKED IN RED REQUIRED ETCHING");

                    // clear selection
                    swModel.ClearSelection2(true);

                    // set justification to center
                    swNote.SetTextJustification((int)swTextJustification_e.swTextJustificationCenter);

                    // get the annotation for this note
                    Annotation swAnn = swNote.IGetAnnotation();

                    // set position
                    swAnn.SetPosition2((outline[0] + outline[2]) / 2, outline[1] - 0.01, 0);
                }

                // add description note
                bRet = AddDescriptionNote(swView, outline, bomTextOffset_x, bomTextOffset_y);

                // get next view
                swView = swView.IGetNextView();
            }

            //string plateThickness = Interaction.InputBox("Enter plate thickness");
            swSheet.SetName("DXF " + fullMatCode);
            CreateScaleBlock(fullMatCode);

            // delete sheet format
            string sheetFormatName = swSheet.GetSheetFormatName();
            if (sheetFormatName != "")
            {
                bRet = swModel.Extension.SelectByID2(sheetFormatName, "SHEET", 0, 0, 0, false, 0, null, 0);
                bRet = swModel.Extension.DeleteSelection2((int)swDeleteSelectionOptions_e.swDelete_Children);
            }

            // message box
            msg.InfoMsg("Prep DXF Function completed!");
        }

        // Method to colour etch edge
        private bool ColorEtchEdges(DrawingDoc swDraw, View swView)
        {
            bool bRet = false;

            // get all visible components
            object[] vComps = (object[])swView.GetVisibleComponents();
            
            // should have only 1 component
            Component2 swComp = (Component2)vComps[0];

            // get all visible edge
            object[] vEdges = (object[])swView.GetVisibleEntities2(swComp,(int)swViewEntityType_e.swViewEntityType_Edge);

            // color all edges to black first
            if (bBlack)
            {
                foreach(object vEdge in vEdges)
                {
                    Entity swEnt = (Entity)vEdge;
                    swEnt.Select4(true, null);
                }

                Color blackColour = Color.Black;
                int blackColourWin = ColorTranslator.ToWin32(blackColour);
                swDraw.SetLineColor(blackColourWin);

                swModel.ClearSelection2(true);
            }

            // process each edge
            foreach(object edge in vEdges)
            {
                Edge swEdge = (Edge)edge;

                // cast to edge
                Entity swEnt = (Entity)swEdge;

                // color edges that has etch in the model v1.8.0
                if(swEnt.ModelName.Contains("etch"))
                {
                    swEnt.Select4(true, null);
                    bRet = true;
                    continue ;
                }

                // get start point
                Vertex stVert = swEdge.IGetStartVertex();

                // if there is a start point
                if(stVert !=null)
                {
                    // get edges
                    object[] vStEdges = (object[])stVert.GetEdges();
                    foreach(object stEdge in vStEdges) // loop all edges
                    {
                        Edge swStEdge = (Edge)stEdge;
                        Curve swCurve = swStEdge.IGetCurve();
                        double[] swCurveData = (double[])swStEdge.GetCurveParams2();
                        double edgeLength = Math.Round(swCurve.GetLength2(swCurveData[6], swCurveData[7]),5);
                        if (edgeLength == 0.0001 || edgeLength == 0.0005)
                        {
                            swEnt.Select4(true, null);
                            bRet = true;
                            break;
                        }
                    }
                }

                // get end point
                Vertex enVert = swEdge.IGetEndVertex();

                // if there is an end point
                if (enVert != null)
                {
                    // get edges
                    object[] vEnEdges = (object[])enVert.GetEdges();
                    foreach (object enEdge in vEnEdges) // loop all edges
                    {
                        Edge swEnEdge = (Edge)enEdge;
                        Curve swCurve = swEnEdge.IGetCurve();
                        double[] swCurveData = (double[])swEnEdge.GetCurveParams2();
                        double edgeLength = Math.Round(swCurve.GetLength2(swCurveData[6], swCurveData[7]), 5);
                        if (edgeLength== 0.0001 || edgeLength == 0.0005)
                        {
                            swEnt.Select4(true, null);
                            bRet = true;
                            break;
                        }
                    }
                }
            }

            Color myColor = Color.Red;
            int winColor = ColorTranslator.ToWin32(myColor);
            swDraw.SetLineColor(winColor);

            swModel.ClearSelection2(true);

            return bRet;
        }

        // Method to add description note
        private bool AddDescriptionNote(View swView, double[] outline, double bomTextOffset_x, double bomTextOffset_y)
        {
            string layerStr = "ANNOTATIONS";
            string balLayerStr = "DIMENSIONS"; // for balloon

            // make sure view is using sheetscale
            swView.UseSheetScale = 1;

            // activate the view first
            bool bret = swDraw.ActivateView(swView.Name);

            // instatiate create balloon option
            BalloonOptions swBalOpt = swModel.Extension.CreateBalloonOptions();
            swBalOpt.Style = (int)swBalloonStyle_e.swBS_None;
            swBalOpt.UpperTextContent = (int)swBalloonTextContent_e.swBalloonTextItemNumber;

            // get viisble edge and select the first edge
            object[] vComps = (object[])swView.GetVisibleComponents();
            Component2 swComp = (Component2)vComps[0];
            object[] vEdges = (object[])swView.GetVisibleEntities2(swComp, (int)swViewEntityType_e.swViewEntityType_Edge);
            Edge swEdge = (Edge)vEdges[0];
            Entity swEnt = (Entity)swEdge;
            swView.SelectEntity(swEnt, false);

            // insert the balloon
            Note swNoteBom = swModel.Extension.InsertBOMBalloon2(swBalOpt);
            swModel.ClearSelection2(true);
            if(swNoteBom == null)
            {
                return false;
            }

            // set balloon leader and position
            Annotation swAnnotBOM = swNoteBom.IGetAnnotation();
            swAnnotBOM.SetLeader3(0, 0, false, false, false, false);
            swAnnotBOM.Layer = balLayerStr; // put in dimensions layer v1.8.0
            if (swAnnotBOM != null)
            {
                bret = swAnnotBOM.SetPosition2(outline[0] - 0.005, outline[1] - 0.005, 0);
                //bret = swAnnotBOM.SetPosition2(outline[0] + bomTextOffset_x + balloonNoteOffsetX, (outline[1] + outline[3]) / 2 + balloonNoteOffsetY, 0);
            }

            // get part number
            string partNumber = swNoteBom.GetText();
            string balloonNumber = "";
            if (isUpdateQty)
            {
                CountTotalQtyFromDrawView2 macro = new CountTotalQtyFromDrawView2();
                balloonNumber = macro.GetBodyQuantityFromView(swApp, swView, topAssemblyPath, selectedConfig).ToString();
            }
            else
            {
                balloonNumber = swNoteBom.GetBomBalloonText(false);
            }

            string noteString = CreateNoteString(partNumber, balloonNumber);

            // remove this session so that the balloon can be used for drawing checker v1.8.0
            // delete balloon afterwards
            // swAnnotBOM.Select3(false, null);
            // swModel.Extension.DeleteSelection2((int)swDeleteSelectionOptions_e.swDelete_Absorbed);
            // swModel.ClearSelection2(true);

            // insert note, set leader and position
            swModel.Extension.SelectByID2(swView.Name, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
            Note swNote = swModel.IInsertNote(noteString);
            swModel.ClearSelection2(true);
            Annotation swAnnot = swNote.IGetAnnotation();
            swAnnot.Layer = layerStr;
            if ( swAnnot != null)
            {
                bret = swAnnot.SetPosition2(outline[0] + 0.015, (outline[1] + outline[3]) / 2 - bomTextOffset_y + 0.005, 0);
            }

            // check if note exceed view outline
            int scaledNoteCount = 0;
            double balloonNoteOffsetX = 0.0135;
            double balloonNoteOffsetY = 0.0055;
            bool bCheck = checkOverlap(outline, swNote);
            while (bCheck)
            {
                // need scaling
                double currentHeight = swNote.GetHeight();
                swNote.SetHeight(currentHeight - 0.0005);
                swModel.ForceRebuild3(false);
                scaledNoteCount += 1;
                balloonNoteOffsetX += 0.0003;
                balloonNoteOffsetY -= 0.0007;

                bCheck = checkOverlap(outline, swNote);
            }

            return true;
        }

        // Method to create note string
        private string CreateNoteString(string partNumber, string balloonNumber)
        {
            string noteString = "";
            noteString += "PN: " + partNumber + "\r\n";
            noteString += "$PRPWLD:" + "\"Description\"\r\n";
            noteString += "MAT: "+basicMaterial+"\r\n";
            noteString += "MAT2: $PRPWLD:" + "\"MATERIAL\"\r\n";
            noteString += "THK: " +plateThickness + "\r\n";
            noteString += "QTY: " + balloonNumber + " OFF";
            return noteString;
        }

        // Method to create scale block
        private bool CreateScaleBlock(string plateThickness)
        {
            string layerName = "DIMENSIONS";
            double X = 0.025;
            double Y = 0.01;
            double delX = 0.1;
            double delY = 0.1;
            SketchManager swSketchMgr = swModel.SketchManager;
            bool bRet = swDraw.ActivateSheet(swSheet.GetName());
            swSketchMgr.AddToDB = true;
            SketchSegment[] skSegment = new SketchSegment[3];
            skSegment[0] = swModel.SketchManager.CreateLine(X, Y, 0, X + delX, Y, 0);
            skSegment[0].Layer = layerName;
            skSegment[1] = swModel.SketchManager.CreateLine(X, Y, 0, X, Y + delY, 0);
            skSegment[1].Layer = layerName;

            Note[] swNote = new Note[3];
            swNote[0] = AddScaleText(X + delX / 2 - 0.01, Y - 0.002, delX * 1000 + "mm");
            swNote[0].IGetAnnotation().Layer = layerName;
            swNote[1] = AddScaleText(X - 0.02, Y + delY / 2, delY * 1000 + "mm");
            swNote[1].IGetAnnotation().Layer = layerName;
            swNote[2] = AddScaleText(X + delX / 2 - 0.02, Y + delY / 2, "THICKNESS: " + plateThickness + "mm");
            swNote[2].IGetAnnotation().Layer = layerName;

            long nbrSelObjects = swModel.Extension.MultiSelect2(skSegment, true, null);
            nbrSelObjects = swModel.Extension.MultiSelect2(swNote, true, null);

            SketchBlockDefinition swSketchBlockDef = swSketchMgr.MakeSketchBlockFromSelected(null);
            object[] vBlockInst = (object[])swSketchBlockDef.GetInstances();
            SketchBlockInstance swSketchBlockInst = (SketchBlockInstance)vBlockInst[0];
            swSketchBlockInst.Layer = layerName;
            swModel.ClearSelection2(true);
            swSketchMgr.AddToDB = false;

            return true;
        }

        // Method to add scale text
        private Note AddScaleText(double xPos, double yPos, string noteText)
        {
            bool bRet = swDraw.ActivateSheet(swSheet.GetName());
            Note swNote = swModel.IInsertNote(noteText);
            Annotation swAnnot = swNote.IGetAnnotation();

            if(swAnnot != null)
            {
                bRet = swAnnot.SetPosition2(xPos, yPos, 0);
            }

            return swNote;
        }

        // Method to check if note is overlapping with view outline
        // v1.8.0 changed to compare note outline and view outline
        private bool checkOverlap(double[] outline, Note swNote)
        {
            // get lower left and upper right corner coordinate
            double[] extent = (double[])swNote.GetExtent();

            // get the height of the text
            double textHeight = swNote.GetHeight();

            // caluclate note height
            double noteHeight = extent[4] - extent[1];

            // caluclate note width
            double noteWidth = extent[3] - extent[0];

            // calculate view outline height
            double outlineHeight = outline[3] - outline[1];

            // calculate view outline width
            double outlineWidth = outline[2] - outline[0];

            // if either height or width of the note is bigger than the view height or width respectively
            if ((noteHeight>outlineHeight) ||(noteWidth>outlineWidth))
            {
                // check if text is 0.5mm , which is the minimum height this add-in set
                if ( textHeight <= 0.0005)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }

            return false;
        }

        // Method to delete all note
        private void DeleteAllNote()
        {
            swModel.ClearSelection2(true);

            View swView = swDraw.IGetFirstView();
            while(swView != null)
            {
                Note swNote = swView.IGetFirstNote();
                while(swNote != null)
                {
                    Annotation swAnn = swNote.IGetAnnotation();
                    swAnn.Select3(true, null);
                    swNote = swNote.IGetNext();
                }
                swView = swView.IGetNextView();
            }

            swModel.Extension.DeleteSelection2((int)swDeleteSelectionOptions_e.swDelete_Absorbed);

            swModel.ClearSelection2(true);
        }
    }
}
