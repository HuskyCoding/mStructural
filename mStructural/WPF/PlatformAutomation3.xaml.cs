using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace mStructural.WPF
{
    /// <summary>
    /// Interaction logic for PlatformAutomation3.xaml
    /// </summary>
    public partial class PlatformAutomation3 : Window
    {
        #region Private Variables
        private SldWorks swApp;
        private ModelDoc2 swModel;
        private appsetting appSet;

        private enum GateType
        {
            None,
            LhsSingleGate,
            RhsSingleGate,
            DoubleGate,
            BomRail
        }

        private enum PlatformType
        {
            Stairways,
            StepType
        }

        #endregion

        // constructor
        public PlatformAutomation3(SldWorks swapp)
        {
            InitializeComponent();
            TypeCb.SelectionChanged += TypeCb_SelectionChanged;
            OverhangCh.Checked += OverhangCh_Checked;
            OverhangCh.Unchecked += OverhangCh_Unchecked;
            FrameCb.SelectionChanged += FrameCb_SelectionChanged;
            swApp = swapp;
            appSet = new appsetting();
            MasterLocTb.Text = appSet.PAMasterPath;
            OutputLocTb.Text = appSet.PATargetPath;
            EventManager.RegisterClassHandler(typeof(TextBox), GotKeyboardFocusEvent, new KeyboardFocusChangedEventHandler(OnGotKeyboardFocus));
        }

        private void GenerateBtn_Click(object sender, RoutedEventArgs e)
        {
            StringBuilder sb = new StringBuilder();

            #region Check Input
            string drawingNo;
            string projectDes = "";
            string outputPath;
            string masterPath;
            PlatformType type;
            bool? isOverhang = false;
            string frame;
            bool? hasJockeyStand = false;
            string handrail;
            bool? isRemovable = false;
            GateType gate;
            bool? hasForklift = false;
            double height = 0;
            double angle = 0;
            double width = 0;
            double landingLength = 0;
            double overhangLength = 0;
            bool bRet;
            string outputProjectString;

            // Check Drawing Number
            drawingNo = DrawingNoTb.Text;
            if (drawingNo.Length != 5)
            {
                MessageBox.Show("Drawing No character length must be exactly 5.");
                return;
            }

            if (!drawingNo.All(char.IsDigit))
            {
                MessageBox.Show("Drawing No can contain character from 0 to 9 only.");
                return;
            }

            // Check Project Description
            projectDes = ProjectDesTb.Text;
            if (projectDes == "")
            {
                MessageBox.Show("Project description is empty.");
                return;
            }

            // check master path
            masterPath = MasterLocTb.Text;
            if (!Directory.Exists(masterPath))
            {
                MessageBox.Show("Master Location directory does not exist.");
                return;
            }

            // check output path
            outputPath = OutputLocTb.Text;
            if (!Directory.Exists(outputPath))
            {
                MessageBox.Show("Output Location directory does not exist.");
                return;
            }

            // option that no need to check
            switch (TypeCb.SelectedIndex)
            {
                case 0:
                    {
                        type = PlatformType.Stairways;
                        break;
                    }
                case 1:
                    {
                        type = PlatformType.StepType;
                        break;
                    }
                default:
                    {
                        type = PlatformType.Stairways;
                        break;
                    }
            }

            isOverhang = OverhangCh.IsChecked;
            frame = FrameCb.Text;
            hasJockeyStand = JockeyStandCh.IsChecked;
            handrail = HandrailCb.Text;
            isRemovable = RemovableCh.IsChecked;
            hasForklift = ForkliftCh.IsChecked;

            // check height
            bRet = double.TryParse(HeightTb.Text, out height);
            if (!bRet)
            {
                MessageBox.Show("Invalid input for Height");
                return;
            }

            // check angle
            bRet = double.TryParse(AngleTb.Text, out angle);
            if (!bRet)
            {
                MessageBox.Show("Invalid input for Angle");
                return;
            }

            if (type == PlatformType.Stairways)
            {
                if (angle > 45 || angle < 20)
                {
                    MessageBox.Show("Invalid input for Angle, stairways angle value must be in between 20 deg to 45 deg.");
                    return;
                }
            }
            else
            {
                if (angle > 70 || angle < 60)
                {
                    MessageBox.Show("Invalid input for Angle, step type angle value must be in between 60 deg to 70 deg.");
                    return;
                }
            }

            // check width
            bRet = double.TryParse(WidthTb.Text, out width);
            if (!bRet)
            {
                MessageBox.Show("Invalid input for Width");
                return;
            }

            // warning for width
            if (type == PlatformType.Stairways)
            {
                if (width < 700)
                {
                    MessageBoxResult msgBoxResult = MessageBox.Show("Width input is less than minimum width (700mm), the tread width will be less than 600mm, do you wish to continue?", "Width Input Warning", MessageBoxButton.YesNo);
                    if(msgBoxResult == MessageBoxResult.No)
                    {
                        return;
                    }
                }
            }
            else
            {
                if (width <= 550 || width>=850)
                {
                    MessageBoxResult msgBoxResult = MessageBox.Show("Width input is out of range (550mm ~ 850mm), the tread width will not fulfill the criteria (450mm ~ 750mm), do you wish to continue?", "Width Input Warning", MessageBoxButton.YesNo);
                    if (msgBoxResult == MessageBoxResult.No)
                    {
                        return;
                    }
                }
            }

            // check landing length
            bRet = double.TryParse(LandingLengthTb.Text, out landingLength);
            if (!bRet)
            {
                MessageBox.Show("Invalid input for Landing Length");
                return;
            }

            // check overhang length
            if (OverhangLengthTb.IsEnabled)
            {
                bRet = double.TryParse(OverhangLengthTb.Text, out overhangLength);
                if (!bRet)
                {
                    MessageBox.Show("Invalid input for Overhang Length");
                    return;
                }
            }

            // output project string
            outputProjectString = outputPath + "\\BT" + drawingNo;

            // get gate type
            switch (GateCb.SelectedIndex)
            {
                case 0:
                    {
                        gate = GateType.None;
                        break;
                    }
                case 1:
                    {
                        gate = GateType.LhsSingleGate;
                        break;
                    }
                case 2:
                    {
                        gate = GateType.RhsSingleGate;
                        break;
                    }
                case 3:
                    {
                        gate = GateType.DoubleGate;
                        break;
                    }
                case 4:
                    {
                        gate = GateType.BomRail;
                        break;
                    }
                default:
                    {
                        gate = GateType.None;
                        break;
                    }
            }

            #endregion

            #region calculation
            double PlatformAngleRad;
            double TreadHorizontalLength;
            double HandrailOffest;
            int TreadInstance;
            double TreadRiser;
            double TreadHypotenuse = 0;
            double TreadGoing = 0;
            double RiserNGoing;
            double TotalHandrailLength;
            double HandrailPostDistance;
            double HandrailPlateDistance;
            double FirstThreadedPlateDistance;
            double ThreadedPlateDim;
            double JockeyStandDis;
            double JockeyPlateDis;
            double WheelDis;
            double WheelBeamDis;

            double HandrailFirstPointHeight = 800;
            double HandrailFirstPointHeightFromFrame;
            double StepThreadedPlatePatternLength;

            double socketGap;

            PlatformAngleRad = angle * Math.PI / 180;
            TreadHorizontalLength = height / Math.Tan(PlatformAngleRad);

            HandrailOffest = 240;

            if (type == PlatformType.Stairways)
            {
                TreadInstance = 2;
                TreadRiser = height / TreadInstance;
                while (TreadRiser > 225)
                {
                    TreadInstance++;
                    TreadRiser = height / TreadInstance;
                }

                if (TreadRiser >= 130)
                {
                    TreadHypotenuse = TreadRiser / Math.Sin(PlatformAngleRad);
                    TreadGoing = Math.Sqrt(TreadHypotenuse * TreadHypotenuse - TreadRiser * TreadRiser);
                    while (TreadGoing > 335)
                    {
                        TreadInstance++;
                        TreadRiser = height / TreadInstance;
                        if (TreadRiser >= 130)
                        {
                            TreadHypotenuse = TreadRiser / Math.Sin(PlatformAngleRad);
                            TreadGoing = Math.Sqrt(TreadHypotenuse * TreadHypotenuse - TreadRiser * TreadRiser);
                        }
                        else
                        {
                            MessageBox.Show("Tread calculation failed, try to increase platform height or decrease platform angle.");
                            return;
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Tread calculation failed");
                    return;
                }

                if (TreadGoing >= 215)
                {
                    RiserNGoing = 2 * TreadRiser + TreadGoing;
                    while (RiserNGoing > 700)
                    {
                        TreadInstance++;
                        TreadRiser = height / TreadInstance;
                        if (TreadRiser >= 130)
                        {
                            TreadHypotenuse = TreadRiser / Math.Sin(PlatformAngleRad);
                            TreadGoing = Math.Sqrt(TreadHypotenuse * TreadHypotenuse - TreadRiser * TreadRiser);
                            if (TreadGoing >= 215)
                            {
                                RiserNGoing = 2 * TreadRiser + TreadGoing;
                            }
                            else
                            {
                                MessageBox.Show("Tread calculation failed.");
                                return;
                            }
                        }
                        else
                        {
                            MessageBox.Show("Tread calculation failed.");
                            return;
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Tread calculation failed, try to increase platform height or decrease platform angle.");
                    return;
                }

                if (RiserNGoing < 215)
                {
                    MessageBox.Show("Tread calculation failed, try to increase platform height or decrease platform angle.");
                    return;
                }

                if (TreadInstance > 18)
                {
                    MessageBox.Show("Number of tread is out of range (2 ~ 18), please refer to AS1657-2013 Clause 7.2.2.");
                }
            }
            else
            {
                TreadInstance = 2;
                TreadRiser = height / TreadInstance;
                while (TreadRiser > 300)
                {
                    TreadInstance++;
                    TreadRiser = height / TreadInstance;
                }

                if (TreadRiser < 200)
                {
                    MessageBox.Show("Tread calculation failed, try to increase platform height or decrease platform angle.");
                    return;
                }

                TreadHypotenuse = TreadRiser / Math.Sin(PlatformAngleRad);
            }

            TotalHandrailLength = TreadHorizontalLength + HandrailOffest - TreadGoing;
            HandrailPostDistance = TotalHandrailLength / 2;
            HandrailPlateDistance = HandrailPostDistance / Math.Cos(PlatformAngleRad);
            ThreadedPlateDim = 105 + 30 * Math.Tan(PlatformAngleRad);
            FirstThreadedPlateDistance = ((HandrailPostDistance - HandrailOffest) / Math.Cos(PlatformAngleRad)) - ThreadedPlateDim;

            HandrailFirstPointHeightFromFrame = 225 * Math.Cos(PlatformAngleRad);
            StepThreadedPlatePatternLength = (height - (HandrailFirstPointHeight - HandrailFirstPointHeightFromFrame)) / Math.Sin(PlatformAngleRad) - (200 - 30 / Math.Tan(PlatformAngleRad));

            socketGap = ((landingLength + overhangLength) - 225) / 2;

            JockeyStandDis = landingLength + 595;
            JockeyPlateDis = landingLength + 825;
            WheelBeamDis = Math.Round(((height - 255) / Math.Tan(PlatformAngleRad)) + landingLength + 215 - (200 / Math.Sin(PlatformAngleRad)) - (375 / 2));
            WheelDis = WheelBeamDis - 25;
            #endregion

            #region construct bitmask
            int bitmask;
            int bitStairway = 1;
            int bitSteptype = 2;
            int bitOverhang = 4;
            int bitNonOverhang = 8;
            int bitFixFrame = 16;
            int bitWheelFrame = 32;
            int bitWelded = 64;
            int bitBolted = 128;
            int bitWeldedRemovable = 256;
            int bitBoltedRemovable = 512;
            int bitType;
            int bitOverhangType;
            int bitFrameType;
            int bitHandrailType;

            if (type == PlatformType.Stairways)
            {
                bitType = bitStairway;
            }
            else
            {
                bitType = bitSteptype;
            }

            if (isOverhang == true)
            {
                bitOverhangType = bitOverhang;
            }
            else
            {
                bitOverhangType = bitNonOverhang;
            }

            if (frame == "Fix")
            {
                bitFrameType = bitFixFrame;
            }
            else
            {
                bitFrameType = bitWheelFrame;
            }

            if (handrail == "Welded")
            {
                if (isRemovable == true)
                {
                    bitHandrailType = bitWeldedRemovable;
                }
                else
                {
                    bitHandrailType = bitWelded;
                }
            }
            else
            {
                if (isRemovable == true)
                {
                    bitHandrailType = bitBoltedRemovable;
                }
                else
                {
                    bitHandrailType = bitBolted;
                }
            }

            bitmask = bitType + bitOverhangType + bitFrameType + bitHandrailType;

            // combination
            // PA00001 = 1 + 8 + 32 + 64 = 105
            // PA00002 = 2 + 8 + 32 + 64 = 106
            // PA00003 = 1 + 8 + 16 + 64 = 89
            // PA00004 = 2 + 8 + 16 + 64 = 90
            // PA00005 = 1 + 8 + 32 + 128 = 169
            // PA00006 = 2 + 8 + 32 + 128 = 170
            // PA00007 = 1 + 8 + 16 + 128 = 153
            // PA00008 = 2 + 8 + 16 + 128 = 154
            // PA00009 = 1 + 8 + 32 + 512 = 553
            // PA00010 = 2 + 8 + 32 + 512 = 554
            // PA00011 = 1 + 8 + 16 + 512 = 537
            // PA00012 = 2 + 8 + 16 + 512 = 538
            // PA00013 = 1 + 8 + 16 + 256 = 281
            // PA00014 = 2 + 8 + 16 + 256 = 282
            // PA00015 = 1 + 8 + 32 + 256 = 297
            // PA00016 = 2 + 8 + 32 + 256 = 298
            // PA00017 = 1 + 4 + 32 + 64 = 101
            // PA00018 = 2 + 4 + 32 + 64 = 102
            // PA00019 = 1 + 4 + 16 + 64 = 85
            // PA00020 = 2 + 4 + 16 + 64 = 86
            // PA00021 = 1 + 4 + 32 + 128 = 165
            // PA00022 = 2 + 4 + 32 + 128 = 166
            // PA00023 = 1 + 4 + 16 + 128 = 149
            // PA00024 = 2 + 4 + 16 + 128 = 150
            // PA00025 = 1 + 4 + 32 + 512 = 549
            // PA00026 = 2 + 4 + 32 + 512 = 550
            // PA00027 = 1 + 4 + 16 + 512 = 533
            // PA00028 = 2 + 4 + 16 + 512 = 534
            // PA00029 = 1 + 4 + 16 + 256 = 277
            // PA00030 = 2 + 4 + 16 + 256 = 278
            // PA00031 = 1 + 4 + 32 + 256 = 293
            // PA00032 = 2 + 4 + 32 + 256 = 294
            #endregion

            #region Pack and Go Model
            string platformMasterName;
            switch (bitmask)
            {
                case 105: { platformMasterName = "PA00001"; break; }
                case 106: { platformMasterName = "PA00002"; break; }
                case 89: { platformMasterName = "PA00003"; break; }
                case 90: { platformMasterName = "PA00004"; break; }
                case 169: { platformMasterName = "PA00005"; break; }
                case 170: { platformMasterName = "PA00006"; break; }
                case 153: { platformMasterName = "PA00007"; break; }
                case 154: { platformMasterName = "PA00008"; break; }
                case 553: { platformMasterName = "PA00009"; break; }
                case 554: { platformMasterName = "PA00010"; break; }
                case 537: { platformMasterName = "PA00011"; break; }
                case 538: { platformMasterName = "PA00012"; break; }
                case 281: { platformMasterName = "PA00013"; break; }
                case 282: { platformMasterName = "PA00014"; break; }
                case 297: { platformMasterName = "PA00015"; break; }
                case 298: { platformMasterName = "PA00016"; break; }
                case 101: { platformMasterName = "PA00017"; break; }
                case 102: { platformMasterName = "PA00018"; break; }
                case 85: { platformMasterName = "PA00019"; break; }
                case 86: { platformMasterName = "PA00020"; break; }
                case 165: { platformMasterName = "PA00021"; break; }
                case 166: { platformMasterName = "PA00022"; break; }
                case 149: { platformMasterName = "PA00023"; break; }
                case 150: { platformMasterName = "PA00024"; break; }
                case 549: { platformMasterName = "PA00025"; break; }
                case 550: { platformMasterName = "PA00026"; break; }
                case 533: { platformMasterName = "PA00027"; break; }
                case 534: { platformMasterName = "PA00028"; break; }
                case 277: { platformMasterName = "PA00029"; break; }
                case 278: { platformMasterName = "PA00030"; break; }
                case 293: { platformMasterName = "PA00031"; break; }
                case 294: { platformMasterName = "PA00032"; break; }
                default: { MessageBox.Show("Bitmask operation failed."); return; }
            }

            string masterFileLoc = masterPath + "\\" + platformMasterName;

            // check if directory exist
            if (!Directory.Exists(masterFileLoc))
            {
                MessageBox.Show(masterFileLoc + " directory does not exist.");
                return;
            }

            string[] allFiles = Directory.GetFiles(masterFileLoc);
            List<string> asmList = new List<string>();
            List<string> drwList = new List<string>();
            List<Tuple<string, string>> nameList = new List<Tuple<string, string>>();
            foreach (string file in allFiles)
            {
                // exclude temp files
                if (!file.Contains("~"))
                {
                    //copy files
                    string sourceFile = file;
                    string oriName = Path.GetFileName(sourceFile);
                    string newName = oriName.Replace(platformMasterName, "BT" + drawingNo);
                    string destFile = outputPath + "\\" + newName;
                    int iRet = swApp.CopyDocument(sourceFile, destFile, "", "", (int)swMoveCopyOptions_e.swMoveCopyOptionsOverwriteExistingDocs);

                    // set read-only flag to off for newly copied files
                    FileInfo fileInfo = new FileInfo(destFile);
                    fileInfo.IsReadOnly = false;

                    // add to respective list
                    nameList.Add(new Tuple<string, string>(sourceFile, destFile));
                    if (destFile.ToUpper().Contains(".SLDASM"))
                        asmList.Add(destFile);

                    if (destFile.ToUpper().Contains(".SLDDRW"))
                        drwList.Add(destFile);
                }
            }

            // update assembly reference
            foreach (string asm in asmList)
            {
                for (int i = 0; i < nameList.Count; i++)
                {
                    string sourceFile = nameList[i].Item1;
                    string destFile = nameList[i].Item2;
                    bRet = swApp.ReplaceReferencedDocument(asm, sourceFile, destFile);
                }
            }

            // update drawing reference
            foreach (string drw in drwList)
            {
                for (int i = 0; i < nameList.Count; i++)
                {
                    string sourceFile = nameList[i].Item1;
                    string destFile = nameList[i].Item2;
                    bRet = swApp.ReplaceReferencedDocument(drw, sourceFile, destFile);
                }
            }

            // open model
            string mainAsm = outputProjectString + "-AS-100.SLDASM";
            int err = -1;
            int warn = -1;
            swModel = swApp.OpenDoc6(mainAsm, (int)swDocumentTypes_e.swDocASSEMBLY, (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", ref err, ref warn);
            swModel.ForceRebuild3(true);
            #endregion

            #region Modify Frame
            string modifyModelName;
            ModelDoc2 modifyModel;
            SelectionMgr modifySelMgr;
            Feature selFeat;
            Dimension modifyDim;
            EquationMgr modifyEqMgr;
            bool bSelRet;

            int deleteOption = (int)swDeleteSelectionOptions_e.swDelete_Children + (int)swDeleteSelectionOptions_e.swDelete_Absorbed;

            modifyModelName = outputProjectString + "-PT-101.SLDPRT";
            modifyModel = swApp.IActivateDoc3(modifyModelName, true, ref err);
            modifySelMgr = modifyModel.ISelectionManager;

            // change main frame sketch dimension
            bSelRet = modifyModel.Extension.SelectByID2("Main Frame Sketch", "SKETCH", 0, 0, 0, false, 0, null, 0); // select sketch
            selFeat = (Feature)modifySelMgr.GetSelectedObject6(1, -1); // cast selected object as Feature
            ChangeFeatureDimension(selFeat, "Height", height / 1000); // change height dimension
            ChangeFeatureDimension(selFeat, "Width", width / 1000); // change width dimension
            ChangeFeatureDimension(selFeat, "Angle", PlatformAngleRad); // change angle dimension
            ChangeFeatureDimension(selFeat, "LandingLength", (landingLength - 25) / 1000); // change landing length dimension
            ChangeFeatureDimension(selFeat, "OverhangLength", overhangLength / 1000); // change overhang length dimension
            ChangeFeatureDimension(selFeat, "WheelBeam", WheelBeamDis / 1000);// change wheel beam dimension

            // change first tread etch sketch dimension 
            bSelRet = modifyModel.Extension.SelectByID2("First Tread Etch Sketch", "SKETCH", 0, 0, 0, false, 0, null, 0); // select sketch
            selFeat = (Feature)modifySelMgr.GetSelectedObject6(1, -1); // cast selected object as feature
            ChangeFeatureDimension(selFeat, "Distance", TreadHypotenuse / 1000);// change diestance dimension

            // change tread etch pattern dimension
            bSelRet = modifyModel.Extension.SelectByID2("Tread Etch", "BODYFEATURE", 0, 0, 0, false, 0, null, 0); // select pattern
            selFeat = (Feature)modifySelMgr.GetSelectedObject6(1, -1); // cast it as feature
            ChangeFeatureDimension(selFeat, "Instance", Convert.ToDouble(TreadInstance - 1)); // change instance
            ChangeFeatureDimension(selFeat, "Distance", TreadHypotenuse / 1000); // change distance

            // select handrail sketch
            bSelRet = modifyModel.Extension.SelectByID2("Handrail Sketch", "SKETCH", 0, 0, 0, false, 0, null, 0);
            if (bSelRet)
            {
                selFeat = (Feature)modifySelMgr.GetSelectedObject6(1, -1);

                if (type == PlatformType.Stairways)
                {
                    ChangeFeatureDimension(selFeat, "HrPost1", HandrailPostDistance / 1000); // change handrail first post dimension
                    ChangeFeatureDimension(selFeat, "HrPost2", HandrailPostDistance / 1000); // change handrail second post dimension
                }

                if (isRemovable == false && isOverhang == true)
                {
                    ChangeFeatureDimension(selFeat, "LandingMidPostD", (landingLength - 25) / 1000); // change mid post length dimension
                }
            }

            // select socket sketch
            bSelRet = modifyModel.Extension.SelectByID2("SocketSketch", "SKETCH", 0, 0, 0, false, 0, null, 0);
            if (bSelRet)
            {
                selFeat = (Feature)modifySelMgr.GetSelectedObject6(1, -1);
                ChangeFeatureDimension(selFeat, "SocketGap", socketGap / 1000); // change socket gap dimension
            }

            ChangeMateDimension(modifyModel, modifySelMgr, "FirstTpMate", FirstThreadedPlateDistance / 1000); // change first threaded plate mate dimension
            if (type == PlatformType.Stairways)
            {
                ChangeMateDimension(modifyModel, modifySelMgr, "WheelPlateMate", WheelDis / 1000); // change wheel plate mate dimension
            }
            else
            {
                ChangeMateDimension(modifyModel, modifySelMgr, "WheelPlateMate", (WheelDis - 45) / 1000); // change wheel plate mate dimension
            }
            ChangeMateDimension(modifyModel, modifySelMgr, "JockeyPlateMate", JockeyPlateDis / 1000); // change jockey plate mate dimension

            // select threaded plate pattern at stair
            bSelRet = modifyModel.Extension.SelectByID2("TpPatternStair", "BODYFEATURE", 0, 0, 0, false, 0, null, 0);
            if (bSelRet)
            {
                selFeat = (Feature)modifySelMgr.GetSelectedObject6(1, -1);

                // change threaded plate pattern dimension
                modifyDim = selFeat.IParameter("TpPatternStairD");
                if (type == PlatformType.Stairways)
                {
                    modifyDim.SetSystemValue3(HandrailPlateDistance / 1000, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                }
                else
                {
                    modifyDim.SetSystemValue3(StepThreadedPlatePatternLength / 1000, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                }
            }

            // select threaded plate pattern at landing
            bSelRet = modifyModel.Extension.SelectByID2("TpPatternLanding", "BODYFEATURE", 0, 0, 0, false, 0, null, 0);
            if (bSelRet)
            {
                selFeat = (Feature)modifySelMgr.GetSelectedObject6(1, -1);
                ChangeFeatureDimension(selFeat, "TpPatternLandingD", ((landingLength + overhangLength - 50) / 2) / 1000);
            }

            if (JockeyStandCh.IsEnabled && hasJockeyStand == false)
            {
                // select the jack adaptor plate and delete
                bSelRet = modifyModel.Extension.SelectByID2("PL10 250 x 170 AL JACK ADAPTOR", "BODYFEATURE", 0, 0, 0, false, 0, null, 0);
                modifyModel.Extension.DeleteSelection2(deleteOption);
            }

            if (ForkliftCh.IsEnabled && hasForklift == false)
            {
                // select the forklife main sketch and delete
                bSelRet = modifyModel.Extension.SelectByID2("Forklift Sketch", "SKETCH", 0, 0, 0, false, 0, null, 0);
                modifyModel.Extension.DeleteSelection2(deleteOption);
            }

            CloseDoc(modifyModel);
            #endregion

            #region Modify Landing Grate
            modifyModelName = outputProjectString + "-PT-102.SLDPRT";
            modifyModel = swApp.IActivateDoc3(modifyModelName, true, ref err);
            modifyEqMgr = modifyModel.GetEquationMgr();

            // change equation
            modifyEqMgr.SetEquationAndConfigurationOption(2, "\"Span\" = " + ((width - 100)).ToString(), (int)swInConfigurationOpts_e.swThisConfiguration, null);
            if (isOverhang == true)
            {
                modifyEqMgr.SetEquationAndConfigurationOption(3, "\"Width\" = " + (((landingLength + overhangLength) - 100) / 2).ToString(), (int)swInConfigurationOpts_e.swThisConfiguration, null);
            }
            else
            {
                modifyEqMgr.SetEquationAndConfigurationOption(3, "\"Width\" = " + ((landingLength - 100)).ToString(), (int)swInConfigurationOpts_e.swThisConfiguration, null);
            }

            PartDoc swPart = (PartDoc)modifyModel; // cast the model doc to part doc
            List<Body2> bodyList = new List<Body2>(); // instantiate list of body2
            object[] bodies = (object[])swPart.GetBodies2(-1, true); // get all bodies for this part
            foreach (object body in bodies)
            {
                Body2 swBody = (Body2)body;
                bodyList.Add(swBody);
            }
            FeatureManager swFeatMgr = modifyModel.FeatureManager;
            swFeatMgr.InsertCombineFeature((int)swBodyOperationType_e.SWBODYADD, null, bodyList.ToArray());

            CloseDoc(modifyModel);
            #endregion

            #region Modify Stair Tread
            modifyModelName = outputProjectString + "-PT-103.SLDPRT";
            modifyModel = swApp.IActivateDoc3(modifyModelName, true, ref err);
            modifyEqMgr = modifyModel.GetEquationMgr();

            // change equation
            modifyEqMgr.SetEquationAndConfigurationOption(2, "\"Span\" = " + ((width - 100)).ToString(), (int)swInConfigurationOpts_e.swThisConfiguration, null);

            CloseDoc(modifyModel);
            #endregion

            #region Modify Bolted Handrail
            if (handrail == "Bolted")
            {
                // modify right hand side handrail
                modifyModelName = outputProjectString + "-PT-105.SLDPRT";
                modifyModel = swApp.IActivateDoc3(modifyModelName, true, ref err);
                modifySelMgr = modifyModel.ISelectionManager;

                bSelRet = modifyModel.Extension.SelectByID2("Main Sketch", "SKETCH", 0, 0, 0, false, 0, null, 0);
                selFeat = (Feature)modifySelMgr.GetSelectedObject6(1, -1);

                // change angle
                modifyDim = selFeat.IParameter("Angle");
                modifyDim.SetSystemValue3(PlatformAngleRad, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                if (isOverhang == true)
                {
                    if (isRemovable == true)
                    {
                        // do nothing
                    }
                    else
                    {
                        // change landing length 1
                        modifyDim = selFeat.IParameter("LandingLength1");
                        modifyDim.SetSystemValue3((((landingLength + overhangLength) - 50) / 2) / 1000, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                        // change landing length 2
                        modifyDim = selFeat.IParameter("LandingLength2");
                        modifyDim.SetSystemValue3((((landingLength + overhangLength) - 50) / 2) / 1000, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                    }
                }
                else
                {
                    if (isRemovable == true)
                    {
                        // do nothing
                    }
                    else
                    {
                        // change landing length
                        modifyDim = selFeat.IParameter("LandingLength");
                        modifyDim.SetSystemValue3((landingLength - 50) / 1000, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                    }
                }

                if (type == PlatformType.Stairways)
                {
                    // change landing length
                    modifyDim = selFeat.IParameter("HrPost1");
                    modifyDim.SetSystemValue3((HandrailPostDistance) / 1000, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                    // change landing length
                    modifyDim = selFeat.IParameter("HrPost2");
                    modifyDim.SetSystemValue3((HandrailPostDistance) / 1000, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                }
                else
                {
                    // change handrail reference height
                    modifyDim = selFeat.IParameter("Height");
                    modifyDim.SetSystemValue3(height / 1000, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                }

                // select threaded plate pattern
                bSelRet = modifyModel.Extension.SelectByID2("TpPattern", "BODYFEATURE", 0, 0, 0, false, 0, null, 0);
                selFeat = (Feature)modifySelMgr.GetSelectedObject6(1, -1);

                // change distance 
                modifyDim = selFeat.IParameter("Distance");

                if (type == PlatformType.Stairways)
                {
                    modifyDim.SetSystemValue3(HandrailPlateDistance / 1000, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                }
                else
                {
                    modifyDim.SetSystemValue3(StepThreadedPlatePatternLength / 1000, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                }

                // select threaded plate pattern at landing
                bSelRet = modifyModel.Extension.SelectByID2("TpPatternL", "BODYFEATURE", 0, 0, 0, false, 0, null, 0);
                if (bSelRet)
                {
                    selFeat = (Feature)modifySelMgr.GetSelectedObject6(1, -1);
                    modifyDim = selFeat.IParameter("Distance");
                    modifyDim.SetSystemValue3(((landingLength + overhangLength - 50) / 2) / 1000, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                }

                CloseDoc(modifyModel);

                // modify left hand side handrail
                modifyModelName = outputProjectString + "-PT-106.SLDPRT";
                modifyModel = swApp.IActivateDoc3(modifyModelName, true, ref err);
                modifySelMgr = modifyModel.ISelectionManager;

                bSelRet = modifyModel.Extension.SelectByID2("Main Sketch", "SKETCH", 0, 0, 0, false, 0, null, 0);
                selFeat = (Feature)modifySelMgr.GetSelectedObject6(1, -1);

                // change angle
                modifyDim = selFeat.IParameter("Angle");
                modifyDim.SetSystemValue3(PlatformAngleRad, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                if (isOverhang == true)
                {
                    if (isRemovable == true)
                    {
                        // do nothing
                    }
                    else
                    {
                        // change landing length 1
                        modifyDim = selFeat.IParameter("LandingLength1");
                        modifyDim.SetSystemValue3((((landingLength + overhangLength) - 50) / 2) / 1000, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                        // change landing length 2
                        modifyDim = selFeat.IParameter("LandingLength2");
                        modifyDim.SetSystemValue3((((landingLength + overhangLength) - 50) / 2) / 1000, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                    }
                }
                else
                {
                    if (isRemovable == true)
                    {
                        // do nothing
                    }
                    else
                    {
                        // change landing length
                        modifyDim = selFeat.IParameter("LandingLength");
                        modifyDim.SetSystemValue3((landingLength - 50) / 1000, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                    }
                }

                if (type == PlatformType.Stairways)
                {
                    // change landing length
                    modifyDim = selFeat.IParameter("HrPost1");
                    modifyDim.SetSystemValue3((HandrailPostDistance) / 1000, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                    // change landing length
                    modifyDim = selFeat.IParameter("HrPost2");
                    modifyDim.SetSystemValue3((HandrailPostDistance) / 1000, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                }
                else
                {
                    // change handrail reference height
                    modifyDim = selFeat.IParameter("Height");
                    modifyDim.SetSystemValue3(height / 1000, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                }

                // select threaded plate pattern
                bSelRet = modifyModel.Extension.SelectByID2("TpPattern", "BODYFEATURE", 0, 0, 0, false, 0, null, 0);
                selFeat = (Feature)modifySelMgr.GetSelectedObject6(1, -1);

                // change distance 
                modifyDim = selFeat.IParameter("Distance");

                if (type == PlatformType.Stairways)
                {
                    modifyDim.SetSystemValue3(HandrailPlateDistance / 1000, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                }
                else
                {
                    modifyDim.SetSystemValue3(StepThreadedPlatePatternLength / 1000, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                }

                // select threaded plate pattern at landing
                bSelRet = modifyModel.Extension.SelectByID2("TpPatternL", "BODYFEATURE", 0, 0, 0, false, 0, null, 0);
                if (bSelRet)
                {
                    selFeat = (Feature)modifySelMgr.GetSelectedObject6(1, -1);
                    modifyDim = selFeat.IParameter("Distance");
                    modifyDim.SetSystemValue3(((landingLength + overhangLength - 50) / 2) / 1000, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                }

                CloseDoc(modifyModel);
            }
            #endregion

            #region Process Gate
            switch (gate)
            {
                case GateType.None:
                    {
                        // do nothing
                        break;
                    }
                case GateType.LhsSingleGate:
                    {
                        if (isRemovable == true)
                        {
                            modifyModelName = outputProjectString + "-AS-200.SLDASM";
                        }
                        else
                        {
                            modifyModelName = outputProjectString + "-AS-150.SLDASM";
                        }

                        // change equation
                        modifyModel = swApp.IActivateDoc3(modifyModelName, true, ref err);
                        modifyEqMgr = modifyModel.GetEquationMgr();
                        modifyEqMgr.SetEquationAndConfigurationOption(0, "\"ENTRANCE WIDTH\" = " + (width - 50).ToString(), (int)swInConfigurationOpts_e.swAllConfiguration, null);
                        CloseDoc(modifyModel);
                        break;
                    }
                case GateType.RhsSingleGate:
                    {
                        if (isRemovable == true)
                        {
                            modifyModelName = outputProjectString + "-AS-200.SLDASM";
                        }
                        else
                        {
                            modifyModelName = outputProjectString + "-AS-150.SLDASM";
                        }

                        // cahnge equation
                        modifyModel = swApp.IActivateDoc3(modifyModelName, true, ref err);
                        modifyEqMgr = modifyModel.GetEquationMgr();
                        modifyEqMgr.SetEquationAndConfigurationOption(0, "\"ENTRANCE WIDTH\" = " + (width - 50).ToString(), (int)swInConfigurationOpts_e.swAllConfiguration, null);
                        CloseDoc(modifyModel);
                        break;
                    }
                case GateType.DoubleGate:
                    {
                        if (isRemovable == true)
                        {
                            modifyModelName = outputProjectString + "-AS-300.SLDASM";
                        }
                        else
                        {
                            modifyModelName = outputProjectString + "-AS-250.SLDASM";
                        }

                        // change equation
                        modifyModel = swApp.IActivateDoc3(modifyModelName, true, ref err);
                        modifyEqMgr = modifyModel.GetEquationMgr();
                        modifyEqMgr.SetEquationAndConfigurationOption(0, "\"GATE OPENING\" = " + (width - 50).ToString(), (int)swInConfigurationOpts_e.swAllConfiguration, null);
                        CloseDoc(modifyModel);
                        break;
                    }
                case GateType.BomRail:
                    {
                        if (isRemovable == true)
                        {
                            modifyModelName = outputProjectString + "-AS-250.SLDASM";
                        }
                        else
                        {
                            modifyModelName = outputProjectString + "-AS-200.SLDASM";
                        }

                        // change equation
                        modifyModel = swApp.IActivateDoc3(modifyModelName, true, ref err);
                        modifyEqMgr = modifyModel.GetEquationMgr();
                        modifyEqMgr.Equation[0] = "\"GATE OPENING\" = " + (width - 50).ToString();
                        CloseDoc(modifyModel);
                        break;
                    }
            }
            #endregion

            #region Modify Removable Handrail
            if (isRemovable == true)
            {
                if (isOverhang == true)
                {
                    // change equation
                    modifyModelName = outputProjectString + "-AS-150.SLDASM";
                    modifyModel = swApp.IActivateDoc3(modifyModelName, true, ref err);
                    modifyEqMgr = modifyModel.GetEquationMgr();
                    modifyEqMgr.Equation[0] = "\"LENGTH\" = " + ((landingLength + overhangLength) / 2).ToString();
                    CloseDoc(modifyModel);
                }
                else
                {
                    // change equation
                    modifyModelName = outputProjectString + "-AS-150.SLDASM";
                    modifyModel = swApp.IActivateDoc3(modifyModelName, true, ref err);
                    modifyEqMgr = modifyModel.GetEquationMgr();
                    modifyEqMgr.Equation[0] = "\"LENGTH\" = " + (landingLength).ToString();
                    CloseDoc(modifyModel);
                }
            }
            #endregion

            #region Modify Assembly
            swModel = swApp.IActivateDoc3(Path.GetFileName(mainAsm), false, ref err);

            // change custom property
            CustomPropertyManager asmCusPropMgr = swModel.Extension.CustomPropertyManager[""];
            asmCusPropMgr.Add3("Description", (int)swCustomInfoType_e.swCustomInfoText, projectDes, (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);

            // select the pattern
            modifySelMgr = swModel.ISelectionManager;
            bSelRet = swModel.Extension.SelectByID2("TreadLP", "COMPPATTERN", 0, 0, 0, false, 0, null, 0);
            selFeat = (Feature)modifySelMgr.GetSelectedObject6(1, -1);

            // change tread pattern instance
            modifyDim = selFeat.IParameter("Instance");
            modifyDim.SetSystemValue3(Convert.ToDouble(TreadInstance), (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

            // change tread pattern distance
            modifyDim = selFeat.IParameter("Distance");
            modifyDim.SetSystemValue3(TreadHypotenuse / 1000, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

            #endregion

            #region Assembly Conditional
            bSelRet = swModel.Extension.SelectByID2("LandingGrateLP", "COMPPATTERN", 0, 0, 0, false, 0, null, 0);
            if (bSelRet)
            {
                selFeat = (Feature)modifySelMgr.GetSelectedObject6(1, -1);

                // change pattern distance
                modifyDim = selFeat.IParameter("Distance");
                modifyDim.SetSystemValue3(((landingLength + overhangLength) / 2 - 50) / 1000, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
            }

            // change linear pattern distance
            bSelRet = swModel.Extension.SelectByID2("WheelLP", "COMPPATTERN", 0, 0, 0, false, 0, null, 0);
            if (bSelRet)
            {
                selFeat = (Feature)modifySelMgr.GetSelectedObject6(1, -1);

                // change tread pattern instance
                modifyDim = selFeat.IParameter("Distance");
                modifyDim.SetSystemValue3(WheelDis / 1000, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
            }

            // change dimension for removable handrail pattern
            bSelRet = swModel.Extension.SelectByID2("RemovableHandrailLP", "COMPPATTERN", 0, 0, 0, false, 0, null, 0);
            if (bSelRet)
            {
                selFeat = (Feature)modifySelMgr.GetSelectedObject6(1, -1);
                ChangeFeatureDimension(selFeat, "Distance", (socketGap + 95) / 1000);
            }

            swModel.ClearSelection2(true);

            // delete unrelated assembly
            if (gate != GateType.LhsSingleGate)
            {
                if (isRemovable == true)
                {
                    swModel.Extension.SelectByID2("BT" + drawingNo + "-AS-200-3@" + Path.GetFileNameWithoutExtension(mainAsm), "COMPONENT", 0, 0, 0, false, 0, null, 0);
                }
                else
                {
                    swModel.Extension.SelectByID2("BT" + drawingNo + "-AS-150-1@" + Path.GetFileNameWithoutExtension(mainAsm), "COMPONENT", 0, 0, 0, false, 0, null, 0);
                }
                swModel.Extension.DeleteSelection2(deleteOption);
            }

            if (gate != GateType.RhsSingleGate)
            {
                if (isRemovable == true)
                {
                    swModel.Extension.SelectByID2("BT" + drawingNo + "-AS-200-4@" + Path.GetFileNameWithoutExtension(mainAsm), "COMPONENT", 0, 0, 0, false, 0, null, 0);
                }
                else
                {
                    swModel.Extension.SelectByID2("BT" + drawingNo + "-AS-150-2@" + Path.GetFileNameWithoutExtension(mainAsm), "COMPONENT", 0, 0, 0, false, 0, null, 0);
                }
                swModel.Extension.DeleteSelection2(deleteOption);
            }

            if (gate != GateType.BomRail)
            {
                if (isRemovable == true)
                {
                    swModel.Extension.SelectByID2("BT" + drawingNo + "-AS-250-4@" + Path.GetFileNameWithoutExtension(mainAsm), "COMPONENT", 0, 0, 0, false, 0, null, 0);
                }
                else
                {
                    swModel.Extension.SelectByID2("BT" + drawingNo + "-AS-200-2@" + Path.GetFileNameWithoutExtension(mainAsm), "COMPONENT", 0, 0, 0, false, 0, null, 0);
                }
                swModel.Extension.DeleteSelection2(deleteOption);
            }

            if (gate != GateType.DoubleGate)
            {
                if (isRemovable == true)
                {
                    swModel.Extension.SelectByID2("BT" + drawingNo + "-AS-300-2@" + Path.GetFileNameWithoutExtension(mainAsm), "COMPONENT", 0, 0, 0, false, 0, null, 0);
                    swModel.Extension.SelectByID2("BT" + drawingNo + "-AS-300-3@" + Path.GetFileNameWithoutExtension(mainAsm), "COMPONENT", 0, 0, 0, true, 0, null, 0);
                }
                else
                {
                    swModel.Extension.SelectByID2("BT" + drawingNo + "-AS-250-2@" + Path.GetFileNameWithoutExtension(mainAsm), "COMPONENT", 0, 0, 0, false, 0, null, 0);
                    swModel.Extension.SelectByID2("BT" + drawingNo + "-AS-250-3@" + Path.GetFileNameWithoutExtension(mainAsm), "COMPONENT", 0, 0, 0, true, 0, null, 0);
                }
                swModel.Extension.DeleteSelection2(deleteOption);
            }

            if (hasJockeyStand == true)
            {
                if (isRemovable == true)
                {
                    if (gate == GateType.LhsSingleGate)
                    {
                        swModel.Extension.SelectByID2("BT" + drawingNo + "-AS-200-3@" + Path.GetFileNameWithoutExtension(mainAsm), "COMPONENT", 0, 0, 0, false, 0, null, 0);
                        Component2 swComp = (Component2)modifySelMgr.GetSelectedObject6(1, -1);
                        swComp.ReferencedConfiguration = "LHS - WITH SIGN";
                    }

                    if (gate == GateType.RhsSingleGate)
                    {
                        swModel.Extension.SelectByID2("BT" + drawingNo + "-AS-200-4@" + Path.GetFileNameWithoutExtension(mainAsm), "COMPONENT", 0, 0, 0, false, 0, null, 0);
                        Component2 swComp = (Component2)modifySelMgr.GetSelectedObject6(1, -1);
                        swComp.ReferencedConfiguration = "RHS - WITH SIGN";
                    }
                }
                else
                {
                    if (gate == GateType.LhsSingleGate)
                    {
                        swModel.Extension.SelectByID2("BT" + drawingNo + "-AS-150-1@" + Path.GetFileNameWithoutExtension(mainAsm), "COMPONENT", 0, 0, 0, false, 0, null, 0);
                        Component2 swComp = (Component2)modifySelMgr.GetSelectedObject6(1, -1);
                        swComp.ReferencedConfiguration = "LHS - WITH SIGN";
                    }

                    if (gate == GateType.RhsSingleGate)
                    {
                        swModel.Extension.SelectByID2("BT" + drawingNo + "-AS-150-2@" + Path.GetFileNameWithoutExtension(mainAsm), "COMPONENT", 0, 0, 0, false, 0, null, 0);
                        Component2 swComp = (Component2)modifySelMgr.GetSelectedObject6(1, -1);
                        swComp.ReferencedConfiguration = "RHS - WITH SIGN";
                    }
                }

                // change linear pattern distance
                bSelRet = swModel.Extension.SelectByID2("JockeyStandLP", "COMPPATTERN", 0, 0, 0, false, 0, null, 0);
                selFeat = (Feature)modifySelMgr.GetSelectedObject6(1, -1);

                // change tread pattern instance
                modifyDim = selFeat.IParameter("Distance");
                modifyDim.SetSystemValue3(JockeyStandDis / 1000, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
            }
            else
            {
                if (isRemovable == true)
                {
                    if (gate == GateType.LhsSingleGate)
                    {
                        swModel.Extension.SelectByID2("BT" + drawingNo + "-AS-200-3@" + Path.GetFileNameWithoutExtension(mainAsm), "COMPONENT", 0, 0, 0, false, 0, null, 0);
                        Component2 swComp = (Component2)modifySelMgr.GetSelectedObject6(1, -1);
                        swComp.ReferencedConfiguration = "LHS";
                    }

                    if (gate == GateType.RhsSingleGate)
                    {
                        swModel.Extension.SelectByID2("BT" + drawingNo + "-AS-200-4@" + Path.GetFileNameWithoutExtension(mainAsm), "COMPONENT", 0, 0, 0, false, 0, null, 0);
                        Component2 swComp = (Component2)modifySelMgr.GetSelectedObject6(1, -1);
                        swComp.ReferencedConfiguration = "RHS";
                    }
                }
                else
                {
                    if (gate == GateType.LhsSingleGate)
                    {
                        swModel.Extension.SelectByID2("BT" + drawingNo + "-AS-150-1@" + Path.GetFileNameWithoutExtension(mainAsm), "COMPONENT", 0, 0, 0, false, 0, null, 0);
                        Component2 swComp = (Component2)modifySelMgr.GetSelectedObject6(1, -1);
                        swComp.ReferencedConfiguration = "LHS";
                    }

                    if (gate == GateType.RhsSingleGate)
                    {
                        swModel.Extension.SelectByID2("BT" + drawingNo + "-AS-150-2@" + Path.GetFileNameWithoutExtension(mainAsm), "COMPONENT", 0, 0, 0, false, 0, null, 0);
                        Component2 swComp = (Component2)modifySelMgr.GetSelectedObject6(1, -1);
                        swComp.ReferencedConfiguration = "RHS";
                    }
                }

                // clear selection first
                swModel.ClearSelection2(true);

                // select jockey stand
                bSelRet = swModel.Extension.SelectByID2("MANUTECH JOCKEY 70 STAND-1@" + Path.GetFileNameWithoutExtension(mainAsm), "COMPONENT", 0, 0, 0, false, 0, null, 0);

                // select jockey stand
                bSelRet = swModel.Extension.SelectByID2("JockeyStandLP", "COMPPATTERN", 0, 0, 0, true, 0, null, 0);

                // delete
                swModel.Extension.DeleteSelection2(deleteOption);
            }

            #endregion

            #region in context editing
            // cast model doc to assembly doc to access to edit part method
            AssemblyDoc swAsm = (AssemblyDoc)swModel;

            // select the frame to edit
            swModel.Extension.SelectByID2("BT" + drawingNo + "-PT-101-1@" + Path.GetFileNameWithoutExtension(mainAsm), "COMPONENT", 0, 0, 0, false, 0, null, 0);

            // go into in context editing
            swAsm.EditPart2(true, false, 1);

            // instantiate feature manager
            FeatureManager featMgr = swModel.FeatureManager;

            // get the selection set folder
            SelectionSetFolder selSetFolderFeat = (SelectionSetFolder)featMgr.GetSelectionSetFolder();

            // select first set for sketch creation
            SelectionSet selSet = (SelectionSet)selSetFolderFeat.GetSelectionSetByName("CW Base Frame");
            selSet.Select();

            // instantiate sketch manager
            SketchManager skMgr = swModel.SketchManager;

            // create new sketch
            skMgr.InsertSketch(true);

            // clear first selection
            swModel.ClearSelection2(true);

            // select second set for convert entity
            selSet = (SelectionSet)selSetFolderFeat.GetSelectionSetByName("CW Base Asm");
            selSet.Select();

            // convert entity
            skMgr.SketchUseEdge2(false);

            // create thin extrude cut feature
            featMgr.FeatureCutThin2(true, false, false, 0, 0, 0.0005, 0, false, false, false, false, 0, 0, false, false, false, false, 0.0001, 0, 0, 1, 0, false, 0, false, false, 0, 0, false);

            // exit in context editing
            swAsm.EditAssembly();

            // save and close assembly
            CloseDoc(swModel);
            #endregion

            #region Modify Drawing
            // Path to CL drawing
            modifyModelName = outputProjectString + "-AS-100_CL_A.SLDDRW";

            // Open CL drawing
            modifyModel = swApp.OpenDoc6(modifyModelName, (int)swDocumentTypes_e.swDocDRAWING, (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", ref err, ref warn);

            // Rebuild model
            modifyModel.ForceRebuild3(true);

            // save as pdf
            string CLPdfName = outputProjectString + "-AS-100_CL_A.PDF";
            bRet = modifyModel.Extension.SaveAs3(CLPdfName, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, null, ref err, ref warn);

            CloseDoc(modifyModel);

            // Path to Q drawing
            modifyModelName = outputProjectString + "-AS-100_Q_A.SLDDRW";

            // Open Q drawing
            modifyModel = swApp.OpenDoc6(modifyModelName, (int)swDocumentTypes_e.swDocDRAWING, (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", ref err, ref warn);

            // Rebuild model
            modifyModel.ForceRebuild3(true);

            // delete dimension
            modifySelMgr = modifyModel.ISelectionManager;
            modifyModel.ClearSelection2(true);
            switch (bitmask)
            {
                case 105:
                    {
                        if (hasJockeyStand == true)
                        {
                            modifyModel.Extension.SelectByID2("RD4@Drawing View3", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }
                        else
                        {
                            modifyModel.Extension.SelectByID2("RD5@Drawing View3", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        if (gate == GateType.BomRail)
                        {
                            modifyModel.Extension.SelectByID2("D4@Sketch27", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }
                        else
                        {
                            modifyModel.Extension.SelectByID2("D2@Sketch27", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        if (hasForklift == false)
                        {
                            modifyModel.Extension.SelectByID2("RD5@Drawing View2", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        break;
                    }
                case 106:
                    {
                        if (gate == GateType.BomRail)
                        {
                            modifyModel.Extension.SelectByID2("D3@Sketch45", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }
                        else
                        {
                            modifyModel.Extension.SelectByID2("D1@Sketch45", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        break;
                    }
                case 89:
                    {
                        if (gate == GateType.BomRail)
                        {
                            modifyModel.Extension.SelectByID2("RD5@Drawing View4", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }
                        else
                        {
                            modifyModel.Extension.SelectByID2("RD4@Drawing View4", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        if (hasForklift == false)
                        {
                            modifyModel.Extension.SelectByID2("RD5@Drawing View2", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        break;
                    }
                case 90:
                    {
                        if (gate == GateType.BomRail)
                        {
                            modifyModel.Extension.SelectByID2("RD4@Drawing View4", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }
                        else
                        {
                            modifyModel.Extension.SelectByID2("RD3@Drawing View4", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        break;
                    }
                case 169:
                    {
                        if (hasJockeyStand == true)
                        {
                            modifyModel.Extension.SelectByID2("RD4@Drawing View3", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }
                        else
                        {
                            modifyModel.Extension.SelectByID2("RD3@Drawing View3", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        if (gate == GateType.BomRail)
                        {
                            modifyModel.Extension.SelectByID2("D3@Sketch45", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }
                        else
                        {
                            modifyModel.Extension.SelectByID2("D2@Sketch45", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        if (hasForklift == false)
                        {
                            modifyModel.Extension.SelectByID2("RD5@Drawing View2", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        break;
                    }
                case 170:
                    {
                        if (gate == GateType.BomRail)
                        {
                            modifyModel.Extension.SelectByID2("D3@Sketch25", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }
                        else
                        {
                            modifyModel.Extension.SelectByID2("D2@Sketch25", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }
                        break;
                    }
                case 153:
                    {
                        if (gate == GateType.BomRail)
                        {
                            modifyModel.Extension.SelectByID2("RD5@Drawing View4", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }
                        else
                        {
                            modifyModel.Extension.SelectByID2("RD3@Drawing View4", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        if (hasForklift == false)
                        {
                            modifyModel.Extension.SelectByID2("RD4@Drawing View2", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        break;
                    }
                case 154:
                    {
                        if (gate == GateType.BomRail)
                        {
                            modifyModel.Extension.SelectByID2("RD4@Drawing View4", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }
                        else
                        {
                            modifyModel.Extension.SelectByID2("RD3@Drawing View4", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        break;
                    }
                case 553:
                    {
                        if (hasJockeyStand == true)
                        {
                            modifyModel.Extension.SelectByID2("RD4@Drawing View3", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }
                        else
                        {
                            modifyModel.Extension.SelectByID2("RD1@Drawing View3", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        if (gate == GateType.BomRail)
                        {
                            modifyModel.Extension.SelectByID2("D3@Sketch27", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }
                        else
                        {
                            modifyModel.Extension.SelectByID2("D2@Sketch27", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        if (hasForklift == false)
                        {
                            modifyModel.Extension.SelectByID2("RD6@Drawing View2", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        break;
                    }
                case 554:
                    {
                        if (gate == GateType.BomRail)
                        {
                            modifyModel.Extension.SelectByID2("D3@Sketch45", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }
                        else
                        {
                            modifyModel.Extension.SelectByID2("D2@Sketch45", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        break;
                    }
                case 537:
                    {
                        if (gate == GateType.BomRail)
                        {
                            modifyModel.Extension.SelectByID2("RD5@Drawing View4", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }
                        else
                        {
                            modifyModel.Extension.SelectByID2("RD4@Drawing View4", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        if (hasForklift == false)
                        {
                            modifyModel.Extension.SelectByID2("RD5@Drawing View2", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        break;
                    }
                case 538:
                    {
                        if (gate == GateType.BomRail)
                        {
                            modifyModel.Extension.SelectByID2("RD4@Drawing View4", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }
                        else
                        {
                            modifyModel.Extension.SelectByID2("RD3@Drawing View4", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        break;
                    }
                case 281:
                    {
                        if (gate == GateType.BomRail)
                        {
                            modifyModel.Extension.SelectByID2("RD5@Drawing View4", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }
                        else
                        {
                            modifyModel.Extension.SelectByID2("RD3@Drawing View4", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        if (hasForklift == false)
                        {
                            modifyModel.Extension.SelectByID2("RD5@Drawing View2", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        break;
                    }
                case 282:
                    {
                        if (gate == GateType.BomRail)
                        {
                            modifyModel.Extension.SelectByID2("RD5@Drawing View4", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }
                        else
                        {
                            modifyModel.Extension.SelectByID2("RD2@Drawing View4", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        break;
                    }
                case 297:
                    {
                        if (hasJockeyStand == true)
                        {
                            modifyModel.Extension.SelectByID2("RD2@Drawing View3", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }
                        else
                        {
                            modifyModel.Extension.SelectByID2("RD1@Drawing View3", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        if (gate == GateType.BomRail)
                        {
                            modifyModel.Extension.SelectByID2("D3@Sketch26", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }
                        else
                        {
                            modifyModel.Extension.SelectByID2("D2@Sketch26", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        if (hasForklift == false)
                        {
                            modifyModel.Extension.SelectByID2("RD6@Drawing View2", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        break;
                    }
                case 298:
                    {
                        if (gate == GateType.BomRail)
                        {
                            modifyModel.Extension.SelectByID2("D3@Sketch45", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }
                        else
                        {
                            modifyModel.Extension.SelectByID2("D2@Sketch45", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        break;
                    }
                case 101:
                    {
                        if (hasJockeyStand == true)
                        {
                            modifyModel.Extension.SelectByID2("RD3@Drawing View3", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }
                        else
                        {
                            modifyModel.Extension.SelectByID2("RD2@Drawing View3", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        if (gate == GateType.BomRail)
                        {
                            modifyModel.Extension.SelectByID2("D3@Sketch27", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }
                        else
                        {
                            modifyModel.Extension.SelectByID2("D1@Sketch27", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        if (hasForklift == false)
                        {
                            modifyModel.Extension.SelectByID2("RD2@Drawing View2", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        break;
                    }
                case 102:
                    {
                        if (gate == GateType.BomRail)
                        {
                            modifyModel.Extension.SelectByID2("D3@Sketch45", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }
                        else
                        {
                            modifyModel.Extension.SelectByID2("D2@Sketch45", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        break;
                    }
                case 85:
                    {
                        if (gate == GateType.BomRail)
                        {
                            modifyModel.Extension.SelectByID2("RD5@Drawing View4", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }
                        else
                        {
                            modifyModel.Extension.SelectByID2("RD4@Drawing View4", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        if (hasForklift == false)
                        {
                            modifyModel.Extension.SelectByID2("RD6@Drawing View2", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        break;
                    }
                case 86:
                    {
                        if (gate == GateType.BomRail)
                        {
                            modifyModel.Extension.SelectByID2("RD3@Drawing View4", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }
                        else
                        {
                            modifyModel.Extension.SelectByID2("RD5@Drawing View4", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        break;
                    }
                case 165:
                    {
                        if (hasJockeyStand == true)
                        {
                            modifyModel.Extension.SelectByID2("RD3@Drawing View3", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }
                        else
                        {
                            modifyModel.Extension.SelectByID2("RD2@Drawing View3", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        if (gate == GateType.BomRail)
                        {
                            modifyModel.Extension.SelectByID2("D3@Sketch45", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }
                        else
                        {
                            modifyModel.Extension.SelectByID2("D2@Sketch45", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        if (hasForklift == false)
                        {
                            modifyModel.Extension.SelectByID2("RD7@Drawing View2", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        break;
                    }
                case 166:
                    {
                        if (gate == GateType.BomRail)
                        {
                            modifyModel.Extension.SelectByID2("D3@Sketch45", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }
                        else
                        {
                            modifyModel.Extension.SelectByID2("D2@Sketch45", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        break;
                    }
                case 149:
                    {
                        if (gate == GateType.BomRail)
                        {
                            modifyModel.Extension.SelectByID2("RD5@Drawing View4", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }
                        else
                        {
                            modifyModel.Extension.SelectByID2("RD4@Drawing View4", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        if (hasForklift == false)
                        {
                            modifyModel.Extension.SelectByID2("RD4@Drawing View2", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        break;
                    }
                case 150:
                    {
                        if (gate == GateType.BomRail)
                        {
                            modifyModel.Extension.SelectByID2("RD5@Drawing View4", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }
                        else
                        {
                            modifyModel.Extension.SelectByID2("RD4@Drawing View4", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        break;
                    }
                case 549:
                    {
                        if (hasJockeyStand == true)
                        {
                            modifyModel.Extension.SelectByID2("RD3@Drawing View3", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }
                        else
                        {
                            modifyModel.Extension.SelectByID2("RD2@Drawing View3", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        if (gate == GateType.BomRail)
                        {
                            modifyModel.Extension.SelectByID2("D3@Sketch45", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }
                        else
                        {
                            modifyModel.Extension.SelectByID2("D1@Sketch45", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        if (hasForklift == false)
                        {
                            modifyModel.Extension.SelectByID2("RD7@Drawing View2", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        break;
                    }
                case 550:
                    {
                        if (gate == GateType.BomRail)
                        {
                            modifyModel.Extension.SelectByID2("D3@Sketch45", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }
                        else
                        {
                            modifyModel.Extension.SelectByID2("D2@Sketch45", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        break;
                    }
                case 533:
                    {
                        if (gate == GateType.BomRail)
                        {
                            modifyModel.Extension.SelectByID2("RD5@Drawing View4", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }
                        else
                        {
                            modifyModel.Extension.SelectByID2("RD4@Drawing View2", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        if (hasForklift == false)
                        {
                            modifyModel.Extension.SelectByID2("RD5@Drawing View2", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        break;
                    }
                case 534:
                    {
                        if (gate == GateType.BomRail)
                        {
                            modifyModel.Extension.SelectByID2("RD5@Drawing View4", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }
                        else
                        {
                            modifyModel.Extension.SelectByID2("RD2@Drawing View4", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        break;
                    }
                case 277:
                    {
                        if (gate == GateType.BomRail)
                        {
                            modifyModel.Extension.SelectByID2("RD5@Drawing View4", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }
                        else
                        {
                            modifyModel.Extension.SelectByID2("RD2@Drawing View4", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        if (hasForklift == false)
                        {
                            modifyModel.Extension.SelectByID2("RD5@Drawing View2", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        break;
                    }
                case 278:
                    {
                        if (gate == GateType.BomRail)
                        {
                            modifyModel.Extension.SelectByID2("RD5@Drawing View4", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }
                        else
                        {
                            modifyModel.Extension.SelectByID2("RD2@Drawing View4", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        break;
                    }
                case 293:
                    {
                        if (hasJockeyStand == true)
                        {
                            modifyModel.Extension.SelectByID2("RD2@Drawing View3", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }
                        else
                        {
                            modifyModel.Extension.SelectByID2("RD1@Drawing View3", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        if (gate == GateType.BomRail)
                        {
                            modifyModel.Extension.SelectByID2("D3@Sketch45", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }
                        else
                        {
                            modifyModel.Extension.SelectByID2("D2@Sketch45", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        if (hasForklift == false)
                        {
                            modifyModel.Extension.SelectByID2("RD6@Drawing View2", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        break;
                    }
                case 294:
                    {
                        if (gate == GateType.BomRail)
                        {
                            modifyModel.Extension.SelectByID2("D3@Sketch29", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }
                        else
                        {
                            modifyModel.Extension.SelectByID2("D2@Sketch29", "DIMENSION", 0, 0, 0, true, 0, null, 0);
                        }

                        break;
                    }
                default: { MessageBox.Show("Bitmask operation failed."); return; }
            }

            // delete selected dimension
            modifyModel.Extension.DeleteSelection2(deleteOption);

            // save as pdf
            string QPdfName = outputProjectString + "-AS-100_Q_A.PDF";
            bRet = modifyModel.Extension.SaveAs3(QPdfName, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, null, ref err, ref warn);

            CloseDoc(modifyModel);
            #endregion

            Close();

            sb.AppendLine("Model Generated!");
            MessageBox.Show(sb.ToString());
        }

        private void OutputLocBtn_Click(object sender, RoutedEventArgs e)
        {
            string folderPath = "";
            System.Windows.Forms.FolderBrowserDialog fbd = new System.Windows.Forms.FolderBrowserDialog();
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                folderPath = fbd.SelectedPath;
                OutputLocTb.Text = folderPath;
                appSet.PATargetPath = folderPath;
                appSet.Save();
            }
        }

        private void MasterLocBtn_Click(object sender, RoutedEventArgs e)
        {
            string folderPath = "";
            System.Windows.Forms.FolderBrowserDialog fbd = new System.Windows.Forms.FolderBrowserDialog();
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                folderPath = fbd.SelectedPath;
                MasterLocTb.Text = folderPath;
                appSet.PAMasterPath = folderPath;
                appSet.Save();
            }
        }

        private void TypeCb_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBox comboBox = (ComboBox)sender;
            if (comboBox.SelectedIndex == 0)
            {
                // if selected is stairways,
                // enable jockey stand option
                if (FrameCb.SelectedIndex != 0)
                {
                    JockeyStandCh.IsEnabled = true;
                }

                // enable forklift option
                ForkliftCh.IsEnabled = true;
            }
            else
            {
                // if selected is step type,
                // enable jockey stand option
                JockeyStandCh.IsChecked = false;
                JockeyStandCh.IsEnabled = false;

                // enable forklift option
                ForkliftCh.IsChecked = false;
                ForkliftCh.IsEnabled = false;
            }
        }

        private void OverhangCh_Checked(object sender, RoutedEventArgs e)
        {
            OverhangLengthTb.IsEnabled = true;
        }

        private void OverhangCh_Unchecked(object sender, RoutedEventArgs e)
        {
            OverhangLengthTb.IsEnabled = false;
        }

        private void FrameCb_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBox comboBox = (ComboBox)sender;
            if (comboBox.SelectedIndex == 0)
            {
                JockeyStandCh.IsChecked = false;
                JockeyStandCh.IsEnabled = false;
            }
            else
            {
                if (TypeCb.SelectedIndex == 0)
                {
                    JockeyStandCh.IsEnabled = true;
                }
            }
        }

        private void OnGotKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            var textBox = sender as TextBox;
            if (textBox != null && !textBox.IsReadOnly && e.KeyboardDevice.IsKeyDown(Key.Tab))
                textBox.SelectAll();
        }

        // method to close doc
        private void CloseDoc(ModelDoc2 swmodel)
        {
            int err = -1;
            int warn = -1;
            swmodel.EditRebuild3();
            swmodel.ClearSelection2(true);
            swmodel.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref err, ref warn);
            swApp.CloseDoc(Path.GetFileNameWithoutExtension(swmodel.GetPathName()));
        }

        // method to change plate dimension
        private void ChangeFeatureDimension(Feature selectedFeature, string parameterName, double value)
        {
            // change tread pattern distance
            Dimension modifyDim = selectedFeature.IParameter(parameterName);
            if (modifyDim != null)
            {
                modifyDim.SetSystemValue3(value, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
            }
        }

        // method to change mate dimension
        private void ChangeMateDimension(ModelDoc2 swModel, SelectionMgr swSelMgr, string mateName, double value)
        {
            bool bRet = swModel.Extension.SelectByID2(mateName, "MATE", 0, 0, 0, false, 0, null, 0);
            if (bRet)
            {
                Feature swFeat = (Feature)swSelMgr.GetSelectedObject6(1, -1);
                Mate2 swMate = (Mate2)swFeat.GetSpecificFeature2();
                DisplayDimension swDispDim = swMate.DisplayDimension2[0];
                Dimension swDim = swDispDim.IGetDimension();
                swDim.SetSystemValue3(value, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
            }
        }
    }
}
