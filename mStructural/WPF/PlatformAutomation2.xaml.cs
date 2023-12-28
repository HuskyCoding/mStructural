using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using static System.Net.WebRequestMethods;

namespace mStructural.WPF
{
    /// <summary>
    /// Interaction logic for PlatformAutomation2.xaml
    /// </summary>
    public partial class PlatformAutomation2 : Window
    {
        #region Private Variables
        SldWorks swApp;
        #endregion

        // Constructor
        public PlatformAutomation2(SldWorks swapp)
        {
            swApp = swapp;
            InitializeComponent();
        }

        #region PA1
        private void pa1GenerateBtn_Click(object sender, RoutedEventArgs e)
        {
            #region Check Input
            double PlatformHeight;
            double PlatformAngle;
            double PlatformWidth;
            double LandingLength;
            int BTNumber;
            string OutputProjectString;
            bool bRet = false;
            string WJConfig;
            string ForkConfig;
            string GateConfig;

            // check if any model is open
            if (swApp.IActiveDoc2 != null)
            {
                MessageBox.Show("Please close all active document.");
                return;
            }

            // get height
            bRet = double.TryParse(pa1HeightTb.Text, out PlatformHeight);
            if (!bRet) { MessageBox.Show("Height input is invalid.");return; }

            // get angle
            bRet = double.TryParse(pa1AngleTb.Text, out PlatformAngle);
            if (!bRet) { MessageBox.Show("Angle input is invalid."); return; }

            // get width
            bRet = double.TryParse(pa1WidthTb.Text, out PlatformWidth);
            if (!bRet) { MessageBox.Show("Width input is invalid."); return; }

            // get landing length
            bRet = double.TryParse(pa1LandingLengthTb.Text, out LandingLength);
            if (!bRet) { MessageBox.Show("Landing Length input is invalid."); return; }

            // check output directory
            bRet = Directory.Exists(pa1OutputLocTb.Text);
            if (!bRet) { MessageBox.Show("Ouput location does not exist."); return; }

            // check BT number
            // check length
            if (pa1DrawingNoTb.Text.Length != 5) { MessageBox.Show("BT Number length should be exactly 5"); return; }

            // check if all are digits
            bRet = int.TryParse(pa1DrawingNoTb.Text, out BTNumber);
            if (!bRet) { MessageBox.Show("Drawing number is invalid, only digits (0~9) are allowed."); return; }

            WJConfig = pa1WJConfigCb.Text;
            ForkConfig = pa1ForkConfigCb.Text;
            GateConfig = pa1GateConfigCb.Text;
            OutputProjectString = pa1OutputLocTb.Text + "\\BT" + BTNumber.ToString();
            #endregion

            #region Calculate
            double PlatformAngleRad;
            double TreadHorizontalLength;
            double HandrailOffest;
            int TreadInstance;
            double TreadRiser;
            double TreadHypotenuse;
            double TreadGoing;
            double RiserNGoing;
            double TotalHandrailLength;
            double HandrailPostDistance;
            double HandrailPlateDistance;
            double FirstTreadPlateDistance;
            double TreadDim;
            double JockeyStandDis;
            double WheelDis;

            PlatformAngleRad = PlatformAngle * Math.PI / 180;
            TreadHorizontalLength = PlatformHeight / Math.Tan(PlatformAngleRad);
            HandrailOffest = 190;

            #region Tread Instance Calculation
            TreadInstance = 2;
            TreadRiser = PlatformHeight / TreadInstance;
            while (TreadRiser>225)
            { 
                TreadInstance++;
                TreadRiser = PlatformHeight / TreadInstance;
            }

            if (TreadRiser>=130)
            {
                TreadHypotenuse = TreadRiser / Math.Sin(PlatformAngleRad);
                TreadGoing = Math.Sqrt(TreadHypotenuse * TreadHypotenuse - TreadRiser * TreadRiser);
                while (TreadGoing > 335)
                {
                    TreadInstance++;
                    TreadRiser = PlatformHeight / TreadInstance;
                    if (TreadRiser >= 130)
                    {
                        TreadHypotenuse = TreadRiser / Math.Sin(PlatformAngleRad);
                        TreadGoing = Math.Sqrt(TreadHypotenuse * TreadHypotenuse - TreadRiser * TreadRiser);
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
                MessageBox.Show("Tread calculation failed");
                return;
            }

            if (TreadGoing>=215)
            {
                RiserNGoing = 2 * TreadRiser + TreadGoing;
                while (RiserNGoing > 700)
                {
                    TreadInstance++;
                    TreadRiser = PlatformHeight / TreadInstance;
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
                MessageBox.Show("Tread calculation failed");
                return;
            }

            if (RiserNGoing < 215)
            {
                MessageBox.Show("Tread calculation failed.");
                return;
            }
            #endregion

            TotalHandrailLength = TreadHorizontalLength + HandrailOffest - TreadGoing;
            HandrailPostDistance = TotalHandrailLength / 2;
            HandrailPlateDistance = HandrailPostDistance/Math.Cos(PlatformAngleRad);
            TreadDim = 105+20*Math.Tan(PlatformAngleRad);
            FirstTreadPlateDistance = ((HandrailPostDistance - HandrailOffest) / Math.Cos(PlatformAngleRad)) - TreadDim;
            JockeyStandDis = LandingLength + 595;
            WheelDis = LandingLength + 1175;
            #endregion

            #region Pack and go model
            // pack and go model
            string[] allFiles = Directory.GetFiles(@"D:\Programming\MasterPlatform\PlatformAutomation2\PA00001");
            List<string> asmList = new List<string>();
            List<Tuple<string,string>> nameList = new List<Tuple<string,string>>();
            foreach(string file in allFiles)
            {
                // exclude temp files
                if (!file.Contains("~"))
                {
                    //copy files
                    string sourceFile = file;
                    string oriName = Path.GetFileName(sourceFile);
                    string newName = oriName.Replace("PA00001", "BT" + BTNumber.ToString());
                    string destFile = pa1OutputLocTb.Text + "\\" + newName;
                    int iRet = swApp.CopyDocument(sourceFile, destFile, "", "", (int)swMoveCopyOptions_e.swMoveCopyOptionsOverwriteExistingDocs);

                    // set read-only flag to off for newly copied files
                    FileInfo fileInfo = new FileInfo(destFile);
                    fileInfo.IsReadOnly = false;

                    // add to respective list
                    nameList.Add(new Tuple<string, string>(sourceFile, destFile));
                    if (destFile.ToUpper().Contains("SLDASM"))
                        asmList.Add(destFile);
                }
            }

            // update reference
            foreach(string asm in asmList)
            {
                for(int i = 0; i < nameList.Count; i++)
                {
                    string sourceFile = nameList[i].Item1;
                    string destFile = nameList[i].Item2;
                    bRet = swApp.ReplaceReferencedDocument(asm,sourceFile, destFile);
                }
            }

            // open model
            string mainAsm = OutputProjectString + "-AS-100.SLDASM";
            int err = -1;
            int warn = -1;
            bool bSelRet;
            ModelDoc2 swModel = swApp.OpenDoc6(mainAsm, (int)swDocumentTypes_e.swDocASSEMBLY, (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", ref err, ref warn);
            #endregion

            #region Modify Frame
            string modifyModelName;
            ModelDoc2 modifyModel;
            SelectionMgr modifySelMgr;
            Feature selFeat;
            Dimension modifyDim;
            EquationMgr modifyEqMgr;

            int deleteOption = (int)swDeleteSelectionOptions_e.swDelete_Children + (int)swDeleteSelectionOptions_e.swDelete_Absorbed;

            modifyModelName = OutputProjectString + "-PT-101.SLDPRT";
            modifyModel = swApp.OpenDoc6(modifyModelName, (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", ref err, ref warn);
            modifySelMgr = modifyModel.ISelectionManager;
            
            // select main frame sketch
            bSelRet = modifyModel.Extension.SelectByID2("Main Frame Sketch", "SKETCH", 0, 0, 0, false, 0, null, 0);
            selFeat = (Feature)modifySelMgr.GetSelectedObject6(1, -1);
            
            // change height dimension
            modifyDim = selFeat.IParameter("Height");
            modifyDim.SetSystemValue3(PlatformHeight/1000, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

            // change width dimension
            modifyDim = selFeat.IParameter ("Width");
            modifyDim.SetSystemValue3(PlatformWidth/1000, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

            // change angle dimension
            modifyDim = selFeat.IParameter("Angle");
            modifyDim.SetSystemValue3(PlatformAngleRad, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

            // change landing length dimension
            modifyDim = selFeat.IParameter("LandingLength");
            modifyDim.SetSystemValue3((LandingLength-25) / 1000, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

            // select handrail sketch
            bSelRet = modifyModel.Extension.SelectByID2("Handrail Sketch", "SKETCH", 0, 0, 0, false, 0, null, 0);
            selFeat = (Feature)modifySelMgr.GetSelectedObject6(1, -1);

            // change handrail first post dimension
            modifyDim = selFeat.IParameter("HrPost1");
            modifyDim.SetSystemValue3(HandrailPostDistance / 1000, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

            // change handrail second post dimension
            modifyDim = selFeat.IParameter("HrPost2");
            modifyDim.SetSystemValue3(HandrailPostDistance / 1000, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

            // select first tread etch sketch
            bSelRet = modifyModel.Extension.SelectByID2("First Tread Etch Sketch", "SKETCH", 0, 0, 0, false, 0, null, 0);
            selFeat = (Feature)modifySelMgr.GetSelectedObject6(1, -1);

            // change first tread etch distance
            modifyDim = selFeat.IParameter("Distance");
            modifyDim.SetSystemValue3(TreadHypotenuse/1000, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

            // select tread etch pattern
            bSelRet = modifyModel.Extension.SelectByID2("Tread Etch", "BODYFEATURE", 0, 0, 0, false, 0, null, 0);
            selFeat = (Feature)modifySelMgr.GetSelectedObject6(1, -1);

            // change tread pattern instance
            modifyDim = selFeat.IParameter("Instance");
            modifyDim.SetSystemValue3(Convert.ToDouble(TreadInstance-1), (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

            // change tread pattern distance
            modifyDim = selFeat.IParameter("Distance");
            modifyDim.SetSystemValue3(TreadHypotenuse / 1000, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

            #region Conditional
            switch (WJConfig)
            {
                case "Wheel and Jockey Stand":
                    {
                        // nothing will change
                        break;
                    }
                case "Wheel":
                    {
                        // select the jack adaptor plate and delete
                        bSelRet = modifyModel.Extension.SelectByID2("PL10 250 x 170 AL JACK ADAPTOR", "BODYFEATURE", 0, 0, 0, false, 0, null, 0);
                        modifyModel.Extension.DeleteSelection2(deleteOption);
                        break;
                    }
            }

            switch (ForkConfig)
            {
                case "Yes":
                    {
                        // do nothing
                        break;
                    }
                case "No":
                    {
                        // select the forklife main sketch and delete
                        bSelRet = modifyModel.Extension.SelectByID2("Forklift Sketch", "SKETCH", 0, 0, 0, false, 0, null, 0);
                        modifyModel.Extension.DeleteSelection2(deleteOption);
                        break;
                    }
            }
            #endregion

            CloseDoc(modifyModel);
            #endregion

            #region Modify Landing Grate
            modifyModelName = OutputProjectString + "-PT-102.SLDPRT";
            modifyModel = swApp.OpenDoc6(modifyModelName, (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", ref err, ref warn);
            modifyEqMgr = modifyModel.GetEquationMgr();

            // change equation
            modifyEqMgr.SetEquationAndConfigurationOption(2, "\"\"Span\"\" = " + ((PlatformWidth-100)).ToString(), (int)swInConfigurationOpts_e.swThisConfiguration, null);
            modifyEqMgr.SetEquationAndConfigurationOption(3, "\"\"Width\"\" = " + ((LandingLength-100)).ToString(), (int)swInConfigurationOpts_e.swThisConfiguration, null);

            CloseDoc(modifyModel);
            #endregion

            #region Modify Stair Tread
            modifyModelName = OutputProjectString + "-PT-103.SLDPRT";
            modifyModel = swApp.OpenDoc6(modifyModelName, (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", ref err, ref warn);
            modifyEqMgr = modifyModel.GetEquationMgr();

            // change equation
            modifyEqMgr.SetEquationAndConfigurationOption(2, "\"\"Span\"\" = " + (PlatformWidth - 100).ToString(), (int)swInConfigurationOpts_e.swThisConfiguration, null);

            CloseDoc(modifyModel);
            #endregion

            #region Process Gate
            switch (GateConfig)
            {
                case "None":
                    {
                        // do nothing
                        break;
                    }
                case "LHS Single Gate":
                    {
                        // change configuration
                        modifyModelName = OutputProjectString + "-AS-150.SLDASM";
                        modifyModel = swApp.OpenDoc6(modifyModelName, (int)swDocumentTypes_e.swDocASSEMBLY, (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", ref err, ref warn);

                        // change equation
                        modifyEqMgr = modifyModel.GetEquationMgr();
                        modifyEqMgr.SetEquationAndConfigurationOption(0, "\"\"ENTRANCE WIDTH\"\" = " + (PlatformWidth - 50).ToString(), (int)swInConfigurationOpts_e.swAllConfiguration, null);
                        CloseDoc(modifyModel);
                        break;
                    }
                case "RHS Single Gate":
                    {
                        // change configuration
                        modifyModelName = OutputProjectString + "-AS-150.SLDASM";
                        modifyModel = swApp.OpenDoc6(modifyModelName, (int)swDocumentTypes_e.swDocASSEMBLY, (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", ref err, ref warn);

                        // change equation
                        modifyEqMgr = modifyModel.GetEquationMgr();
                        modifyEqMgr.SetEquationAndConfigurationOption(0, "\"\"ENTRANCE WIDTH\"\" = " + (PlatformWidth - 50).ToString(), (int)swInConfigurationOpts_e.swAllConfiguration, null);
                        CloseDoc(modifyModel);
                        break;
                    }
                case "Double Gate":
                    {
                        // change equation
                        modifyModelName = OutputProjectString + "-AS-250.SLDASM";
                        modifyModel = swApp.OpenDoc6(modifyModelName, (int)swDocumentTypes_e.swDocASSEMBLY, (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", ref err, ref warn);
                        modifyEqMgr = modifyModel.GetEquationMgr();
                        modifyEqMgr.SetEquationAndConfigurationOption(0, "\"\"GATE OPENING\"\" = " + (PlatformWidth - 50).ToString(), (int)swInConfigurationOpts_e.swThisConfiguration, null);
                        CloseDoc(modifyModel);
                        break;
                    }
                case "Bom Rail":
                    {
                        // change equation
                        modifyModelName = OutputProjectString + "-AS-200.SLDASM";
                        modifyModel = swApp.OpenDoc6(modifyModelName, (int)swDocumentTypes_e.swDocASSEMBLY, (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", ref err, ref warn);
                        modifyEqMgr = modifyModel.GetEquationMgr();
                        modifyEqMgr.SetEquationAndConfigurationOption(0, "\"\"GATE OPENING\"\" = " + (PlatformWidth - 50).ToString(), (int)swInConfigurationOpts_e.swThisConfiguration, null);
                        CloseDoc(modifyModel);
                        break;
                    }
            }
            #endregion

            #region Modify Tread Pattern
            // activate main assembly
            swModel = (ModelDoc2)swApp.ActivateDoc3(Path.GetFileName(mainAsm), false, (int)swRebuildOptions_e.swRebuildAll, ref err);

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
            // delete unrelated assembly
            if (GateConfig != "LHS Single Gate")
            {
                swModel.Extension.SelectByID2("BT"+BTNumber+"-AS-150-1@" + Path.GetFileNameWithoutExtension(mainAsm), "COMPONENT", 0, 0, 0, false, 0, null, 0);
                swModel.Extension.DeleteSelection2(deleteOption);
            }

            if(GateConfig!="RHS Single Gate")
            {
                swModel.Extension.SelectByID2("BT" + BTNumber + "-AS-150-2@" + Path.GetFileNameWithoutExtension(mainAsm), "COMPONENT", 0, 0, 0, false, 0, null, 0);
                swModel.Extension.DeleteSelection2(deleteOption);
            }

            if(GateConfig!= "Bom Rail")
            {
                swModel.Extension.SelectByID2("BT" + BTNumber+"-AS-200-2@" + Path.GetFileNameWithoutExtension(mainAsm), "COMPONENT", 0, 0, 0, false, 0, null, 0);
                swModel.Extension.DeleteSelection2(deleteOption);
            }

            if(GateConfig!= "Double Gate")
            {
                swModel.Extension.SelectByID2("BT" + BTNumber+"-AS-250-2@" + Path.GetFileNameWithoutExtension(mainAsm), "COMPONENT", 0, 0, 0, false, 0, null, 0);
                swModel.Extension.SelectByID2("BT" + BTNumber+"-AS-250-3@" + Path.GetFileNameWithoutExtension(mainAsm), "COMPONENT", 0, 0, 0, true, 0, null, 0);
                swModel.Extension.DeleteSelection2(deleteOption);
            }

            switch (WJConfig)
            {
                case "Wheel and Jockey Stand":
                    {
                        // change configuration
                        if(GateConfig == "LHS Single Gate")
                        {
                            swModel.Extension.SelectByID2("BT"+BTNumber+"-AS-150-1@" + Path.GetFileNameWithoutExtension(mainAsm), "COMPONENT", 0, 0, 0, false, 0, null, 0);
                            Component2 swComp = (Component2)modifySelMgr.GetSelectedObject6(1, -1);
                            swComp.ReferencedConfiguration = "LHS - WITH SIGN";
                        }

                        if (GateConfig == "RHS Single Gate")
                        {
                            swModel.Extension.SelectByID2("BT" + BTNumber + "-AS-150-2@" + Path.GetFileNameWithoutExtension(mainAsm), "COMPONENT", 0, 0, 0, false, 0, null, 0);
                            Component2 swComp = (Component2)modifySelMgr.GetSelectedObject6(1, -1);
                            swComp.ReferencedConfiguration = "RHS - WITH SIGN";
                        }

                        // change linear pattern distance
                        bSelRet = swModel.Extension.SelectByID2("JockeyStandLP", "COMPPATTERN", 0, 0, 0, false, 0, null, 0);
                        selFeat = (Feature)modifySelMgr.GetSelectedObject6(1, -1);

                        // change tread pattern instance
                        modifyDim = selFeat.IParameter("Distance");
                        modifyDim.SetSystemValue3(JockeyStandDis/1000, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                        // change linear pattern distance
                        bSelRet = swModel.Extension.SelectByID2("WheelLP", "COMPPATTERN", 0, 0, 0, false, 0, null, 0);
                        selFeat = (Feature)modifySelMgr.GetSelectedObject6(1, -1);

                        // change tread pattern instance
                        modifyDim = selFeat.IParameter("Distance");
                        modifyDim.SetSystemValue3(WheelDis / 1000, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                        break;
                    }

                case "Wheel":
                    {
                        if (GateConfig == "LHS Single Gate")
                        {
                            swModel.Extension.SelectByID2("BT" + BTNumber + "-AS-150-1@" + Path.GetFileNameWithoutExtension(mainAsm), "COMPONENT", 0, 0, 0, false, 0, null, 0);
                            Component2 swComp = (Component2)modifySelMgr.GetSelectedObject6(1, -1);
                            swComp.ReferencedConfiguration = "LHS";
                        }

                        if (GateConfig == "RHS Single Gate")
                        {
                            swModel.Extension.SelectByID2("BT" + BTNumber + "-AS-150-2@" + Path.GetFileNameWithoutExtension(mainAsm), "COMPONENT", 0, 0, 0, false, 0, null, 0);
                            Component2 swComp = (Component2)modifySelMgr.GetSelectedObject6(1, -1);
                            swComp.ReferencedConfiguration = "RHS";
                        }

                        // select jockey stand
                        bSelRet = swModel.Extension.SelectByID2("MANUTECH JOCKEY 70 STAND-1@"+Path.GetFileNameWithoutExtension(mainAsm), "COMPONENT", 0, 0, 0, false, 0, null, 0);

                        // select jockey stand
                        bSelRet = swModel.Extension.SelectByID2("JockeyStandLP", "COMPPATTERN", 0, 0, 0, true, 0, null, 0);

                        // suppress
                        swModel.Extension.DeleteSelection2(deleteOption);

                        break;
                    }
            }
            #endregion

            #region in context editing
            // cast model doc to assembly doc to access to edit part method
            AssemblyDoc swAsm = (AssemblyDoc)swModel;

            // select the frame to edit
            swModel.Extension.SelectByID2("BT"+BTNumber+ "-PT-101-1@" + Path.GetFileNameWithoutExtension(mainAsm), "COMPONENT", 0, 0, 0, false, 0, null, 0);
            
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

            #endregion

            swModel.ClearSelection2(true);
            swModel.EditRebuild3();
            swModel.ShowNamedView2("*Isometric", -1);
            swModel.ViewZoomtofit2();
            Close();
            MessageBox.Show("Model generation completed.");
        }

        private void pa1OutputLocBrowseBtn_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog fbd = new System.Windows.Forms.FolderBrowserDialog();
            System.Windows.Forms.DialogResult result = fbd.ShowDialog();
            if(result == System.Windows.Forms.DialogResult.OK)
            {
                pa1OutputLocTb.Text = fbd.SelectedPath;
                System.Environment.SpecialFolder root = fbd.RootFolder;
            }
        }
        #endregion

        // method to close doc
        private void CloseDoc(ModelDoc2 swmodel)
        {
            int err = -1;
            int warn = -1;
            swmodel.EditRebuild3();
            swmodel.ClearSelection2(true);
            swmodel.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref err, ref warn);
        }
    }
}
