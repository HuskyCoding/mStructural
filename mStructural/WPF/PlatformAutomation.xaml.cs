using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Media.Imaging;

namespace mStructural.WPF
{
    /// <summary>
    /// Interaction logic for PlatformAutomation.xaml
    /// </summary>
    public partial class PlatformAutomation : Window
    {
        #region Private Member
        private SldWorks swApp;
        private appsetting appset;

        private enum PlatformType
        {
            OverhangLeftOpen
        }
        private PlatformType platformType;

        private enum HandrailType
        {
            Bolton,
            Welded,
            RemovableBolton,
            RemovableWelded
        }
        private HandrailType handrailType;

        private enum Configuration
        {
            WheelAndJockey,
            WheelOnly,
            Fixed
        }
        private Configuration config;

        private string configurationName;
        private string configName;
        private string handrailName;
        
        #endregion

        // Constructor
        public PlatformAutomation(SWIntegration swintegration)
        {
            InitializeComponent();
            swApp = swintegration.SwApp;
            appset = new appsetting();
            MasterPathTb.Text = appset.PAMasterPath;
            TargetPathTb.Text = appset.PATargetPath;
        }

        private void BrowseBtn_Click(object sender, RoutedEventArgs e)
        {
            string folderPath = "";
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                folderPath = fbd.SelectedPath;
                MasterPathTb.Text = folderPath;
                appset.PAMasterPath = folderPath;
                appset.Save();
            }
        }

        private void tBrowseBtn_Click(object sender, RoutedEventArgs e)
        {
            string folderPath = "";
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                folderPath = fbd.SelectedPath;
                TargetPathTb.Text = folderPath;
                appset.PATargetPath = folderPath;
                appset.Save();
            }
        }

        private void MasterPathTb_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            // clear current selection
            PlatformTypeCb.SelectedIndex = -1;

            // clear items
            PlatformTypeCb.Items.Clear();

            try
            {
                if(Directory.Exists(MasterPathTb.Text))
                {
                    string[] dirs = Directory.GetDirectories(MasterPathTb.Text, "*",SearchOption.TopDirectoryOnly);
                    foreach(string dir in dirs)
                    {
                        string folderName  = Path.GetFileName(dir);
                        PlatformTypeCb.Items.Add(folderName);
                    }
                    PlatformTypeCb.SelectedIndex = 0;
                }
            }
            catch (System.Exception)
            {
            }
        }

        private void PlatformTypeCb_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            switch (PlatformTypeCb.SelectedValue)
            {
                case ("OverhangLeftOpen"):
                    {
                        platformType = PlatformType.OverhangLeftOpen;
                    }
                    break;
            }
        }

        private static readonly Regex _regex = new Regex("[^0-9.]+");
        private static bool IsTextAllowed(string text)
        {
            return !_regex.IsMatch(text);
        }

        private void NumericInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            e.Handled = !IsTextAllowed(e.Text);
        }

        private void GenerateBtn_Click(object sender, RoutedEventArgs e)
        {
            // check if text box are empty
            bool inputEmpty = false;
            if(TargetPathTb.Text == "") inputEmpty = true;
            if(OLOHeight.Text == "") inputEmpty = true;
            if(OLOAngle.Text == "") inputEmpty = true;
            if(OLOOverhangLength.Text == "") inputEmpty = true;
            if(OLOOverhangWidth.Text == "") inputEmpty = true;

            // collect variables
            double oloHeight = double.Parse(OLOHeight.Text)/1000;
            double oloAngle = double.Parse(OLOAngle.Text) * Math.PI / 180;
            double oloOLength = double.Parse(OLOOverhangLength.Text)/1000;
            double oloOWidth = double.Parse(OLOOverhangWidth.Text) / 1000;
            double oloStairWidth = double.Parse(OLOStairWidthTb.Text) / 1000;

            if (!inputEmpty)
            {
                int iRet = -1;
                SelectionMgr swSelMgr;

                // choose configuration
                switch (HandrailCb.Text)
                {
                    case "Bolt-on":
                        switch (ConfigCb.Text)
                        {
                            case "Wheel and Jockey":
                                configurationName = "Bolt-on";
                                break;
                            case "Wheel Only":
                                configurationName = "Bolt-on Wheel Only";
                                break;
                            case "Fixed":
                                configurationName = "Bolt-on Fixed";
                                break;
                        }
                        break;
                    case "Welded":
                        switch (ConfigCb.Text)
                        {
                            case "Wheel and Jockey":
                                configurationName = "Welded";
                                break;
                            case "Wheel Only":
                                configurationName = "Welded Wheel Only";
                                break;
                            case "Fixed":
                                configurationName = "Welded Fixed";
                                break;
                        }
                        break;
                    case "Removable Bolt-on":
                        switch (ConfigCb.Text)
                        {
                            case "Wheel and Jockey":
                                configurationName = "Removable Bolt-on";
                                break;
                            case "Wheel Only":
                                configurationName = "Removable Bolt-on Wheel Only";
                                break;
                            case "Fixed":
                                configurationName = "Removable Bolt-on Fixed";
                                break;
                        }
                        break;
                    case "Removable Welded":
                        switch (ConfigCb.Text)
                        {
                            case "Wheel and Jockey":
                                configurationName = "Removable Welded";
                                break;
                            case "Wheel Only":
                                configurationName = "Removable Welded Wheel Only";
                                break;
                            case "Fixed":
                                configurationName = "Removable Welded Fixed";
                                break;
                        }
                        break;
                }

                // pack and go
                string[] allFiles = Directory.GetFiles(MasterPathTb.Text + "\\" + PlatformTypeCb.Text);
                List<string> asmList = new List<string>();
                List<Tuple<string, string>> nameList = new List<Tuple<string, string>>();
                foreach (string file in allFiles)
                {
                    if (!file.Contains("~"))
                    {
                        // copy files
                        string sourceFile = file;
                        string oriName = Path.GetFileName(sourceFile);
                        string newName = oriName.Replace("XXXXX", DrawingNoTb.Text);
                        string destFile = TargetPathTb.Text + "\\" + newName;
                        iRet = swApp.CopyDocument(sourceFile, destFile, "", "", (int)swMoveCopyOptions_e.swMoveCopyOptionsOverwriteExistingDocs);

                        // set read-only flag to off for newly copied files
                        FileInfo fileInfo = new FileInfo(destFile);
                        fileInfo.IsReadOnly = false;

                        // add to respective list
                        nameList.Add(new Tuple<string,string>(sourceFile, destFile));
                        if(destFile.ToUpper().Contains("SLDASM"))
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
                        bool bRet = swApp.ReplaceReferencedDocument(asm, sourceFile, destFile);
                    }
                }

                // open main assembly
                string mainAsm = TargetPathTb.Text + "\\BT" + DrawingNoTb.Text + "-AS-100.SLDASM";
                int err = -1;
                int warn = -1;
                bool bSelRet;
                ModelDoc2 swModel = swApp.OpenDoc6(mainAsm, (int)swDocumentTypes_e.swDocASSEMBLY, (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel
                    , configurationName, ref err, ref warn);

                #region Calculation
                int ThreadInstance = 2;
                double ThreadHeight = oloHeight / ThreadInstance;
                while(ThreadHeight > 0.225)
                {
                    ThreadInstance++;
                    ThreadHeight = oloHeight / ThreadInstance;
                }
                double ThreadHDim = ThreadHeight / Math.Sin(oloAngle);
                double ThreadGoing = Math.Sqrt(Math.Pow(ThreadHDim, 2)-Math.Pow(ThreadHeight, 2));
                double TwoRPlusG = 2 * ThreadHeight + ThreadGoing;

                double HrOffset = 0.19;
                double L = oloHeight / Math.Tan(oloAngle);
                double Lr = L + HrOffset - ThreadGoing;
                double HrD = Lr / 2;
                double HrP = HrD / Math.Cos(oloAngle);

                double FirstThreadPlate = ((HrD - HrOffset) / Math.Cos(oloAngle)) - (0.105+0.030*Math.Tan(oloAngle));
                #endregion

                #region Modify Frame
                // change height dimension
                string frameName = TargetPathTb.Text + "\\BT" + DrawingNoTb.Text + "-PT-101.SLDPRT";
                ModelDoc2 frameModel = swApp.OpenDoc6(frameName, (int)swDocumentTypes_e.swDocPART,
                    (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, configurationName, ref err, ref warn);
                frameModel.ShowConfiguration2(configurationName);
                SelectionMgr frameSelMgr = frameModel.ISelectionManager;

                bSelRet = frameModel.Extension.SelectByID2("3DSketch1", "SKETCH", 0, 0, 0, false, 0, null, 0);
                Feature skFeat = (Feature)frameSelMgr.GetSelectedObject6(1, -1);
                Dimension heightDim = skFeat.IParameter("D1");
                heightDim.SetSystemValue3(oloHeight, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                // change stair width
                Dimension stairWidthDim = skFeat.IParameter("D2");
                stairWidthDim.SetSystemValue3(oloStairWidth, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                // change angle
                Dimension angleDim = skFeat.IParameter("D6");
                angleDim.SetSystemValue3(oloAngle, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                // cahnge overhang width
                Dimension oWidthDim = skFeat.IParameter("D4");
                oWidthDim.SetSystemValue3(oloOWidth, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                // change overhang length
                Dimension oLengthDim = skFeat.IParameter("D3");
                oLengthDim.SetSystemValue3(oloOLength, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                // conditional dimension
                if(configurationName.Contains("Welded"))
                {
                    if (configurationName.Contains("Fixed"))
                    {
                        bSelRet = frameModel.Extension.SelectByID2("3DSketch5", "SKETCH", 0, 0, 0, false, 0, null, 0);
                        Feature weldedHrSkFeat1 = (Feature)frameSelMgr.GetSelectedObject6(1, -1);
                        Dimension weldedHrSkDim1 = weldedHrSkFeat1.IParameter("D4");
                        weldedHrSkDim1.SetSystemValue3(HrD, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                        Dimension weldedHrSkDim2 = weldedHrSkFeat1.IParameter("D5");
                        weldedHrSkDim2.SetSystemValue3(HrD, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                        if (!configurationName.Contains("Removable"))
                        {
                            bSelRet = frameModel.Extension.SelectByID2("3DSketch4", "SKETCH", 0, 0, 0, false, 0, null, 0);
                            Feature weldedHrSkFeat2 = (Feature)frameSelMgr.GetSelectedObject6(1, -1);
                            Dimension weldedHrSkDim3 = weldedHrSkFeat2.IParameter("D5");
                            weldedHrSkDim3.SetSystemValue3(HrD, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                            Dimension weldedHrSkDim4 = weldedHrSkFeat2.IParameter("D6");
                            weldedHrSkDim4.SetSystemValue3(HrD, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                        }
                    }
                    else
                    {
                        bSelRet = frameModel.Extension.SelectByID2("3DSketch3", "SKETCH", 0, 0, 0, false, 0, null, 0);
                        Feature weldedHrSkFeat1 = (Feature)frameSelMgr.GetSelectedObject6(1, -1);
                        Dimension weldedHrSkDim1 = weldedHrSkFeat1.IParameter("D4");
                        weldedHrSkDim1.SetSystemValue3(HrD, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                        Dimension weldedHrSkDim2 = weldedHrSkFeat1.IParameter("D5");
                        weldedHrSkDim2.SetSystemValue3(HrD, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                        if (!configurationName.Contains("Removable"))
                        {
                            bSelRet = frameModel.Extension.SelectByID2("3DSketch2", "SKETCH", 0, 0, 0, false, 0, null, 0);
                            Feature weldedHrSkFeat2 = (Feature)frameSelMgr.GetSelectedObject6(1, -1);
                            Dimension weldedHrSkDim3 = weldedHrSkFeat2.IParameter("D5");
                            weldedHrSkDim3.SetSystemValue3(HrD, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                            Dimension weldedHrSkDim4 = weldedHrSkFeat2.IParameter("D6");
                            weldedHrSkDim4.SetSystemValue3(HrD, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                        }
                    }

                }
                else
                {
                    if (configurationName.Contains("Fixed"))
                    {
                        // change plate dimension
                        bSelRet = frameModel.Extension.SelectByID2("Distance11", "MATE", 0, 0, 0, false, 0, null, 0);
                        Feature mateFeat = (Feature)frameSelMgr.GetSelectedObject6(1, -1);
                        Mate2 mateMate = (Mate2)mateFeat.GetSpecificFeature2();
                        DisplayDimension mateDispDim = mateMate.DisplayDimension2[0];
                        Dimension mateDim = mateDispDim.IGetDimension();
                        int setDimRet = mateDim.SetSystemValue3(FirstThreadPlate, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                        // change pattern dimension 
                        bSelRet = frameModel.Extension.SelectByID2("LPattern4", "BODYFEATURE", 0, 0, 0, false, 0, null, 0);
                        Feature frPatternFeat = (Feature)frameSelMgr.GetSelectedObject6(1, -1);
                        Dimension frPatternDim = frPatternFeat.IParameter("D3");
                        frPatternDim.SetSystemValue3(HrP, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                    }
                    else
                    {
                        // change plate dimension
                        bSelRet = frameModel.Extension.SelectByID2("Distance7", "MATE", 0, 0, 0, false, 0, null, 0);
                        Feature mateFeat = (Feature)frameSelMgr.GetSelectedObject6(1, -1);
                        Mate2 mateMate = (Mate2)mateFeat.GetSpecificFeature2();
                        DisplayDimension mateDispDim = mateMate.DisplayDimension2[0];
                        Dimension mateDim = mateDispDim.IGetDimension();
                        int setDimRet = mateDim.SetSystemValue3(FirstThreadPlate, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                        // change pattern dimension 
                        bSelRet = frameModel.Extension.SelectByID2("LPattern1", "BODYFEATURE", 0, 0, 0, false, 0, null, 0);
                        Feature frPatternFeat = (Feature)frameSelMgr.GetSelectedObject6(1, -1);
                        Dimension frPatternDim = frPatternFeat.IParameter("D3");
                        frPatternDim.SetSystemValue3(HrP, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                    }
                }

                frameModel.ClearSelection2(true);

                DeleteConfiguration(frameModel);
                DeleteSuppressedFeature(frameModel);
                DeleteSuppressedFeature(frameModel);

                // method to close
                CloseDoc(frameModel);
                #endregion

                #region Modify Handrail #1
                if (configurationName.Contains("Bolt-on") && !configurationName.Contains("Removable"))
                {
                    string firstHrName = TargetPathTb.Text + "\\BT" + DrawingNoTb.Text + "-PT-104.SLDPRT";
                    ModelDoc2 firstHrModel = swApp.OpenDoc6(firstHrName, (int)swDocumentTypes_e.swDocPART,
                        (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", ref err, ref warn);
                    SelectionMgr firstHrSelMgr = firstHrModel.ISelectionManager;
                    bSelRet = firstHrModel.Extension.SelectByID2("Sketch1", "SKETCH", 0, 0, 0, false, 0, null, 0);
                    Feature hrSkFeat = (Feature)firstHrSelMgr.GetSelectedObject6(1, -1);
                    Dimension hrHeightDim = hrSkFeat.IParameter("D1");
                    hrHeightDim.SetSystemValue3(oloHeight, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                    Dimension hrGapDim = hrSkFeat.IParameter("D8");
                    hrGapDim.SetSystemValue3(oloOLength + 0.05, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                    Dimension hrDisDim = hrSkFeat.IParameter("D12");
                    hrDisDim.SetSystemValue3(HrD, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                    Dimension hrAngleDim = hrSkFeat.IParameter("D11");
                    hrAngleDim.SetSystemValue3(oloAngle, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                    // change stair width dim
                    bSelRet = firstHrModel.Extension.SelectByID2("Sketch2", "SKETCH", 0, 0, 0, false, 0, null, 0);
                    Feature firstHrStairWidthFeat = (Feature)firstHrSelMgr.GetSelectedObject6(1, -1);
                    Dimension firstHrStairWidthDim = firstHrStairWidthFeat.IParameter("D1");
                    firstHrStairWidthDim.SetSystemValue3(oloStairWidth - 0.05, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                    bSelRet = firstHrModel.Extension.SelectByID2("LPattern1", "BODYFEATURE", 0, 0, 0, false, 0, null, 0);
                    Feature hrPatternFeat = (Feature)firstHrSelMgr.GetSelectedObject6(1, -1);
                    Dimension hrPatternDim = hrPatternFeat.IParameter("D3");
                    hrPatternDim.SetSystemValue3(HrP, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                    CloseDoc(firstHrModel);
                }
                #endregion

                #region Modify Handrail #2
                if (configurationName.Contains("Bolt-on"))
                {
                    string secondHrName = TargetPathTb.Text + "\\BT" + DrawingNoTb.Text + "-PT-105.SLDPRT";
                    ModelDoc2 secondHrModel = swApp.OpenDoc6(secondHrName, (int)swDocumentTypes_e.swDocPART,
                        (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", ref err, ref warn);
                    SelectionMgr secondHrSelMgr = secondHrModel.ISelectionManager;
                    bSelRet = secondHrModel.Extension.SelectByID2("Sketch1", "SKETCH", 0, 0, 0, false, 0, null, 0);
                    Feature hrSkFeat2 = (Feature)secondHrSelMgr.GetSelectedObject6(1, -1);
                    Dimension hrHeightDim2 = hrSkFeat2.IParameter("D1");
                    hrHeightDim2.SetSystemValue3(oloHeight, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                    Dimension hrGapDim2 = hrSkFeat2.IParameter("D5");
                    hrGapDim2.SetSystemValue3(oloOLength + 0.05, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                    Dimension hrDisDim2 = hrSkFeat2.IParameter("D11");
                    hrDisDim2.SetSystemValue3(HrD, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                    Dimension hrAngleDim2 = hrSkFeat2.IParameter("D10");
                    hrAngleDim2.SetSystemValue3(oloAngle, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                    bSelRet = secondHrModel.Extension.SelectByID2("LPattern1", "BODYFEATURE", 0, 0, 0, false, 0, null, 0);
                    Feature hrPatternFeat2 = (Feature)secondHrSelMgr.GetSelectedObject6(1, -1);
                    Dimension hrPatternDim2 = hrPatternFeat2.IParameter("D3");
                    hrPatternDim2.SetSystemValue3(HrP, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                    CloseDoc(secondHrModel);
                }
                #endregion

                #region Modify landing thread
                string landingGrateName= TargetPathTb.Text + "\\BT" + DrawingNoTb.Text + "-PT-103.SLDPRT";
                ModelDoc2 landingGrateModel = swApp.OpenDoc6(landingGrateName, (int)swDocumentTypes_e.swDocPART,
                    (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", ref err, ref warn);
                SelectionMgr landingGrateSelMgr = landingGrateModel.ISelectionManager;
                EquationMgr landingGrateEqMgr = landingGrateModel.GetEquationMgr();
                
                // change equation
                landingGrateEqMgr.SetEquationAndConfigurationOption(2, "\"\"Span\"\" = "+((oloStairWidth - 0.1 + oloOWidth)*1000).ToString(), (int)swInConfigurationOpts_e.swThisConfiguration, null);
                landingGrateEqMgr.SetEquationAndConfigurationOption(3, "\"\"Width\"\" = "+((oloOLength-0.05)*1000).ToString(), (int)swInConfigurationOpts_e.swThisConfiguration, null);

                // get feature
                bSelRet = landingGrateModel.Extension.SelectByID2("Sketch16", "SKETCH", 0, 0, 0, false, 0, null, 0);
                Feature landingGrateCutFeat = (Feature)landingGrateSelMgr.GetSelectedObject6(1, -1);

                // change parameter
                Dimension landingGrateCutLength= landingGrateCutFeat.IParameter("D1");
                landingGrateCutLength.SetSystemValue3(oloOWidth+0.005, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                CloseDoc(landingGrateModel);
                #endregion

                #region Modify Stair Thread
                string stairThreadName = TargetPathTb.Text + "\\BT" + DrawingNoTb.Text + "-PT-102.SLDPRT";
                ModelDoc2 stairThreadModel = swApp.OpenDoc6(stairThreadName, (int)swDocumentTypes_e.swDocPART,
                    (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", ref err, ref warn);
                EquationMgr stairThreadEqMgr = stairThreadModel.GetEquationMgr();

                // change equation
                stairThreadEqMgr.SetEquationAndConfigurationOption(2, "\"\"Span\"\" = " + ((oloStairWidth - 0.1) * 1000).ToString(), (int)swInConfigurationOpts_e.swThisConfiguration, null);

                // close doc
                CloseDoc(stairThreadModel);
                #endregion

                #region Modify Self Closing Gate
                string selfClosingGateName = TargetPathTb.Text + "\\BT" + DrawingNoTb.Text + "-AS-150.SLDASM";
                ModelDoc2 selfClosingGateModel = swApp.OpenDoc6(selfClosingGateName, (int)swDocumentTypes_e.swDocASSEMBLY,
                    (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", ref err, ref warn);
                EquationMgr selfClosingGateEqMgr = selfClosingGateModel.GetEquationMgr();

                // change equation
                selfClosingGateEqMgr.SetEquationAndConfigurationOption(0, "\"\"ENTRANCE WIDTH\"\" = " + ((oloStairWidth - 0.05) * 1000).ToString(), (int)swInConfigurationOpts_e.swThisConfiguration, null);

                // close doc
                CloseDoc(selfClosingGateModel);
                #endregion

                #region Modify Removable Handrail
                if (configurationName.Contains("Removable"))
                {
                    string removableHrName = TargetPathTb.Text + "\\BT" + DrawingNoTb.Text + "-PT-201.SLDPRT";
                    ModelDoc2 removableHrModel = swApp.OpenDoc6(removableHrName, (int)swDocumentTypes_e.swDocPART,
                        (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", ref err, ref warn);
                    SelectionMgr removableHrSelMgr= removableHrModel.ISelectionManager;

                    // get feature
                    bSelRet = removableHrModel.Extension.SelectByID2("Sketch1", "SKETCH", 0, 0, 0, false, 0, null, 0);
                    Feature removeableHrSkFeat = (Feature)removableHrSelMgr.GetSelectedObject6(1, -1);

                    // change parameter
                    Dimension removeableHrWidthDim = removeableHrSkFeat.IParameter("D2");
                    removeableHrWidthDim.SetSystemValue3(oloStairWidth - 0.07, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                    // close doc
                    CloseDoc(removableHrModel);

                    // different set of removable handrail
                    if (Math.Round(oloOLength,4) != Math.Round((oloStairWidth - 0.05),4)) // to compare double need to round up first
                    {
                        // suppress handrail
                        swModel = (ModelDoc2)swApp.ActivateDoc3(Path.GetFileName(mainAsm), false, (int)swRebuildOptions_e.swRebuildAll, ref err);
                        swSelMgr = swModel.ISelectionManager;
                        string removableHrNameInst2 = "BT" + DrawingNoTb.Text + "-AS-200-1@BT" + DrawingNoTb.Text + "-AS-100";
                        swModel.Extension.SelectByID2(removableHrNameInst2, "COMPONENT", 0, 0, 0, false, 0, null, 0);
                        Component2 removableHrInst2Comp = (Component2)swSelMgr.GetSelectedObject6(1, -1);
                        removableHrInst2Comp.SetSuppression2((int)swComponentSuppressionState_e.swComponentSuppressed);
                        
                        // unsuppress handrail
                        string removableHr2Name = "BT" + DrawingNoTb.Text + "-AS-250-1@BT" + DrawingNoTb.Text + "-AS-100";
                        swModel.Extension.SelectByID2(removableHr2Name, "COMPONENT", 0, 0, 0, false, 0, null, 0);
                        Component2 removableHr2Comp = (Component2)swSelMgr.GetSelectedObject6(1, -1);
                        removableHr2Comp.SetSuppression2((int)swComponentSuppressionState_e.swComponentFullyResolved);

                        // select the part of the unsuppressed handrail
                        removableHr2Name = TargetPathTb.Text + "\\BT" + DrawingNoTb.Text + "-PT-251.SLDPRT";
                        ModelDoc2 removableHr2Model = swApp.OpenDoc6(removableHr2Name, (int)swDocumentTypes_e.swDocPART,
                            (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", ref err, ref warn);
                        SelectionMgr removableHr2SelMgr = removableHr2Model.ISelectionManager;

                        // get feature
                        bSelRet = removableHr2Model.Extension.SelectByID2("Sketch1", "SKETCH", 0, 0, 0, false, 0, null, 0);
                        Feature removeableHr2SkFeat = (Feature)removableHr2SelMgr.GetSelectedObject6(1, -1);

                        // change parameter
                        Dimension removeableHr2WidthDim = removeableHr2SkFeat.IParameter("D2");
                        removeableHr2WidthDim.SetSystemValue3(oloOLength - 0.02, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                        // close doc
                        CloseDoc(removableHrModel);

                    }

                }
                #endregion

                #region Modify Thread Pattern
                swModel = (ModelDoc2)swApp.ActivateDoc3(Path.GetFileName(mainAsm), false, (int)swRebuildOptions_e.swRebuildAll, ref err);
                swSelMgr = swModel.ISelectionManager;

                string patternFeatureName = "";

                if (configurationName.Contains("Fixed"))
                {
                    patternFeatureName = "LocalLPattern2";
                }
                else
                {
                    patternFeatureName = "LocalLPattern1";
                }

                bSelRet = swModel.Extension.SelectByID2(patternFeatureName, "COMPPATTERN", 0, 0, 0, false, 0, null, 0);
                Feature patternFeature = (Feature)swSelMgr.GetSelectedObject6(1, -1);
                Dimension threadInsDim = patternFeature.IParameter("D1");
                threadInsDim.SetSystemValue3(Convert.ToDouble(ThreadInstance), (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);

                Dimension threadDisDim = patternFeature.IParameter("D3");
                threadDisDim.SetSystemValue3(ThreadHDim, (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration, null);
                #endregion

                #region Assembly Mate
                if (configurationName.Contains("Removable") && configurationName.Contains("Fixed"))
                {
                    Feature removableHrMate;
                    bSelRet = swModel.Extension.SelectByID2("Concentric81", "MATE", 0, 0, 0, false, 0, null, 0);
                    removableHrMate = (Feature)swSelMgr.GetSelectedObject6(1, -1);
                    removableHrMate.SetSuppression2((int)swFeatureSuppressionAction_e.swSuppressFeature, (int)swInConfigurationOpts_e.swThisConfiguration, null);
                    bSelRet = swModel.Extension.SelectByID2("Concentric82", "MATE", 0, 0, 0, false, 0, null, 0);
                    removableHrMate = (Feature)swSelMgr.GetSelectedObject6(1, -1);
                    removableHrMate.SetSuppression2((int)swFeatureSuppressionAction_e.swSuppressFeature, (int)swInConfigurationOpts_e.swThisConfiguration, null);
                    bSelRet = swModel.Extension.SelectByID2("Coincident376", "MATE", 0, 0, 0, false, 0, null, 0);
                    removableHrMate = (Feature)swSelMgr.GetSelectedObject6(1, -1);
                    removableHrMate.SetSuppression2((int)swFeatureSuppressionAction_e.swSuppressFeature, (int)swInConfigurationOpts_e.swThisConfiguration, null);
                    bSelRet = swModel.Extension.SelectByID2("Concentric83", "MATE", 0, 0, 0, false, 0, null, 0);
                    removableHrMate = (Feature)swSelMgr.GetSelectedObject6(1, -1);
                    removableHrMate.SetSuppression2((int)swFeatureSuppressionAction_e.swUnSuppressFeature, (int)swInConfigurationOpts_e.swThisConfiguration, null);
                    bSelRet = swModel.Extension.SelectByID2("Concentric84", "MATE", 0, 0, 0, false, 0, null, 0);
                    removableHrMate = (Feature)swSelMgr.GetSelectedObject6(1, -1);
                    removableHrMate.SetSuppression2((int)swFeatureSuppressionAction_e.swUnSuppressFeature, (int)swInConfigurationOpts_e.swThisConfiguration, null);
                    bSelRet = swModel.Extension.SelectByID2("Coincident377", "MATE", 0, 0, 0, false, 0, null, 0);
                    removableHrMate = (Feature)swSelMgr.GetSelectedObject6(1, -1);
                    removableHrMate.SetSuppression2((int)swFeatureSuppressionAction_e.swUnSuppressFeature, (int)swInConfigurationOpts_e.swThisConfiguration, null);
                }
                #endregion

                swModel.EditRebuild3();

                DeleteConfiguration(swModel);

                // delete all supressed component
                Component2 swRootComp = swModel.ConfigurationManager.ActiveConfiguration.IGetRootComponent2();
                object[] vChildComps = (object[])swRootComp.GetChildren();
                foreach (object childComp in vChildComps)
                {
                    Component2 swChildComp = (Component2)childComp;
                    if (swChildComp.IsSuppressed())
                    {
                        swChildComp.Select4(true, null, false);
                    }
                }
                AssemblyDoc swAsm = (AssemblyDoc)swModel;
                swAsm.DeleteSelections((int)swAssemblyDeleteOptions_e.swDelete_SelectedComponents);
                swModel.ClearSelection2(true);

                // delete all suppressed feature
                DeleteSuppressedFeature(swModel);
                DeleteSuppressedFeature(swModel);

                swModel.EditRebuild3();
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Some input are empty.");
                return;
            }
        }

        private void CloseDoc(ModelDoc2 swmodel)
        {
            int err = -1;
            int warn = -1;
            swmodel.EditRebuild3();
            swmodel.ClearSelection2(true);
            swmodel.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref err, ref warn);
        }

        private void DeleteSuppressedFeature(ModelDoc2 swmodel)
        {
            Feature deleteFeat = swmodel.IFirstFeature();
            while (deleteFeat != null)
            {
                if (deleteFeat.IsSuppressed())
                {
                    deleteFeat.Select2(true, 0);
                }
                deleteFeat = deleteFeat.IGetNextFeature();
            }
            swmodel.Extension.DeleteSelection2((int)swDeleteSelectionOptions_e.swDelete_Absorbed + (int)swDeleteSelectionOptions_e.swDelete_Children);
        }

        private void DeleteConfiguration(ModelDoc2 swmodel)
        {
            // get all configuration
            string[] configNames = (string[])swmodel.GetConfigurationNames();
            foreach (string configName in configNames)
            {
                swmodel.DeleteConfiguration2(configName);
            }
        }

        private void oloConfigChanged(object sender, SelectionChangedEventArgs e)
        {
            System.Windows.Controls.ComboBox modifiedCb = (System.Windows.Controls.ComboBox)sender;
            ComboBoxItem selectedItem = (ComboBoxItem)modifiedCb.SelectedItem;
            configName = selectedItem.Content.ToString();
            if(HandrailCb != null)
            {
                handrailName = HandrailCb.Text;
            }
            else
            {
                handrailName = "Wheel and Jockey";
            }
            UpdateImage(configName, handrailName);
        }

        private void oloHandrailChanged(object sender, SelectionChangedEventArgs e)
        {
            System.Windows.Controls.ComboBox modifiedCb = (System.Windows.Controls.ComboBox)sender;
            ComboBoxItem selectedItem = (ComboBoxItem)modifiedCb.SelectedItem;
            handrailName = selectedItem.Content.ToString();
            if (ConfigCb != null)
            {
                configName = ConfigCb.Text;
            }
            else
            {
                configName = "Bolt-on";
            }
            UpdateImage(configName, handrailName);
        }

        private void UpdateImage(string configname, string handrailname)
        {
            string imageSourceString = "";

            // choose configuration
            switch (handrailname)
            {
                case "Bolt-on":
                    switch (configname)
                    {
                        case "Wheel and Jockey":
                            imageSourceString = "OverhangLeftOpenBolton";
                            break;
                        case "Wheel Only":
                            imageSourceString = "OverhangLeftOpenBoltonWheelOnly";
                            break;
                        case "Fixed":
                            imageSourceString = "OverhangLeftOpenBoltonFixed";
                            break;
                    }
                    break;
                case "Welded":
                    switch (configname)
                    {
                        case "Wheel and Jockey":
                            imageSourceString = "OverhangLeftOpenWelded";
                            break;
                        case "Wheel Only":
                            imageSourceString = "OverhangLeftOpenWeldedWheelOnly";
                            break;
                        case "Fixed":
                            imageSourceString = "OverhangLeftOpenWeldedFixed";
                            break;
                    }
                    break;
                case "Removable Bolt-on":
                    switch (configname)
                    {
                        case "Wheel and Jockey":
                            imageSourceString = "OverhangLeftOpenRemovableBolton";
                            break;
                        case "Wheel Only":
                            imageSourceString = "OverhangLeftOpenRemovableBoltonWheelOnly";
                            break;
                        case "Fixed":
                            imageSourceString = "OverhangLeftOpenRemovableBoltonFixed";
                            break;
                    }
                    break;
                case "Removable Welded":
                    switch (configname)
                    {
                        case "Wheel and Jockey":
                            imageSourceString = "OverhangLeftOpenRemovableWelded";
                            break;
                        case "Wheel Only":
                            imageSourceString = "OverhangLeftOpenRemovableWeldedWheelOnly";
                            break;
                        case "Fixed":
                            imageSourceString = "OverhangLeftOpenRemovableWeldedFixed";
                            break;
                    }
                    break;
            }

            if (oloTypeImage != null)
            {
                oloTypeImage.Source = new BitmapImage(new Uri("/mStructural;component/Resources/Platform Automation/OverhangLeftOpen/" + imageSourceString +".png",UriKind.Relative));
            }
        }
    }
}
