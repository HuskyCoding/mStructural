using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System;
using System.Runtime.InteropServices;
using System.Collections;

namespace mStructural.Function
{
    public class ExportBom
    {
        #region Private Variables
        private SldWorks swApp;
        private Message msg;
        private Macros macros;
        #endregion

        // constructor
        public ExportBom(SWIntegration swIntegration)
        {
            swApp = swIntegration.SwApp;
            msg = new Message(swApp);
            macros = new Macros(swApp);
        }

        // Method to run the function
        public void Run()
        {
            // get active doc
            ModelDoc2 swModel = swApp.IActiveDoc2;

            // check
            bool bRet = macros.checkDocNullandType(swModel, swDocumentTypes_e.swDocDRAWING); if (!bRet) return;

            // show confirm export
            bool? toProduction;
            WPF.ProductionExportForm pef = new WPF.ProductionExportForm();
            bool? dialogResult = pef.ShowDialog();
            if (dialogResult == true)
            {
                toProduction = pef.ToProductionCh.IsChecked;
            }
            else
            {
                return;
            }

            // cast to drawing document
            DrawingDoc swDraw = (DrawingDoc)swModel;

            // get current model name
            string sFileName = Path.GetFileNameWithoutExtension(swModel.GetPathName());

            // get current path name
            string sPathName = Path.GetDirectoryName(swModel.GetPathName());

            // find the \
            int index = sPathName.LastIndexOf("\\");

            // add production in path
            if(toProduction == true)
            {
                sPathName = sPathName.Substring(0, index) + "\\Production";
            }

            // create directory
            Directory.CreateDirectory(sPathName);

            // initialize custom prop manager
            CustomPropertyManager swCusPropMgr = swModel.Extension.CustomPropertyManager[""];

            // get revision custom prop
            string val = "";
            string valout = "";
            bRet = swCusPropMgr.Get4("Revision", false, out val, out valout);

            // Create excel
            Application xlApp = new Application();

            // check if excel is installed
            if(xlApp == null)
            {
                msg.ErrorMsg("Error in starting excel");
                return;
            }

            // make excel visible
            xlApp.Visible = true;

            // add a workbook
            Workbook xlWb = xlApp.Workbooks.Add();

            // get all sheet names
            string[] sheetNames = (string[])swDraw.GetSheetNames();

            // loop for each sheet
            foreach(string sheetName in sheetNames)
            {
                // activate the specific sheet
                swDraw.ActivateSheet(sheetName);

                Hashtable excelSheetNames = new Hashtable();
                int hashIndex = 0;

                // get the first view
                View swView = swDraw.IGetFirstView();

                // loop all view
                if(swView != null)
                {
                    // get first table annotation
                    TableAnnotation swTable = swView.GetFirstTableAnnotation();
                    
                    if(swTable != null)
                    {
                        // set table index = 1
                        int tableIndex = 1;

                        // loop table annotation
                        while (swTable != null)
                        {
                            string firstBalloon = swTable.DisplayedText2[1, 0, false];
                            switch (swTable.Type)
                            {
                                case (int)swTableAnnotationType_e.swTableAnnotation_WeldmentCutList:
                                    {
                                        // cast to wcl annotation
                                        WeldmentCutListAnnotation clTable = (WeldmentCutListAnnotation)swTable;
                                        WeldmentCutListFeature clFeat = clTable.WeldmentCutListFeature;

                                        string modelName = "";

                                        // name of the configuration
                                        string configName = clFeat.Configuration;

                                        while(swView != null)
                                        {
                                            Note swNote = swView.IGetFirstNote();
                                            while(swNote != null)
                                            {
                                                if(firstBalloon == swNote.GetBomBalloonText(true))
                                                {
                                                    modelName = Path.GetFileNameWithoutExtension(swView.GetReferencedModelName());
                                                }
                                                swNote = swNote.IGetNext();
                                            }
                                            swView = swView.IGetNextView();
                                        }

                                        // get active sheet
                                        Worksheet xlWs = (Worksheet)xlWb.ActiveSheet;

                                        // check if the modelname longer than 31 char
                                        if (modelName.Length <= 31)
                                        {
                                            // directly rename if it is sheet1, else create new sheet first
                                            if (xlWs.Name == "Sheet1")
                                            {
                                                int j = 1;
                                                while (excelSheetNames.ContainsValue(modelName))
                                                {
                                                    modelName = modelName + "(" + j.ToString() + ")";
                                                    j++;
                                                    if (j > 10) break;
                                                }
                                                xlWs.Name = modelName;
                                                excelSheetNames.Add(hashIndex, modelName);
                                                hashIndex++;
                                            }
                                            else
                                            {
                                                int j = 1;
                                                while (excelSheetNames.ContainsValue(modelName))
                                                {
                                                    modelName = modelName + "(" + j.ToString() + ")";
                                                    j++;
                                                    if (j > 10) break;
                                                }
                                                xlWs = (Worksheet)xlWb.Sheets.Add(After: xlWb.Sheets[xlWb.Worksheets.Count]);
                                                xlWs.Name = modelName;
                                                excelSheetNames.Add(hashIndex, modelName);
                                                hashIndex++;
                                            }
                                        }

                                        /*// directly rename if it is sheet1, else create new sheet first
                                        if (xlWs.Name == "Sheet1")
                                        {
                                            xlWs.Name = modelName + "_" + configName;
                                        }
                                        else
                                        {
                                            xlWs = (Worksheet)xlWb.Sheets.Add(After: xlWb.Sheets[xlWb.Worksheets.Count]);
                                            xlWs.Name = modelName + "_" + configName;
                                        }*/

                                        // write to excel
                                        for (int i = 0; i < swTable.RowCount; i++)
                                        {
                                            for (int j = 0; j < swTable.ColumnCount; j++)
                                            {
                                                xlWs.Cells[i + 1, j + 1] = swTable.Text2[i, j, false];
                                            }
                                        }

                                        AutofitWs(xlWs);

                                        releaseObject(xlWs); // excel object need to release after use
                                        tableIndex++; // increment table index
                                    }
                                    break;
                                case (int)swTableAnnotationType_e.swTableAnnotation_BillOfMaterials:
                                    {
                                        // cast to wcl annotation
                                        BomTableAnnotation bomTable = (BomTableAnnotation)swTable;
                                        BomFeature bomFeat = bomTable.BomFeature;

                                        // name of the configuration
                                        string configName = "";
                                        switch (bomFeat.TableType)
                                        {
                                            case (int)swBomType_e.swBomType_Indented:
                                            case (int)swBomType_e.swBomType_PartsOnly:
                                                {
                                                    configName = bomFeat.Configuration;
                                                    break;
                                                }
                                            case (int)swBomType_e.swBomType_TopLevelOnly:
                                                {
                                                    object bomVisibleRef = null;
                                                    configName = bomFeat.GetConfigurations(true, bomVisibleRef)[0];
                                                    break;
                                                }
                                            }

                                        // name of reference model
                                        string modelName = bomFeat.GetReferencedModelName();
                                        modelName = Path.GetFileNameWithoutExtension(modelName);

                                        // get active sheet
                                        Worksheet xlWs = (Worksheet)xlWb.ActiveSheet;

                                        // check if the modelname longer than 31 char
                                        if (modelName.Length <= 31)
                                        {
                                            // directly rename if it is sheet1, else create new sheet first
                                            if (xlWs.Name == "Sheet1")
                                            {
                                                int j = 1;
                                                while (excelSheetNames.ContainsValue(modelName))
                                                {
                                                    modelName = modelName + "(" + j.ToString() + ")";
                                                    j++;
                                                    if (j > 10) break;
                                                }
                                                xlWs.Name = modelName;
                                                excelSheetNames.Add(hashIndex, modelName);
                                                hashIndex++;
                                            }
                                            else
                                            {
                                                int j = 1;
                                                while (excelSheetNames.ContainsValue(modelName))
                                                {
                                                    modelName = modelName + "(" + j.ToString() + ")";
                                                    j++;
                                                    if (j > 10) break;
                                                }
                                                xlWs = (Worksheet)xlWb.Sheets.Add(After: xlWb.Sheets[xlWb.Worksheets.Count]);
                                                xlWs.Name = modelName;
                                                excelSheetNames.Add(hashIndex, modelName);
                                                hashIndex++;
                                            }
                                        }

                                        /*// directly rename if it is sheet1, else create new sheet first
                                        if (xlWs.Name == "Sheet1")
                                        {
                                            xlWs.Name = modelName + "_" + configName;
                                        }
                                        else
                                        {
                                            xlWs = (Worksheet)xlWb.Sheets.Add(After: xlWb.Sheets[xlWb.Worksheets.Count]);
                                            xlWs.Name = modelName + "_" + configName;
                                        }*/

                                        // write to excel
                                        for (int i = 0; i < swTable.RowCount; i++)
                                        {
                                            for (int j = 0; j < swTable.ColumnCount; j++)
                                            {
                                                xlWs.Cells[i + 1, j + 1] = swTable.Text2[i, j, false];
                                            }
                                        }

                                        AutofitWs(xlWs);

                                        releaseObject(xlWs); // excel object need to release after use
                                        tableIndex++; // increment table index
                                    }
                                    break;
                            }

                            // get next table annotation
                            swTable = swTable.GetNext();
                        }

                    }
                }
            }

            // get file save name
            string saveName = sPathName + "\\" + sFileName +"_Rev" + valout +".xlsx";

            // save excel file
            xlApp.ActiveWorkbook.SaveAs(saveName);

            releaseObject(xlWb); // excel object need to release after use
            releaseObject(xlApp); // excel object need to release after use

            msg.InfoMsg("Export BOM completed!");
        }

        // Method to release excel object
        private void releaseObject(object obj)
        {
            try
            {
                Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                msg.ErrorMsg(ex.Message);
            }
            finally
            {
                GC.Collect();
            }
        }

        // method to fit worksheet
        private void AutofitWs(Worksheet ws)
        {
            Range myRange = ws.Range["A1"].CurrentRegion;
            myRange.EntireColumn.AutoFit();
            myRange.Borders.LineStyle = XlLineStyle.xlContinuous;
            myRange.Borders.Weight = XlBorderWeight.xlThin;
        }
    }
}
