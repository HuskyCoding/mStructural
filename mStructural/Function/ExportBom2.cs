using Microsoft.Office.Interop.Excel;
using mStructural.Classes;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace mStructural.Function
{
    public class ExportBom2
    {
        // Main Method
        public StringBuilder Run(SldWorks swapp, string drawingpath, bool isproduction, ObservableCollection<ExportBomClass> bomcol, string rev)
        {
            StringBuilder sb = new StringBuilder();
            Log(sb, "Initiate BOM Export.");

            // open drawing
            int err = -1;
            int warn = -1;
            ModelDoc2 swModel = swapp.OpenDoc6(drawingpath, (int)swDocumentTypes_e.swDocDRAWING, (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel, "", ref err, ref warn);

            // cast as drawing doc
            DrawingDoc swDraw = swModel as DrawingDoc;

            // get current model name
            string sFileName = Path.GetFileNameWithoutExtension(swModel.GetPathName());

            // get current path name
            string sPathName = Path.GetDirectoryName(swModel.GetPathName());

            // find the \
            int index = sPathName.LastIndexOf("\\");

            // add production in path
            if (isproduction == true)
            {
                sPathName = sPathName.Substring(0, index) + "\\Production";
            }

            // create directory
            Directory.CreateDirectory(sPathName);

            // Create excel
            Application xlApp = new Application();
            xlApp.DisplayAlerts = false;

            // check if excel is installed
            if (xlApp == null)
            {
                Log(sb, "Failed to start excel");
                return sb;
            }

            // make excel visible
            xlApp.Visible = true;

            // add a workbook
            Workbook xlWb = xlApp.Workbooks.Add();
            
            Hashtable sheetNames = new Hashtable();
            int hashIndex = 0;
            foreach(ExportBomClass ebc in bomcol)
            {
                string modelName = Path.GetFileNameWithoutExtension(ebc.SwModel.GetPathName());
                TableAnnotation swTable = ebc.SwTableAnn;

                // get active sheet
                Worksheet xlWs = (Worksheet)xlWb.ActiveSheet;

                // check if the modelname longer than 31 char
                if (modelName.Length <= 31)
                {
                    // directly rename if it is sheet1, else create new sheet first
                    if (xlWs.Name == "Sheet1")
                    {
                        int j = 1;
                        while(sheetNames.ContainsValue(modelName))
                        {
                            modelName = modelName +"("+j.ToString()+")";
                            j++;
                            if (j > 10) break;
                        }
                        xlWs.Name = modelName;
                        sheetNames.Add(hashIndex, modelName);
                        hashIndex++;
                    }
                    else
                    {
                        int j = 1;
                        while (sheetNames.ContainsValue(modelName))
                        {
                            modelName = modelName + "(" + j.ToString() + ")";
                            j++;
                            if (j > 10) break;
                        }
                        xlWs = (Worksheet)xlWb.Sheets.Add(After: xlWb.Sheets[xlWb.Worksheets.Count]);
                        xlWs.Name = modelName;
                        sheetNames.Add(hashIndex, modelName);
                        hashIndex++;
                    }
                }
                else
                {
                    Log(sb, "Model name: " + modelName + " is longer than 31 characters, default sheet name: " + xlWs.Name + " will be use instead.");
                }

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

                // rename body
                if (ebc.ExportTubeLaser)
                {
                    if(swTable.Type == (int)swTableAnnotationType_e.swTableAnnotation_WeldmentCutList)
                    {
                        // lock row height
                        for(int i = 0; i < swTable.RowCount; i++)
                        {
                            swTable.SetLockRowHeight(i, true);
                        }

                        swTable.InsertColumn2((int)swTableItemInsertPosition_e.swTableItemInsertPosition_Last, 0, "", (int)swInsertTableColumnWidthStyle_e.swInsertColumn_DefaultWidth);
                        swTable.SetColumnType2(swTable.ColumnCount - 1, (int)swTableColumnTypes_e.swWeldTableColumnType_CutListName, false);

                        ModelDoc2 tableModel = swapp.ActivateDoc3(modelName, false, (int)swRebuildOnActivation_e.swRebuildActiveDoc, ref err);
                        if(tableModel != null)
                        {
                            tableModel.ShowConfiguration2(ebc.TubeLaserConfiguration);

                            SelectionMgr tableModelSelMgr = tableModel.SelectionManager;

                            for(int i = 1; i < swTable.RowCount; i++)
                            {
                                tableModel.Extension.SelectByID2(swTable.DisplayedText2[i, swTable.ColumnCount - 1, false], "BDYFOLDER", 0, 0, 0, false, 0, null, 0);
                                Feature swFeat = tableModelSelMgr.GetSelectedObject6(1, -1);
                                BodyFolder swBodyFolder = swFeat.GetSpecificFeature2();
                                object[] bodies = swBodyFolder.GetBodies();
                                int j = 0;
                                if (bodies!=null)
                                {
                                    foreach(object body in bodies)
                                    {
                                        Body2 swBody = (Body2)body;
                                        swBody.Name = modelName + "_" + swTable.DisplayedText2[i, 0, false] + "_" + j;
                                        j++;
                                    }
                                }
                            }

                            tableModel.ForceRebuild3(false);
                            tableModel.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref err, ref warn);
                            swapp.CloseDoc(modelName);
                        }
                        
                        swTable.DeleteColumn2(swTable.ColumnCount - 1, false);

                        // unlock row height after done
                        for (int i = 0; i < swTable.RowCount; i++)
                        {
                            swTable.SetLockRowHeight(i, false);
                        }
                    }

                    Log(sb, "Renamed bodies for " + modelName +" in configuration " +ebc.BomConfiguration);
                }
            }

            // get file save name
            string saveName = sPathName + "\\" + sFileName + "_Rev" + rev + ".xlsx";

            // save excel file
            xlWb.SaveAs(saveName,AccessMode:XlSaveAsAccessMode.xlNoChange);
            xlWb.Close(0);
            xlApp.Quit();

            releaseObject(xlWb); // excel object need to release after use
            releaseObject(xlApp); // excel object need to release after use

            Log(sb, "BOM exported.");
            return sb;
        }

        private void Log(StringBuilder sb, string text)
        {
            sb.AppendLine(DateTime.Now.ToString() + " : " + text);
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
