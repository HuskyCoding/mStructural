using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Collections.Generic;
using System.Linq;

namespace mStructural.Function
{
    // class for sorting
    class GFG : IComparer<string>
    {
        public int Compare(string x, string y)
        {
            if (x == null || y == null)
            {
                return 0;
            }

            // "CompareTo()" method
            return x.CompareTo(y);

        }
    }

    public class SortCutlist
    {
        #region Private Variables
        private SldWorks swApp = default (SldWorks);
        private Message msg;
        private Macros macros;
        #endregion

        // constructor
        public SortCutlist(SWIntegration swintegration)
        {
            swApp = swintegration.SwApp;
            macros = new Macros(swApp);
            msg = new Message(swApp);
        }

        // Run function
        public void Run()
        {
            ModelDoc2 swModel = default (ModelDoc2);
            swModel = (ModelDoc2)swApp.ActiveDoc;

            bool bRet = macros.checkDocNullandType(swModel, swDocumentTypes_e.swDocDRAWING); if (!bRet) return; // check

            SelectionMgr swSelMgr = default (SelectionMgr);
            swSelMgr = (SelectionMgr)swModel.SelectionManager;

            TableAnnotation clTable = default (TableAnnotation);

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

            // get description column
            int desColNo = -1;
            int lenColNo = -1;
            int rowCount;
            int colCount;
            int i;

            rowCount = clTable.RowCount;
            colCount = clTable.ColumnCount;

            // get description column
            for (i=0;i<colCount;i++)
            {
                string columnTitle = clTable.DisplayedText2[0, i, false];
                if(columnTitle.ToLower() == "description")
                {
                    desColNo= i;
                    break;
                }
            }

            // get length column
            for (i = 0; i < colCount; i++)
            {
                string columnTitle = clTable.DisplayedText2[0, i, false];
                if (columnTitle.ToLower() == "length")
                {
                    lenColNo = i;
                    break;
                }
            }

            // check if description column exist
            if (desColNo == -1)
            {
                msg.ErrorMsg("There is no description column in this weldment cutlist");
                return;
            }

            // check if length column exist
            if (lenColNo == -1)
            {
                msg.ErrorMsg("There is no length column in this weldment cutlist");
                return;
            }

            // get all desctiption in description column and create a distinct list
            List<string> desListDup = new List<string>();
            List<string> desList = new List<string>();
            for (i=0;i<rowCount; i++)
            {
                desListDup.Add(clTable.DisplayedText2[i,desColNo,false]);
            }
            desList = desListDup.Distinct().ToList();

            // create lists for sorting
            List<string> rhsList = new List<string>();
            List<string> shsList = new List<string>();
            List<string> chsList = new List<string>();
            List<string> tList = new List<string>();
            List<string> iList = new List<string>();
            List<string> angleList = new List<string>();
            List<string> mrList = new List<string>();
            List<string> rbList = new List<string>();
            List<string> pfcList = new List<string>();
            List<string> ubList = new List<string>();
            List<string> ucList = new List<string>();
            List<string> cdsList = new List<string>();
            List<string> cList = new List<string>();
            List<string> fbList = new List<string>();
            List<string> plateList = new List<string>();
            List<string> etcList = new List<string>();
            int tpInt = 0;

            // categorise distinct description
            foreach (string des in desList)
            {
                if(des.ToLower() == "description")
                {
                    continue;
                }

                if(des.Length>11)
                {
                    if (des.ToLower().Substring(des.Length - 3) == "rhs")
                    {
                        rhsList.Add(des);
                    }
                    else if (des.ToLower().Substring(des.Length - 3) == "shs")
                    {
                        shsList.Add(des);
                    }
                    else if (des.ToLower().Substring(des.Length - 3) == "chs")
                    {
                        chsList.Add(des);
                    }
                    else if (des.ToLower().Substring(0, 4) == "tube")
                    {
                        tList.Add(des);
                    }
                    else if (des.ToLower().Substring(0, 6) == "i beam")
                    {
                        iList.Add(des);
                    }
                    else if (des.ToLower().Substring(0, 5) == "angle")
                    {
                        angleList.Add(des);
                    }
                    else if (des.ToLower().Substring(0, 11) == "machine rod")
                    {
                        mrList.Add(des);
                    }
                    else if (des.ToLower().Substring(0, 9) == "round bar")
                    {
                        rbList.Add(des);
                    }
                    else if (des.ToLower().Substring(des.Length - 3) == "pfc")
                    {
                        pfcList.Add(des);
                    }
                    else if (des.ToLower().Substring(3, 2) == "ub" && int.TryParse(des.Substring(0, 3), out tpInt))
                    {
                        ubList.Add(des);
                    }
                    else if (des.ToLower().Substring(3, 2) == "uc" && int.TryParse(des.Substring(0, 3), out tpInt))
                    {
                        ucList.Add(des);
                    }
                    else if (des.ToLower().Substring(des.Length - 3) == "cds")
                    {
                        cdsList.Add(des);
                    }
                    else if (des.ToLower().Substring(0, 1) == "c" && int.TryParse(des.Substring(1, 1), out tpInt))
                    {
                        cList.Add(des);
                    }
                    else if (des.ToLower().Substring(0, 2) == "fb")
                    {
                        fbList.Add(des);
                    }
                    else if (des.ToLower().Substring(0,5) == "plate")
                    {
                        plateList.Add(des);
                    }
                    else
                    {
                        etcList.Add(des);
                    }
                }            
                else if (des.Length > 9)
                {
                    if (des.ToLower().Substring(des.Length - 3) == "rhs")
                    {
                        rhsList.Add(des);
                    }
                    else if (des.ToLower().Substring(des.Length - 3) == "shs")
                    {
                        shsList.Add(des);
                    }
                    else if (des.ToLower().Substring(des.Length - 3) == "chs")
                    {
                        chsList.Add(des);
                    }
                    else if (des.ToLower().Substring(0, 4) == "tube")
                    {
                        tList.Add(des);
                    }
                    else if (des.ToLower().Substring(0, 6) == "i beam")
                    {
                        iList.Add(des);
                    }
                    else if (des.ToLower().Substring(0, 5) == "angle")
                    {
                        angleList.Add(des);
                    }
                    else if (des.ToLower().Substring(0, 9) == "round bar")
                    {
                        rbList.Add(des);
                    }
                    else if (des.ToLower().Substring(des.Length - 3) == "pfc")
                    {
                        pfcList.Add(des);
                    }
                    else if(des.ToLower().Substring(3,2)=="ub"&& int.TryParse(des.Substring(0,3),out tpInt))
                    {
                        ubList.Add(des);
                    }
                    else if (des.ToLower().Substring(3, 2) == "uc" && int.TryParse(des.Substring(0, 3), out tpInt))
                    {
                        ucList.Add(des);
                    }
                    else if (des.ToLower().Substring(des.Length - 3) == "cds")
                    {
                        cdsList.Add(des);
                    }
                    else if (des.ToLower().Substring(0, 1) == "c" && int.TryParse(des.Substring(1, 1), out tpInt))
                    {
                        cList.Add(des);
                    }
                    else if (des.ToLower().Substring(0, 2) == "fb")
                    {
                        fbList.Add(des);
                    }
                    else if (des.ToLower().Substring(0, 5) == "plate")
                    {
                        plateList.Add(des);
                    }
                    else
                    {
                        etcList.Add(des);
                    }
                }
                else if (des.Length > 6)
                {
                    if (des.ToLower().Substring(des.Length - 3) == "rhs")
                    {
                        rhsList.Add(des);
                    }
                    else if (des.ToLower().Substring(des.Length - 3) == "shs")
                    {
                        shsList.Add(des);
                    }
                    else if (des.ToLower().Substring(des.Length - 3) == "chs")
                    {
                        chsList.Add(des);
                    }
                    else if (des.ToLower().Substring(0, 4) == "tube")
                    {
                        tList.Add(des);
                    }
                    else if (des.ToLower().Substring(0, 6) == "i beam")
                    {
                        iList.Add(des);
                    }
                    else if (des.ToLower().Substring(0, 5) == "angle")
                    {
                        angleList.Add(des);
                    }
                    else if (des.ToLower().Substring(des.Length - 3) == "pfc")
                    {
                        pfcList.Add(des);
                    }
                    else if (des.ToLower().Substring(3, 2) == "ub" && int.TryParse(des.Substring(0, 3), out tpInt))
                    {
                        ubList.Add(des);
                    }
                    else if (des.ToLower().Substring(3, 2) == "uc" && int.TryParse(des.Substring(0, 3), out tpInt))
                    {
                        ucList.Add(des);
                    }
                    else if (des.ToLower().Substring(des.Length - 3) == "cds")
                    {
                        cdsList.Add(des);
                    }
                    else if (des.ToLower().Substring(0, 1) == "c" && int.TryParse(des.Substring(1, 1), out tpInt))
                    {
                        cList.Add(des);
                    }
                    else if (des.ToLower().Substring(0, 2) == "fb")
                    {
                        fbList.Add(des);
                    }
                    else
                    {
                        etcList.Add(des);
                    }
                }
                else
                {
                    etcList.Add(des);
                }
            }

            // sort the list
            GFG gg = new GFG();
            rhsList.Sort(gg);
            shsList.Sort(gg);
            chsList.Sort(gg);
            tList.Sort(gg);
            iList.Sort(gg);
            angleList.Sort(gg);
            mrList.Sort(gg);
            rbList.Sort(gg);
            pfcList.Sort(gg);
            ubList.Sort(gg);
            ucList.Sort(gg);
            cdsList.Sort(gg);
            cList.Sort(gg);
            fbList.Sort(gg);
            plateList.Sort(gg);
            etcList.Sort(gg);

            int lastProRow = 1;

            // process lists
            ProcessList(rhsList, clTable, desColNo, lenColNo, lastProRow, out lastProRow);
            ProcessList(shsList, clTable, desColNo, lenColNo, lastProRow, out lastProRow);
            ProcessList(chsList, clTable, desColNo, lenColNo, lastProRow, out lastProRow);
            ProcessList(tList, clTable, desColNo, lenColNo, lastProRow, out lastProRow);
            ProcessList(iList, clTable, desColNo, lenColNo, lastProRow, out lastProRow);
            ProcessList(angleList, clTable, desColNo, lenColNo, lastProRow, out lastProRow);
            ProcessList(mrList, clTable, desColNo, lenColNo, lastProRow, out lastProRow);
            ProcessList(rbList, clTable, desColNo, lenColNo, lastProRow, out lastProRow);
            ProcessList(pfcList, clTable, desColNo, lenColNo, lastProRow, out lastProRow);
            ProcessList(ubList, clTable, desColNo, lenColNo, lastProRow, out lastProRow);
            ProcessList(ucList, clTable, desColNo, lenColNo, lastProRow, out lastProRow);
            ProcessList(cdsList, clTable, desColNo, lenColNo, lastProRow, out lastProRow);
            ProcessList(cList, clTable, desColNo, lenColNo, lastProRow, out lastProRow);
            ProcessList(fbList, clTable, desColNo, lenColNo, lastProRow, out lastProRow);
            ProcessList2(plateList, clTable, desColNo, lastProRow, out lastProRow);
            ProcessList2(etcList, clTable, desColNo, lastProRow, out lastProRow);

            // rebuild after rearrange
            swModel.ForceRebuild3(false);
        }

        // Function to process lists
        private void ProcessList(List<string> list, TableAnnotation table, int descolno, int lencolno, int curprorow, out int lastprorow)
        {       
            int curProRow = curprorow;
            int i;
            int j;
            bool bRet = false;

            // get lowest length
            int lowestLenRow = -1;
            foreach (string item in list)
            {
                double lowestLen = 0;
                // get lowest length
                for (i = 0; i < table.RowCount; i++)
                {
                    if (table.DisplayedText2[i, descolno, false] == item)
                    {
                        string rowLenStr = table.DisplayedText2[i, lencolno, false];
                        double rowLen = 0;
                        if (double.TryParse(rowLenStr, out rowLen))
                        {
                            if (lowestLen == 0)
                            {
                                lowestLen = rowLen;
                                lowestLenRow = i;
                            }
                            else if (rowLen < lowestLen)
                            {
                                lowestLen = rowLen;
                                lowestLenRow = i;
                            }
                        }
                    }
                }

                // move lowest length to first row
                bRet = table.MoveRow(lowestLenRow, (int)swTableItemInsertPosition_e.swTableItemInsertPosition_Before, curProRow);

                // sort the rest
                for (i = curProRow + 1; i < table.RowCount; i++)
                {
                    if (table.DisplayedText2[i, descolno, false] == item)
                    {
                        string comRowLenStr = table.DisplayedText2[i, lencolno, false];
                        double comRowLen = 0;
                        string curRowLenStr = table.DisplayedText2[curProRow, lencolno, false];
                        double curRowLen = 0;
                        if (double.TryParse(curRowLenStr, out curRowLen))
                        {
                            if (double.TryParse(comRowLenStr, out comRowLen))
                            {
                                if (comRowLen >= curRowLen)
                                {
                                    bRet = table.MoveRow(i, (int)swTableItemInsertPosition_e.swTableItemInsertPosition_After, curProRow);
                                }
                                else
                                {
                                    for (j = 1; j < curProRow; j++)
                                    {
                                        curRowLenStr = table.DisplayedText2[curProRow - j, lencolno, false];
                                        if (double.TryParse(curRowLenStr, out curRowLen))
                                        {
                                            if (comRowLen >= curRowLen)
                                            {
                                                bRet = table.MoveRow(i, (int)swTableItemInsertPosition_e.swTableItemInsertPosition_After, curProRow - j);
                                                break;
                                            }
                                        }
                                    }
                                }
                                curProRow += 1;
                            }
                        }
                    }
                }
                curProRow += 1;
            }

            lastprorow = curProRow;
        }

        // Function to process lists without length properties
        private void ProcessList2(List<string> list, TableAnnotation table, int descolno, int curprorow, out int lastprorow)
        {
            int curProRow = curprorow;
            int i;
            bool bRet = false;
            foreach (string item in list)
            {
                for(i=0; i<table.RowCount; i++)
                {
                    if (table.DisplayedText2[i, descolno, false] == item)
                    {
                        bRet = table.MoveRow(i, (int)swTableItemInsertPosition_e.swTableItemInsertPosition_Before, curProRow);
                        curProRow += 1;
                    }
                }
            }
            lastprorow = curProRow;
        }
    }
}
