using mStructural.Classes;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace mStructural.Function
{
    public class ExportPdf2
    {
        public StringBuilder Run(SldWorks swapp, string drawingpath, bool isproduction, ObservableCollection<ExportPdfClass> pdfcol, string rev)
        {
            StringBuilder sb = new StringBuilder();

            // open drawing doc
            int err = -1;
            int warn = -1;
            ModelDoc2 swModel = swapp.OpenDoc6(drawingpath, (int)swDocumentTypes_e.swDocDRAWING, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref err, ref warn);

            // cast to drawing doc
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

            // get file save name
            string saveName = sPathName + "\\" + sFileName + "_Rev" + rev + ".PDF";

            // get export data
            ExportPdfData swEpd = (ExportPdfData)swapp.GetExportFileData((int)swExportDataFileType_e.swExportPdfData);

            // initialize dispatch wrapper and intermediate object
            List<DispatchWrapper> dwList = new List<DispatchWrapper>();

            foreach (ExportPdfClass epc in pdfcol)
            {
                if (epc.Include)
                {
                    Sheet swSheet = swDraw.Sheet[epc.SheetName];
                    DispatchWrapper dw = new DispatchWrapper(swSheet);
                    dwList.Add(dw);
                }
            }

            DispatchWrapper[] dwArr = dwList.ToArray();

            // save the drawings to pdf
            bool bRet = swEpd.SetSheets((int)swExportDataSheetsToExport_e.swExportData_ExportSpecifiedSheets, (dwArr));
            swEpd.ViewPdfAfterSaving = true;
            bRet = swModel.Extension.SaveAs(saveName, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, swEpd, ref err, ref warn);

            // if error
            if (!bRet)
            {
                Log(sb, "Saving pdf error with error code: " + err.ToString());
            }
            else
            {
                Log(sb, "Exported PDF...");
            }

            return sb;
        }

        private void Log(StringBuilder sb, string text)
        {
            sb.AppendLine(DateTime.Now.ToString() + " : " + text);
        }
    }
}
