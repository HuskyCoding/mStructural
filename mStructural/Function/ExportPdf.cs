using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;

namespace mStructural.Function
{
    public class ExportPdf
    {
        #region Private Variables
        private SldWorks swApp;
        private Macros macros;
        private Message msg;
        private SwErrors swErrors;
        #endregion
        
        // constructor
        public ExportPdf(SWIntegration swintegration)
        {
            swApp = swintegration.SwApp;
            macros = new Macros(swApp);
            msg = new Message(swApp);
            swErrors = new SwErrors();
        }

        // main method
        public void Run()
        {
            // get the current modeldoc
            ModelDoc2 swModel = swApp.IActiveDoc2;

            // check if any active doc or the type is drawing
            if (!macros.checkDocNullandType(swModel, swDocumentTypes_e.swDocDRAWING)) return;

            // show confirm export message
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

            // warning message before run
            // if (msg.WarnMsg("Press OK to save the drawing as PDF.") != 4) return;

            // initialize custom prop manager
            CustomPropertyManager swCusPropMgr = swModel.Extension.CustomPropertyManager[""];

            // get revision custom prop
            string val = "";
            string valout = "";
            bool bRet = swCusPropMgr.Get4("Revision", false, out val, out valout);

            // get current model name
            string sFileName = Path.GetFileNameWithoutExtension(swModel.GetPathName());

            // get current path name
            string sPathName = Path.GetDirectoryName(swModel.GetPathName());

            // find the \
            int index = sPathName.LastIndexOf("\\");

            // add production in path
            if (toProduction == true)
            {
                sPathName = sPathName.Substring(0, index) + "\\Production";
            }

            // create directory
            Directory.CreateDirectory(sPathName);

            // get file save name
            string saveName = sPathName + "\\" + sFileName + "_Rev" + valout.ToUpper() + ".PDF";

            // get export data
            ExportPdfData swEpd = (ExportPdfData)swApp.GetExportFileData((int)swExportDataFileType_e.swExportPdfData);

            // cast drawing doc
            DrawingDoc swDraw = (DrawingDoc)swModel;

            // get all sheet
            string[] obj = (string[])swDraw.GetSheetNames();

            // create new list without dxf in the sheet name
            List<string> targetSheet = new List<string>();
            foreach (string s in obj)
            {
                if (!s.ToLower().Contains("dxf"))
                {
                    targetSheet.Add(s);
                }
            }

            // initialize dispatch wrapper and intermediate object
            object[] objs = new object[targetSheet.Count];
            DispatchWrapper[] arrObjIn = new DispatchWrapper[targetSheet.Count];

            for(int i=0;i< targetSheet.Count; i++)
            {
                bRet = swDraw.ActivateSheet((targetSheet[i]));
                Sheet swSheet = swDraw.IGetCurrentSheet();
                objs[i] = swSheet;
                arrObjIn[i] = new DispatchWrapper(objs[i]);
            }

            // save the drawings to pdf
            int err = -1, warn = -1;
            bRet = swEpd.SetSheets((int)swExportDataSheetsToExport_e.swExportData_ExportSpecifiedSheets, (arrObjIn));
            swEpd.ViewPdfAfterSaving = true;
            bRet = swModel.Extension.SaveAs(saveName, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, swEpd, ref err, ref warn);

            // if error
            if (!bRet)
            {
                msg.ErrorMsg(swErrors.GetFileSaveError(err));
            }
            else
            {
                msg.InfoMsg("Export PDF completed.");
            }
        }
    }
}
