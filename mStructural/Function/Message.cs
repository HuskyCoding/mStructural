using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace mStructural.Function
{
    public class Message
    {
        #region Public Variables
        public string ModelDocNull = "No active document found!";
        public string NotDrawingDoc = "Only support drawing document!";
        public string NotPartDoc = "Only support part document!";
        public string NotAssemDoc = "Only support assembly document!";
        public string NotPartNorAssemDoc = "Only support part or assembly document!";
        #endregion

        #region Private Variables
        private SldWorks swApp;
        #endregion

        // constructor
        public Message(SldWorks swapp)
        {
            swApp = swapp;
        }

        public void InfoMsg(string message)
        {
            swApp.SendMsgToUser2(message, (int)swMessageBoxIcon_e.swMbInformation, (int)swMessageBoxBtn_e.swMbOk);
        }

        public int WarnMsg(string message)
        {
            int result;
            result = swApp.SendMsgToUser2(message, (int)swMessageBoxIcon_e.swMbWarning, (int)swMessageBoxBtn_e.swMbOkCancel);
            return result;
        }

        public void ErrorMsg(string message)
        {
            swApp.SendMsgToUser2(message, (int)swMessageBoxIcon_e.swMbStop, (int)swMessageBoxBtn_e.swMbOk);
        }
    }
}
