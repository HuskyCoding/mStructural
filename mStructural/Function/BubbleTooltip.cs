using SolidWorks.Interop.sldworks;
using System.Runtime.InteropServices;
using System;
using SolidWorks.Interop.swconst;

namespace mStructural.Function
{
    public class BubbleTooltip
    {
        #region Private Variables
        private SldWorks swApp;
        #endregion

        public BubbleTooltip(SldWorks swapp)
        {
            swApp = swapp;
        }

        public void DisplayBubbleTooltip(string title, string displaytext)
        {
            double[] curPos = GetCursorPosition();
            int cursorX = Convert.ToInt32(curPos[0]);
            int cursorY = Convert.ToInt32(curPos[1]);
            swApp.ShowBubbleTooltipAt2(cursorX, cursorY, (int)swArrowPosition.swArrowLeftOrRightTop, title, displaytext, (int)swBitMaps.swBitMapNone, "", "", 0, (int)swLinkString.swLinkStringNone, "", "");
        }

        // Method to get cursor position
        private double[] GetCursorPosition()
        {
            double[] cursorPos = new double[2];
            POINT defPnt = new POINT();
            GetCursorPos(out defPnt);
            cursorPos[0] = defPnt.X;
            cursorPos[1] = defPnt.Y;
            return cursorPos;
        }

        // struct for point
        [StructLayout(LayoutKind.Sequential)]
        struct POINT
        {
            public Int32 X;
            public Int32 Y;
        }

        // import unmanaged method to get cursor position
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool GetCursorPos(out POINT point);
    }
}
