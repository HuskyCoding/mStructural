using mStructural.Function;
using SolidWorks.Interop.sldworks;
using System.Collections.Generic;
using System.Windows;

namespace mStructural.WPF
{
    /// <summary>
    /// Interaction logic for HideRefWPF.xaml
    /// </summary>
    public partial class HideRefWPF : Window
    {
        #region Private Variables
        private SldWorks swApp;
        #endregion

        public HideRefWPF(SldWorks swapp)
        {
            swApp = swapp;
            InitializeComponent();
        }

        private void OkBtn_Click(object sender, RoutedEventArgs e)
        {
            HideRef hideRef = new HideRef();

            List<string> hideTypeList = new List<string>();
            if (OriginCh.IsChecked == true) hideTypeList.Add("OriginProfileFeature");
            if (PlaneCh.IsChecked == true) hideTypeList.Add("RefPlane");
            if (AxeCh.IsChecked == true) hideTypeList.Add("RefAxis");
            if (PointCh.IsChecked == true) hideTypeList.Add("RefPoint");
            if (CoorCh.IsChecked == true) hideTypeList.Add("CoordSys");
            if (Sk2dCh.IsChecked == true) hideTypeList.Add("ProfileFeature");
            if (Sk3dCh.IsChecked == true) hideTypeList.Add("3DProfileFeature");
            if (PtCurveCh.IsChecked == true) hideTypeList.Add("3DSplineCurve");
            if (CompCurveCh.IsChecked == true) hideTypeList.Add("CompositeCurve");
            if (HelixCh.IsChecked == true) hideTypeList.Add("Helix");

            hideRef.Run(swApp, hideTypeList);

            Close();
        }

        private void DeselectBtn_Click(object sender, RoutedEventArgs e)
        {
            OriginCh.IsChecked = false;
            PlaneCh.IsChecked = false;
            AxeCh.IsChecked = false;
            PointCh.IsChecked = false;
            CoorCh.IsChecked = false;
            Sk2dCh.IsChecked = false;
            Sk3dCh.IsChecked = false;
            PtCurveCh.IsChecked = false;
            CompCurveCh.IsChecked = false;
            HelixCh.IsChecked = false;
        }

        private void SelectBtn_Click(object sender, RoutedEventArgs e)
        {
            OriginCh.IsChecked = true;
            PlaneCh.IsChecked = true;
            AxeCh.IsChecked = true;
            PointCh.IsChecked = true;
            CoorCh.IsChecked = true;
            Sk2dCh.IsChecked = true;
            Sk3dCh.IsChecked = true;
            PtCurveCh.IsChecked = true;
            CompCurveCh.IsChecked = true;
            HelixCh.IsChecked = true;
        }
    }
}
