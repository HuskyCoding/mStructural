using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;

namespace mStructural.MacroFeatures
{
    public class MacroFeature : SwComFeature
    {
        public object Edit(object app, object modelDoc, object feature)
        {
            PMP myPMP = new PMP((SldWorks)app, 1, (Feature)feature);
            myPMP.Show();
            return null;
        }

        public object Regenerate(object app, object modelDoc, object feature)
        {

            return null;
        }

        public object Security(object app, object modelDoc, object feature)
        {
            return null;
        }
    }
}
