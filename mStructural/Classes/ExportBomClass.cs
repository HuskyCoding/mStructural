using SolidWorks.Interop.sldworks;
using System.Collections.ObjectModel;

namespace mStructural.Classes
{
    public class ExportBomClass
    {
        public string TableName { get; set; }
        public TableAnnotation SwTableAnn { get; set; } 
        public ModelDoc2 SwModel { get; set; }
        public string BalloonRef { get; set; }
        public bool ExportTubeLaser { get; set; }
        public string BomConfiguration { get; set; }
        public string TubeLaserConfiguration { get; set; }
        public ObservableCollection<string> TubeLaserConfigurationList { get; set; }
    }
}
