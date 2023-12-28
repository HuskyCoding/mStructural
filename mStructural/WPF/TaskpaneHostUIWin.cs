using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace mStructural
{
    [ProgId(SWIntegration.SWTASKPANE_PROGID)]
    public partial class TaskpaneHostUIWin : Form
    {
        public TaskpaneHostUIWin(SWIntegration swintegration)
        {
            InitializeComponent();

            // Instantiate ui
            TaskpaneHostUI uc = new TaskpaneHostUI(swintegration);

            // assign child to eh
            elementHost.Child = uc;
        }
    }
}
