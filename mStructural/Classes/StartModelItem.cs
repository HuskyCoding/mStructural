using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace mStructural.Classes
{
    public class StartModelItem : INotifyPropertyChanged
    {
        private int itemNo;
        public int ItemNo
        {
            get
            {
                return itemNo;
            }
            set
            {
                itemNo = value;
            }
        }

        private string originalFileName;
        public string OriginalFileName
        {
            get
            {
                return originalFileName;
            }
            set
            {
                originalFileName = value;
            }
        }

        private string newFileName;
        public string NewFileName
        {
            get
            {
                return newFileName;
            }
            set
            {
                newFileName = value;
            }
        }

        private bool isTopLevel;
        public bool IsTopLevel
        {
            get
            {
                return isTopLevel;
            }
            set
            {
                isTopLevel = value;
            }
        }

        private bool include;
        public bool Include
        {
            get
            {
                return include;
            }
            set
            {
                include = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string propertyName="")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
