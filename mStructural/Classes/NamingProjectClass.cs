using SolidWorks.Interop.sldworks;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace mStructural.Classes
{
    public class NamingProjectClass: INotifyPropertyChanged
    {
        private string oldName;
        public string OldName
        {
            get
            {
                return oldName;
            }
            set
            {
                oldName = value;
            }
        }

        private string newName;
        public string NewName
        {
            get
            {
                return newName;
            }
            set
            {
                newName = value;
            }
        }

        private string docType;
        public string DocType
        {
            get
            {
                return docType;
            }
            set
            {
                docType = value;
            }
        }

        private Component2 swComp;
        public Component2 SwComp
        {
            get
            {
                return swComp;
            }
            set
            {
                swComp = value;
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
                IncludeChanged();
            }
        }

        private bool fasterner;
        public bool Fasterner
        {
            get
            {
                return fasterner;
            }
            set
            {
                fasterner = value;
                OnPropertyChanged();
            }
        }

        private bool isSelected;
        public bool IsSelected
        {
            get
            {
                return isSelected;
            }
            set
            {
                isSelected = value;
            }
        }


        // Method to check on all child if parent is checked
        private void IncludeChanged()
        {
            foreach(NamingProjectClass child in ChildrenNPC)
            {
                child.Include = Include;
            }
        }

        private ObservableCollection<NamingProjectClass> childrenNPC = new ObservableCollection<NamingProjectClass>();
        public ObservableCollection<NamingProjectClass> ChildrenNPC
        {
            get
            {
                return childrenNPC;
            }
            set
            {
                childrenNPC = value;
            }
        }

        private NamingProjectClass parentNPC;
        public NamingProjectClass ParentNPC
        {
            get
            {
                return parentNPC;
            }
            set
            {
                parentNPC = value;
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
