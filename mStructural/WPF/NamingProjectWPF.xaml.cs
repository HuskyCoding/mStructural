using mStructural.Classes;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace mStructural.WPF
{
    /// <summary>
    /// Interaction logic for NamingProjectWPF.xaml
    /// </summary>
    public partial class NamingProjectWPF : Window
    {
        #region Private Variables
        private ModelDoc2 swModel;

        private ObservableCollection<NamingProjectClass> comList;
        private List<string> comNameList;

        private int asmCounter = 100;
        private int fasternerCounter = 100;

        private StringBuilder sb;

        private NamingProjectClass selectedNpc;
        #endregion

        // constructor
        public NamingProjectWPF(ModelDoc2 swmodel)
        {
            InitializeComponent();
            swModel = swmodel;
            
            // get the model name without extension and file path for bt number use
            string ModelDocName = Path.GetFileNameWithoutExtension(swModel.GetPathName());

            // set bt number text according to model name
            BtNumberTb.Text = ModelDocName.Substring(2, 5);

            // populate the tree view
            PopulateList();

            // set tree view source
            TV1.ItemsSource = comList;

            // add text change event to bt number text box 
            BtNumberTb.TextChanged += BtNumberTb_TextChanged;
        }

        #region Methods to populate initial item source
        // method to traverse tree view
        private void PopulateList()
        {
            // create new collection for main item source
            comList = new ObservableCollection<NamingProjectClass>();

            // create new name list for repetitive elimination
            comNameList = new List<string>();

            // add collection change event to main item source collection
            comList.CollectionChanged += items_CollectionChanged;

            // get active configuration to get the root component from assembly
            Configuration swConf = swModel.IGetActiveConfiguration();

            // get root component
            Component2 swComp = swConf.GetRootComponent3(true);

            // create new npc class for the root component
            NamingProjectClass npc = new NamingProjectClass();
           
            // set properties for root component
            npc.OldName = Path.GetFileNameWithoutExtension(swModel.GetPathName());
            npc.NewName = npc.OldName;
            npc.DocType = "ASM";
            npc.Include = true;
            npc.Fasterner = false;
            npc.ParentNPC = null;

            // Traverse the child components
            TraverseComponent(swComp, asmCounter, npc);

            // add this root component to tree view item source, only need the main parent, the reset will be added as the child to this main.
            comList.Add(npc);
        }

        // Method to Traverse Components
        private void TraverseComponent(Component2 swcomp, int asmcount, NamingProjectClass parentnpc)
        {
            // reset part count for each parent
            int intPartCount = 0;

            // get all childrens
            object[] vChildArr = (object[])swcomp.GetChildren();

            // if child exists, loop
            foreach (object vChild in vChildArr)
            {
                // cast the object returned from get children to component2
                Component2 swChildComp = (Component2)vChild;

                // get the model doc to this component2
                ModelDoc2 compModel = swChildComp.IGetModelDoc();

                // skip if the model is suppressed or is patterned component
                if (!swChildComp.IsSuppressed() && !swChildComp.IsPatternInstance())
                {
                    // skip if the model is read only or it is envelope
                    if (!compModel.IsOpenedReadOnly())
                    {
                        // get the model name of this child without extension
                        string docPath = Path.GetFileNameWithoutExtension(compModel.GetPathName());

                        // check if this file is already processed
                        if (!comNameList.Contains(docPath))
                        {
                            // create new npc for this child
                            NamingProjectClass npc = new NamingProjectClass();

                            // set properties
                            npc.OldName = docPath;

                            // conditioning
                            if (compModel.GetType() == (int)swDocumentTypes_e.swDocPART)
                            {
                                // if it is part file, increment the part count by 1 first
                                intPartCount++;

                                // set doc type and new name
                                npc.DocType = "PART";
                                npc.NewName = "BT"+BtNumberTb.Text+"-PT-" + (asmcount + intPartCount).ToString();
                            }
                            else if (compModel.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY)
                            {   
                                // if it is assembly, increment the count by 50
                                asmCounter += 50;

                                // set doc type and new name
                                npc.DocType = "ASM";
                                npc.NewName = "BT"+BtNumberTb.Text+"-AS-" + asmCounter.ToString();
                            }
                            
                            // set include as true and fasterner as false as default properties
                            npc.Include = true;
                            npc.Fasterner = false;

                            // set comp name
                            npc.SwComp = swChildComp;

                            // set parent class
                            npc.ParentNPC = parentnpc;

                            // add collection change event for children of this child
                            parentnpc.ChildrenNPC.CollectionChanged += items_CollectionChanged;

                            // add this child to the parent
                            parentnpc.ChildrenNPC.Add(npc);

                            // add this model name into the comparison list
                            comNameList.Add(docPath);

                            // traverse the child of this child
                            TraverseComponent(swChildComp, asmCounter, npc);
                        }
                    }
                }
            }
        }
        #endregion

        #region Methods to renumber
        // method to renumber name
        private void Renumber()
        {
            // reset asm and fasterner count
            asmCounter = 100;
            fasternerCounter = 100;

            // loop
            foreach(NamingProjectClass npc in comList)
            {
                TraverseRenumbering(npc, asmCounter);
            }
        }

        // method for renumbering traversal
        private void TraverseRenumbering(NamingProjectClass npc, int asmcount)
        {
            // set initial part count to 0 as initial value for each child
            int intPartCount = 0;

            // loop for each child
            foreach (NamingProjectClass childNpc in npc.ChildrenNPC)
            {
                // check if it is included
                if (childNpc.Include)
                {
                    if (childNpc.Fasterner)
                    {
                        // if fasterner toggle button is checked, increment the fasterner count by 10
                        fasternerCounter += 10;

                        // set new name
                        childNpc.NewName = "BT" + BtNumberTb.Text + "-FT-" + fasternerCounter;
                    }
                    else
                    {
                        // if it is not defined as fasterner
                        if (childNpc.DocType == "PART")
                        {
                            // if it is a part, increase part count by 1
                            intPartCount++;

                            // set new name
                            childNpc.NewName = "BT" + BtNumberTb.Text + "-PT-" + (asmcount + intPartCount).ToString();
                        }
                        else if(childNpc.DocType == "ASM")
                        {
                            // if it is an assembly, increase assembly count by 50
                            asmCounter += 50;

                            // set new name
                            childNpc.NewName = "BT" + BtNumberTb.Text + "-AS-" + asmCounter.ToString();
                        }
                    }
                }
                else
                {
                    // remove from rename if not included
                    childNpc.NewName = "";
                }

                // Traverse for children
                TraverseRenumbering(childNpc, asmCounter);
            }
        }
        #endregion

        #region Methods for collection changed event
        // Method that i copy from internet
        private void items_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (e.OldItems != null)
            {
                foreach (INotifyPropertyChanged item in e.OldItems)
                    item.PropertyChanged -= item_PropertyChanged;
            }
            if (e.NewItems != null)
            {
                foreach (INotifyPropertyChanged item in e.NewItems)
                    item.PropertyChanged += item_PropertyChanged;
            }
        }

        // Method that also copy from internet, but put all the collection changed logic in here
        private void item_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            // renumber the new name
            Renumber();

            // refresh the tree view data
            TV1.Items.Refresh();

            // update tree view layout
            TV1.UpdateLayout();
        }
        #endregion

        // Method for bt number changed
        private void BtNumberTb_TextChanged(object sender, TextChangedEventArgs e)
        {
            // reset counters
            asmCounter = 100;
            fasternerCounter = 100;

            // populate list
            PopulateList();

            // assign item source for tree view
            TV1.ItemsSource = comList;

            // update the tree view
            TV1.UpdateLayout();
        }

        // Method when confirm button is pressed
        private void ConfirmBtn_Click(object sender, RoutedEventArgs e)
        {
            // initiate new string builder
            sb = new StringBuilder();
            sb.Append("Completed!\r\n");
            
            // Traverse for each child in main item source
            foreach(NamingProjectClass npc in comList)
            {
                TraverseRename(npc);
            }

            // clear selection
            swModel.ClearSelection2(true);

            // close this window
            Close();

            // show log
            MessageBox.Show(sb.ToString());
        }

        // Method to traverse rename component
        private void TraverseRename(NamingProjectClass parentNpc)
        {
            // loop
            foreach(NamingProjectClass npc in parentNpc.ChildrenNPC)
            {
                // if include checkbox is checked
                if (npc.Include)
                {
                    // case to local variable, just a good practice
                    Component2 swComp = npc.SwComp;

                    // select the component
                    swComp.Select4(false, null, false);

                    // ranem
                    int iStatus = swModel.Extension.RenameDocument(npc.NewName);

                    // log
                    sb.AppendLine(npc.OldName + " renamed to " + npc.NewName + ", Status: " + iStatus);
                    
                    // go to child
                    TraverseRename(npc);
                }
            }
        }

        private void UpBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // check if selected item
                NamingProjectClass parentNpc = selectedNpc.ParentNPC;
                if(parentNpc == null)
                {
                    return;
                }

                ObservableCollection<NamingProjectClass> childCol = parentNpc.ChildrenNPC;
                int selectedItemIndex = childCol.IndexOf(selectedNpc);

                if(selectedItemIndex != 0)
                {
                    childCol.Move(selectedItemIndex, selectedItemIndex - 1);
                }

                // renumber the new name
                Renumber();

                // refresh the tree view data
                TV1.Items.Refresh();

                // update tree view layout
                TV1.UpdateLayout();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void DownBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // check if selected item
                NamingProjectClass parentNpc = selectedNpc.ParentNPC;
                if (parentNpc == null)
                {
                    return;
                }

                ObservableCollection<NamingProjectClass> childCol = parentNpc.ChildrenNPC;
                int selectedItemIndex = childCol.IndexOf(selectedNpc);

                if (selectedItemIndex != childCol.Count-1)
                {
                    childCol.Move(selectedItemIndex, selectedItemIndex +1);
                }

                // renumber the new name
                Renumber();

                // refresh the tree view data
                TV1.Items.Refresh();

                // update tree view layout
                TV1.UpdateLayout();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void TV1_Selected(object sender, RoutedEventArgs e)
        {
            TreeViewItem tvi = e.OriginalSource as TreeViewItem;
            selectedNpc = tvi.DataContext as NamingProjectClass;
        }
    }
}
