using mStructural.Classes;
using mStructural.Function;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;

namespace mStructural.WPF
{
    /// <summary>
    /// Interaction logic for StartModelMenu.xaml
    /// </summary>

    public partial class StartModelMenu : Window
    {
        #region Public Variables
        public ObservableCollection<StartModelItem> SmCol {get;set;}
        #endregion

        #region Private Variables
        private SldWorks swApp;
        private string StartModelPath;
        private BubbleTooltip bubbleTT;
        #endregion

        public StartModelMenu(SldWorks swapp)
        {
            InitializeComponent();

            // grab the solidworks app object
            swApp = swapp;

            // instantiate macros
            bubbleTT = new BubbleTooltip(swApp);

            // populate start model list at \Library\02-DESIGN LIBRARY\BT LIBRARY\START MODELS
            appsetting appset = new appsetting();
            StartModelPath = appset.StartModelPath;

            // get all start model in the pdm location
            string[] startModelList = Directory.GetDirectories(StartModelPath, "*", SearchOption.TopDirectoryOnly);

            // populate the start model to dropdown list
            foreach (string startModel in startModelList)
            {
                string folderName = new DirectoryInfo(startModel).Name;
                startModelListCb.Items.Add(folderName);
            }
        }

        private void startModelListCb_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            try
            {
                // get all part in the list
                string[] smItemStrs = Directory.GetFiles(StartModelPath + "\\" + startModelListCb.SelectedValue.ToString());

                // put into hashtable
                Hashtable smItems = new Hashtable();
                foreach(string smItem in smItemStrs)
                {
                    smItems.Add(Path.GetFileName(smItem), smItem);
                }

                // get all dependencies for toplevel checking
                Hashtable allDependencies = getDependancies(smItemStrs);

                SmCol = new ObservableCollection<StartModelItem>();
                SmCol.CollectionChanged += items_CollectionChanged;
                int i = 1;
                int instanceNo;

                // get the instance number from text box
                bool bRet = int.TryParse(ModelInstTb.Text, out instanceNo);

                foreach (string smItemStr in smItemStrs)
                {
                    StartModelItem smItem = new StartModelItem();

                    // check if it is toplevel
                    smItem.IsTopLevel = !allDependencies.ContainsValue(smItemStr);

                    // if top level, then loop through structure
                    if( smItem.IsTopLevel)
                    {
                        string filename = Path.GetFileName(smItemStr);
                        smItem.ItemNo = i;
                        smItem.OriginalFileName = filename;

                        // check if it is part or assembly
                        string partType = "";
                        string ext = "";
                        if (filename.ToLower().Contains(".sldasm"))
                        {
                            partType = "AS";
                            ext = ".sldasm";
                        }
                        else if (filename.ToLower().Contains(".sldprt"))
                        {
                            partType = "PT";
                            ext = ".sldprt";
                        }
                    
                        int drawingNo;
                        int.TryParse(DrawNoTB.Text,out drawingNo);
                        smItem.NewFileName = "BT" + drawingNo.ToString("00000") + "-" + partType + "-" + instanceNo.ToString("000") + ext;

                        smItem.Include = true;
                        SmCol.Add(smItem);
                        i++;

                        // get all dependecies for this file only
                        Hashtable dependencies = getDependancy(smItemStr);
                        if(dependencies.Count > 0)
                        {
                            foreach(DictionaryEntry entry in dependencies)
                            {
                                if(smItems.ContainsValue(entry.Value.ToString()))
                                {
                                    instanceNo++;
                                    GenerateName(entry, instanceNo,i);
                                    i++;
                                }
                            }
                        }
                    }
                }

                DG1.ItemsSource = SmCol;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void BrowseBtn_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog fbd = new System.Windows.Forms.FolderBrowserDialog();
            System.Windows.Forms.DialogResult result = fbd.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                OutputLocTb.Text = fbd.SelectedPath;
            }
        }

        private void GenerateBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (OutputLocTb.Text == "")
                {
                    MessageBox.Show("Output Location is empty.");
                    return;
                }

                if (!Directory.Exists(OutputLocTb.Text))
                {
                    MessageBox.Show("Output directory does not exist.");
                    return;
                }

                if (SmCol == null)
                {
                    MessageBox.Show("Collection is empty");
                    return;
                }

                string topLevelAsm = "";

                foreach (StartModelItem sr in SmCol)
                {
                    if (sr.Include)
                    {
                        // copy file
                        string sourceFile = StartModelPath + "\\" + startModelListCb.Text + "\\" + sr.OriginalFileName;
                        string destFile = OutputLocTb.Text + "\\" + sr.NewFileName;
                        int iRet = swApp.CopyDocument(sourceFile, destFile, "", "", (int)swMoveCopyOptions_e.swMoveCopyOptionsOverwriteExistingDocs);

                        // set read-only flag to off for newly copied files
                        FileInfo fileInfo = new FileInfo(destFile);
                        fileInfo.IsReadOnly = false;

                        // check the column if it is an assembly
                        if (sr.IsTopLevel)
                        {
                            topLevelAsm = OutputLocTb.Text + "\\" + sr.NewFileName;
                        }
                    }
                }

                // replace reference
                foreach (StartModelItem sr in SmCol)
                {
                    if (!sr.IsTopLevel && sr.Include)
                    {
                        bool bRet = swApp.ReplaceReferencedDocument(topLevelAsm, StartModelPath + "\\" + startModelListCb.Text + "\\" + sr.OriginalFileName, OutputLocTb.Text + "\\" + sr.NewFileName);
                    }
                }

                Close();
                bubbleTT.DisplayBubbleTooltip("Completed", "Start Model generated!");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private Hashtable getDependancies(string[] allFiles)
        {
            Hashtable allDependencies = new Hashtable();
            // get all dependency and make a unique list
            foreach (string file in allFiles)
            {
                string[] depends = (string[])swApp.GetDocumentDependencies2(file, true, true, false);
                if (depends != null)
                {
                    int index = 0;
                    while (index < depends.GetUpperBound(0))
                    {
                        try
                        {
                            allDependencies.Add(depends[index], depends[index + 1]);
                        }
                        catch { }
                        index += 2;
                    }
                }
            }

            return allDependencies;
        }

        private Hashtable getDependancy(string file)
        {
            Hashtable dependencies = new Hashtable();
            string[] depends = (string[])swApp.GetDocumentDependencies2(file, true, true, false);
            if (depends != null)
            {
                int index = 0;
                while (index < depends.GetUpperBound(0))
                {
                    try
                    {
                        dependencies.Add(depends[index], depends[index + 1]);
                    }
                    catch { }
                    index += 2;
                }
            }

            return dependencies;
        }

        private void DrawNoTB_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            e.Handled = !IsTextAllowed(e.Text);
        }

        private void ModelInstTb_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            e.Handled = !IsTextAllowed(e.Text);
        }

        private static readonly Regex _regex = new Regex("[^0-9]+"); //regex that matches disallowed text
        private static bool IsTextAllowed(string text)
        {
            return !_regex.IsMatch(text);
        }

        private void DrawNoTB_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            if (SmCol != null)
            {
                foreach(StartModelItem smItem in SmCol)
                {
                    if (smItem.Include||smItem.IsTopLevel)
                    {
                        StringBuilder sb = new StringBuilder(smItem.NewFileName);
                        sb.Remove(0, 7);
                        int drawingNo;
                        int.TryParse(DrawNoTB.Text, out drawingNo);
                        sb.Insert(0, "BT" + drawingNo.ToString("00000"));
                        smItem.NewFileName = sb.ToString();
                    }
                }

                DG1.Items.Refresh();
            }
        }

        private void ModelInstTb_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            if (SmCol != null)
            {
                int index = 1;
                foreach (StartModelItem smItem in SmCol)
                {
                    if (smItem.IsTopLevel)
                    {
                        StringBuilder sb = new StringBuilder(smItem.NewFileName);
                        sb.Remove(11, 3);
                        int ModelInstNo;
                        int.TryParse(ModelInstTb.Text, out ModelInstNo);
                        sb.Insert(11, ModelInstNo.ToString("000"));
                        smItem.NewFileName = sb.ToString();
                    }
                    else
                    {
                        if (smItem.Include)
                        {
                            StringBuilder sb = new StringBuilder(smItem.NewFileName);
                            sb.Remove(11, 3);
                            int ModelInstNo;
                            int.TryParse(ModelInstTb.Text, out ModelInstNo);
                            ModelInstNo += index;
                            sb.Insert(11, ModelInstNo.ToString("000"));
                            smItem.NewFileName = sb.ToString();
                            index++;
                        }
                    }
                }

                DG1.Items.Refresh();
            }
        }

        private void RegenerateName()
        {
            bool bRet;
            int drawingNo;
            int instanceNo;
            bRet = int.TryParse(DrawNoTB.Text, out drawingNo);
            bRet = int.TryParse(ModelInstTb.Text, out instanceNo);

            string partType = "'";
            string ext = "";

            foreach (StartModelItem smItem in SmCol)
            {
                if (!smItem.IsTopLevel)
                {
                    if (smItem.Include)
                    {
                        // check if it is part or assembly
                        if (smItem.OriginalFileName.ToLower().Contains(".sldasm"))
                        {
                            partType = "AS";
                            ext = ".sldasm";
                        }
                        else if (smItem.OriginalFileName.ToLower().Contains(".sldprt"))
                        {
                            partType = "PT";
                            ext = ".sldprt";
                        }
                        instanceNo++;
                        smItem.NewFileName = smItem.NewFileName = "BT" + drawingNo.ToString("00000") + "-" + partType + "-" + instanceNo.ToString("000") + ext;
                    }
                    else
                    {
                        smItem.NewFileName = "";
                    }
                }
            }
        }

        private void GenerateName(DictionaryEntry entry, int instanceNo, int itemNo)
        {
            StartModelItem smItem = new StartModelItem();
            string filename = Path.GetFileName(entry.Value.ToString());
            smItem.ItemNo = itemNo;
            smItem.OriginalFileName = filename;
            smItem.IsTopLevel = false;

            string partType= "";
            string ext = "";
            int drawingNo;

            // check if it is part or assembly
            if (filename.ToLower().Contains(".sldasm"))
            {
                partType = "AS";
                ext = ".sldasm";
            }
            else if (filename.ToLower().Contains(".sldprt"))
            {
                partType = "PT";
                ext = ".sldprt";
            }
            int.TryParse(DrawNoTB.Text, out drawingNo);
            smItem.NewFileName = "BT" + drawingNo.ToString("00000") + "-" + partType + "-" + instanceNo.ToString("000") + ext;

            smItem.Include = true;
            SmCol.Add(smItem);
        }

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

        private void item_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            RegenerateName();
            DG1.ItemsSource = null;
            DG1.ItemsSource = SmCol;
        }
    }
}
