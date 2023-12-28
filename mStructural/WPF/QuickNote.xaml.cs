using mStructural.Function;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Windows;

namespace mStructural.WPF
{
    /// <summary>
    /// Interaction logic for QuickNote.xaml
    /// </summary>
    public partial class QuickNote : Window
    {
        #region Private Variables
        SldWorks swApp;
        Macros macros;
        Message msg;
        appsetting appset;
        string edgeString;
        #endregion

        // constructor
        public QuickNote(SWIntegration swintegration)
        {
            // initialize method and component
            InitializeComponent();
            swApp = swintegration.SwApp;
            macros = new Macros(swApp);
            msg = new Message(swApp);

            // initialize app setting
            appset = new appsetting();

            // get edge string from setting
            edgeString = appset.EdgeNote;

            // split the string with ;
            string[] edgeStrArr = edgeString.Split(';');
            foreach(string edgeStr in edgeStrArr )
            {
                // add each string in array to combobox item
                EdgeNoteCb.Items.Add(edgeStr);
            }

            // select first item in the combobox
            EdgeNoteCb.SelectedIndex = 0;
        }

        private void FilenameDescBtn_Click(object sender, RoutedEventArgs e)
        {
            // get active model doc
            ModelDoc2 swModel = swApp.IActiveDoc2;
            
            // check
            if (!macros.checkDocNullandType(swModel, swDocumentTypes_e.swDocDRAWING)) return;

            // check if any view is selected
            SelectionMgr swSelMgr = swModel.ISelectionManager;

            // list of view
            List<View> viewList = new List<View>();
            
            // placeholder for output note
            List<Note> noteList = new List<Note>();

            // array of view
            View[] swViews = default(View[]);

            // get count of selected object
            int selObjCount = swSelMgr.GetSelectedObjectCount2(-1);

            // if more than 1 object selected
            if (selObjCount > 0)
            {
                // gather all the view and put in the view list
                try
                {
                    for (int i = 1; i < selObjCount + 1; i++)
                    {
                        // if selected object is view
                        if (swSelMgr.GetSelectedObjectType3(i, -1) == (int)swSelectType_e.swSelDRAWINGVIEWS)
                        {
                            // try to cast the selected object to solidworks view object
                            View swView = (View)swSelMgr.GetSelectedObject6(i, -1);

                            // add view to list
                            viewList.Add(swView);
                        }
                    }

                    // create array from list
                    swViews = viewList.ToArray();

                    // clear selection
                    swModel.ClearSelection2(true);
                }
                catch (Exception ex)
                {
                    msg.ErrorMsg(ex.ToString());
                }
            }
            else
            {
                msg.ErrorMsg("Please select at least a view to proceed.");
                return;
            }

            // process each view
            foreach (View swView in swViews)
            {
                // string for the name and description
                string filenameNDescription = "$PRPVIEW:\"SW-File Name(File Name)\"\r\n$PRPVIEW:\"Description\"";

                // select view
                bool boolstatus = swModel.Extension.SelectByID2(swView.Name, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
                
                // create note
                Note swNote = (Note)swModel.InsertNote(filenameNDescription);
                
                // set justification to center
                swNote.SetTextJustification((int)swTextJustification_e.swTextJustificationLeft);

                // set bold
                swNote.PropertyLinkedText = "<FONT style=B>" + filenameNDescription;

                // get annotation to set position
                Annotation swAnn = (Annotation)swNote.GetAnnotation();
                
                // get view outline
                double[] outline = (double[])swView.GetOutline();
                
                // set annotation position to top right of the view outline
                swAnn.SetPosition(outline[2], outline[3], 0);
            }
        }

        private void EdgeNoteBtn_Click(object sender, RoutedEventArgs e)
        {
            // get active model doc
            ModelDoc2 swModel = swApp.IActiveDoc2;

            // check
            if (!macros.checkDocNullandType(swModel, swDocumentTypes_e.swDocDRAWING)) return;

            // check if any view is selected
            SelectionMgr swSelMgr = swModel.ISelectionManager;

            // get count of selected object
            int selObjCount = swSelMgr.GetSelectedObjectCount2(-1);

            // lsit of entities
            List<Entity> entList = new List<Entity>();

            // list of silhouette edge
            List<SilhouetteEdge> seList = new List<SilhouetteEdge>();

            // if more than 1 object selected
            if (selObjCount > 0)
            {
                // add selected data to entity list if it is edge or vertices
                for(int i = 1; i < selObjCount+1; i++)
                {
                    // check selection type
                    int selObjType = swSelMgr.GetSelectedObjectType3(i, -1);
                    
                    // 1 is vertex, 3 is edge
                    if(selObjType == 1|| selObjType==3)
                    {
                        // cast to entity and save to list
                        Entity swEnt = (Entity)swSelMgr.GetSelectedObject6(i,-1); ;
                        entList.Add(swEnt);
                    }
                    else if (selObjType == 46)
                    {
                        SilhouetteEdge swSE = (SilhouetteEdge)swSelMgr.GetSelectedObject6(i, -1);
                        seList.Add(swSE);
                    }
                }

                // process each entity
                foreach(Entity ent in entList)
                {
                    // select the entity and insert note
                    ent.Select4(false, null);
                    swModel.InsertNote(EdgeNoteCb.Text);
                }

                // process each silhouette edge
                foreach (SilhouetteEdge se in seList)
                {
                    // select the edge and insert note
                    se.Select2(false, null);
                    swModel.InsertNote(EdgeNoteCb.Text);
                }

                // clear selection
                swModel.ClearSelection2(true);
            }
            else
            {
                msg.ErrorMsg("Select any edge or vertex before run this function.");
                return;
            }
        }

        private void PlatformNoteBtn_Click(object sender, RoutedEventArgs e)
        {
            // Common text
            string commonNoteStr = "-REFLECTIVE TAPE ALL AROUND\r\n" +
                                                    "-BEND TECH STICKER ON EACH SIDE\r\n" +
                                                    "-STD CERT PLATE AT EYE LEVEL";

            // get active model doc
            ModelDoc2 swModel = swApp.IActiveDoc2;

            // check
            if (!macros.checkDocNullandType(swModel, swDocumentTypes_e.swDocDRAWING)) return;

            // check if any view is selected
            SelectionMgr swSelMgr = swModel.ISelectionManager;

            // get count of selected object
            int selObjCount = swSelMgr.GetSelectedObjectCount2(-1);

            // lsit of entities
            List<Entity> entList = new List<Entity>();

            // list of silhouette edge
            List<SilhouetteEdge> seList = new List<SilhouetteEdge>();

            // if more than 1 object selected
            if (selObjCount > 0)
            {
                // add selected data to entity list if it is edge or vertices
                for (int i = 1; i < selObjCount + 1; i++)
                {
                    // check selection type
                    int selObjType = swSelMgr.GetSelectedObjectType3(i, -1);

                    // 1 is vertex, 3 is edge
                    if (selObjType == 1 || selObjType == 3)
                    {
                        // cast to entity and save to list
                        Entity swEnt = (Entity)swSelMgr.GetSelectedObject6(i, -1); ;
                        entList.Add(swEnt);
                    }
                    else if(selObjType == 46)
                    {
                        SilhouetteEdge swSE = (SilhouetteEdge)swSelMgr.GetSelectedObject6(i,-1);
                        seList.Add(swSE);
                    }
                }

                // process each entity
                foreach (Entity ent in entList)
                {
                    // select the entity and insert note
                    ent.Select4(false, null);
                    swModel.InsertNote(commonNoteStr);
                }

                // process each silhouette edge
                foreach(SilhouetteEdge se in seList)
                {
                    // select the edge and insert note
                    se.Select2(false, null);
                    swModel.InsertNote(commonNoteStr);
                }

                // clear selection
                swModel.ClearSelection2(true);
            }
            else
            {
                msg.ErrorMsg("Select any edge or vertex before run this function.");
                return;
            }
        }

        private void DescScaleBtn_Click(object sender, RoutedEventArgs e)
        {
            // get active model doc
            ModelDoc2 swModel = swApp.IActiveDoc2;

            // check
            if (!macros.checkDocNullandType(swModel, swDocumentTypes_e.swDocDRAWING)) return;

            // check if any view is selected
            SelectionMgr swSelMgr = swModel.ISelectionManager;

            // list of view
            List<View> viewList = new List<View>();

            // placeholder for output note
            List<Note> noteList = new List<Note>();

            // array of view
            View[] swViews = default(View[]);

            // get count of selected object
            int selObjCount = swSelMgr.GetSelectedObjectCount2(-1);

            // if more than 1 object selected
            if (selObjCount > 0)
            {
                // gather all the view and put in the view list
                try
                {
                    for (int i = 1; i < selObjCount + 1; i++)
                    {
                        // if selected object is view
                        if (swSelMgr.GetSelectedObjectType3(i, -1) == (int)swSelectType_e.swSelDRAWINGVIEWS)
                        {
                            // try to cast the selected object to solidworks view object
                            View swView = (View)swSelMgr.GetSelectedObject6(i, -1);

                            // add view to list
                            viewList.Add(swView);
                        }
                    }

                    // create array from list
                    swViews = viewList.ToArray();

                    // clear selection
                    swModel.ClearSelection2(true);
                }
                catch (Exception ex)
                {
                    msg.ErrorMsg(ex.ToString());
                }
            }
            else
            {
                msg.ErrorMsg("Please select at least a view to proceed.");
                return;
            }

            // process each view
            foreach (View swView in swViews)
            {
                // place holder for the description string
                string descScaleStr = "";

                // check if view is using sheet scale
                if(swView.UseSheetScale == 1)
                {
                    // string for the description
                    descScaleStr = "$PRPWLD:\"DESCRIPTION\"";
                }
                else
                {
                    // string for the description and scale
                    descScaleStr = "$PRPWLD:\"DESCRIPTION\"\r\nSCALE $PRP:\"SW-View Scale(View Scale)\"";
                }


                // select view
                bool boolstatus = swModel.Extension.SelectByID2(swView.Name, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);

                // create note
                Note swNote = (Note)swModel.InsertNote(descScaleStr);

                // set justification to center
                swNote.SetTextJustification((int)swTextJustification_e.swTextJustificationCenter);

                // get annotation to set position
                Annotation swAnn = (Annotation)swNote.GetAnnotation();

                // get view outline
                double[] outline = (double[])swView.GetOutline();

                // set annotation position to top right of the view outline
                swAnn.SetPosition((outline[0] + outline[2])/2, outline[1]-0.005, 0);
            }
        }
    }
}
