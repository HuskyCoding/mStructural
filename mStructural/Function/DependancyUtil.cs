using SolidWorks.Interop.sldworks;
using System.Collections;
using System.Collections.Generic;

namespace mStructural.Function
{
    public class DependancyUtil
    {

         public string GetTopModelStrValue(SldWorks swApp, string drawingPath)
         {
             List<string> modelListInDraw = new List<string>();
             string topModelPathName = "";

             Hashtable drawDep = getDependancy(swApp, drawingPath);
             if (drawDep.Count > 0)
             {
                 foreach (DictionaryEntry entry in drawDep)
                 {
                     modelListInDraw.Add(entry.Value.ToString());
                 }
             }

             string[] modelArrInDraw = modelListInDraw.ToArray();

             Hashtable allDep = getDependancies(swApp, modelArrInDraw);

             foreach (string modelInDraw in modelArrInDraw)
             {
                 if (!allDep.ContainsKey(modelInDraw))
                 {
                     topModelPathName = modelInDraw;
                 }
             }

             return topModelPathName;
         }
        
        public string GetTopModelStrKey(SldWorks swApp, string drawingPath)
        {
            List<string> modelListInDraw = new List<string>();
            string topModelPathName = "";

            Hashtable drawDep = getDependancy(swApp, drawingPath);
            if (drawDep.Count > 0)
            {
                foreach (DictionaryEntry entry in drawDep)
                {
                    modelListInDraw.Add(entry.Key.ToString());
                }
            }

            string[] modelArrInDraw = modelListInDraw.ToArray();

            Hashtable allDep = getDependancies(swApp, modelArrInDraw);

            foreach (string modelInDraw in modelArrInDraw)
            {
                if (!allDep.ContainsKey(modelInDraw))
                {
                    topModelPathName = modelInDraw;
                }
            }

            return topModelPathName;
        }

        public Hashtable getDependancies(SldWorks swApp, string[] allFiles)
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

        public Hashtable getDependancy(SldWorks swApp, string file)
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
    }
}
