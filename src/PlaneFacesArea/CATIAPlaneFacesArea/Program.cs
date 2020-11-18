using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HybridShapeTypeLib;
using INFITF;
using MECMOD;
using ProductStructureTypeLib;
using SPATypeLib;

namespace CATIAPlaneFacesArea
{
    class Program
    {
        static void Main(string[] args)
        {

            var catiaType = Type.GetTypeFromProgID("CATIA.Application");
            var catiaInstane = (Application)Activator.CreateInstance(catiaType);
            var activeDocument = catiaInstane.ActiveDocument;
            if (activeDocument == null)
            {
                Console.WriteLine("Veuillez ouvrir un assemblage.");
                Console.ReadKey();
                return;
            }
            else if (!(activeDocument is ProductDocument))
            {
                Console.WriteLine("Veuillez ouvrir un assemblage.");
                Console.ReadKey();
                return;
            }

            var productDocument = (ProductDocument)activeDocument;
            var selection = productDocument.Selection;

            var rootProduct = productDocument.Product;
            var products = rootProduct.Products;
            var workbenchId = catiaInstane.GetWorkbenchId();
            var areas = 0.0;
            double PlanarFacesAreaForPart = 0.0;
            double totalAreaPlanarFaces = 0.0;
            for (int i = 1; i <= products.Count; i++)
            {

                var product = products.Item(i);
                var instanceName = product.get_Name();
                var masterShape = product.GetMasterShapeRepresentation(true);
                var partDocumentOf = masterShape as PartDocument;


                var documentName = product.GetMasterShapeRepresentationPathName();
                Console.WriteLine($"Processing document {documentName}");

                selection.Clear();
                selection.Add(partDocumentOf.Part);


                catiaInstane.StartWorkbench("PrtCfg");
                workbenchId = catiaInstane.GetWorkbenchId();

                var part = partDocumentOf.Part;
                var partSelection = partDocumentOf.Selection;
                partSelection.Clear();
                partSelection.Add(part.MainBody);


                var hybridBody = part.HybridBodies.Add();
                var hsf = (HybridShapeFactory)part.HybridShapeFactory;

                selection.Search("Topology.CGMFace,sel");
                var extracts = new List<HybridShapeExtract>();
                for (var selectedIndex = 1; selectedIndex <= selection.Count2; selectedIndex++)
                {
                    var selectedItem = selection.Item2(selectedIndex).Reference;
                    var extract = hsf.AddNewExtract(selectedItem);
                    extracts.Add(extract);
                    hybridBody.AppendHybridShape(extract);
                }
                part.Update();
                catiaInstane.StartWorkbench("Inertia");
                var spaWb = (SPAWorkbench)partDocumentOf.GetWorkbench("SPAWorkbench");

                foreach (var extract in extracts)
                {
                    try
                    {

                        var measurableOnExtract = spaWb.GetMeasurable(part.CreateReferenceFromObject(extract));
                        Console.WriteLine(measurableOnExtract.GeometryName);
                        var planecomponents = new object[9]; //new object[9];//Enumerable.Range(0,9).Select(x=> 0.0).ToArray();
                        areas += measurableOnExtract.Area;
                        Console.WriteLine("Aire de la face = {0}", measurableOnExtract.Area);
                        if (measurableOnExtract.GeometryName == CatMeasurableName.CatMeasurablePlane)
                        {
                            //measurableOnExtract.GetPlane(planecomponents);
                        }
                    }
                    catch (Exception e)
                    {
                    }
                }
                //Console.WriteLine($"The computed area is {areas}");
                selection.Clear();
                selection.Add(rootProduct);

                catiaInstane.StartWorkbench("Assembly");

                workbenchId = catiaInstane.GetWorkbenchId();
                PlanarFacesAreaForPart = areas;
                totalAreaPlanarFaces += PlanarFacesAreaForPart;

                selection.Clear();
                selection.Add(hybridBody);
                selection.Delete();
                Console.WriteLine("Aire totale temporaire = {0}", totalAreaPlanarFaces);
            }


            Console.WriteLine("L'aire de toutes les faces planes vaut: {0}", areas);
            Console.ReadKey();
        }
    }
}
