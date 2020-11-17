using System;
using SldWorks;
using SwConst;

namespace PlaneFacesArea
{
    class Program
    {
        static void Main(string[] args)
        {
            SldWorks.SldWorks swApp;
            ModelDoc2 swModelDoc;
            AssemblyDoc swAssemblyDoc;
            Object[] components;
            double totalAreaPlanarFaces = 0;

            swApp = new SldWorks.SldWorks();
            swApp.Visible = true;
            swModelDoc = (ModelDoc2)swApp.ActiveDoc;
            if (swModelDoc == null)
            {
                Console.WriteLine("Veuillez ouvrir un assemblage.");
                Console.ReadKey();
                return;
            }
            else if (swModelDoc.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                Console.WriteLine("Veuillez ouvrir un assemblage.");
                Console.ReadKey();
                return;
            }
            swAssemblyDoc = (AssemblyDoc)swModelDoc;
            components = (object[])swAssemblyDoc.GetComponents(false);
            foreach (var item in components)
            {
                Component2 swComponent;
                swComponent = (Component2)item;
                Console.WriteLine(swComponent.Name);
                if (swComponent.IsSuppressed() == true)
                {
                    continue;
                }
                else
                {
                    ModelDoc2 swComponentDoc;
                    swComponentDoc = (ModelDoc2)swComponent.GetModelDoc2(); // Connexion au document (pièce ou assemblage) lié au composant.
                    if (swComponentDoc == null)
                    {
                        // Le ModelDoc2 sera Nothing si le fichier lié au composant n'est pas ouvert en mémoire(lightweight, supprimé, fichier non trouvé, etc.).
                        continue;
                    }
                    // On doit s'assurer que ce composant est lié à une pièce et non à un assemblage. Sinon, le cast vers le type PartDoc créera un crash(une exception).
                    if (swComponentDoc.GetType() == (int)swDocumentTypes_e.swDocPART)
                    {
                        PartDoc swCompPartDoc;
                        double PlanarFacesAreaForPart = 0;
                        swCompPartDoc = (PartDoc)swComponentDoc;
                        PlanarFacesAreaForPart = GetAreaOfPlanarFacesOfPart(swCompPartDoc);
                        totalAreaPlanarFaces += PlanarFacesAreaForPart;
                        Console.WriteLine("Aire totale temporaire = {0}", totalAreaPlanarFaces);
                    }
                }
            }
            Console.WriteLine("L'aire de toutes les faces planes vaut: {0}", totalAreaPlanarFaces);
            Console.ReadKey();
        }
        static double GetAreaOfPlanarFacesOfPart(PartDoc _swPartDoc)
        {
            double area = 0;
            Surface swSurface;
            Face2 swFace;
            Body2 swBody;
            Object[] bodies;
            object[] faces;

            bodies = (object[])_swPartDoc.GetBodies2((int)swBodyType_e.swSolidBody, true);
            for (int i = 0; i < bodies.Length; i++)
            {
                Object objBody;
                objBody = bodies[i];
                swBody = (Body2)objBody;
                faces = (object[])swBody.GetFaces();
                for (int j = 0; j < faces.Length; j++)
                {
                    Object objFace;
                    objFace = faces[j];
                    swFace = (Face2)objFace;
                    swSurface = (Surface)swFace.GetSurface();
                    if (swSurface.IsPlane())
                    {
                        // obtenir l'aire et l'additionner au reste
                        area += swFace.GetArea();
                        Console.WriteLine("Aire de la face = {0}", area);
                    }
                }
            }
            return area;
        }
    }
}
