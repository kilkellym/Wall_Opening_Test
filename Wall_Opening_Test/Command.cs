#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Windows.Controls;

#endregion

namespace Wall_Opening_Test
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            // select wall
            Reference wallRef = uidoc.Selection.PickObject(ObjectType.Element, "Select Wall");
            Wall wall = doc.GetElement(wallRef) as Wall;

            // get intersecting generic families
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(FamilyInstance));
            collector.WherePasses(new ElementIntersectsElementFilter(wall));

            using(Transaction t = new Transaction(doc))
            {
                t.Start("Insert wall openings");

                foreach (Element elem in collector.ToList())
                    InsertWallOpening(doc, elem, wall);

                t.Commit();
            }
            
            return Result.Succeeded;
        }

        private static Solid GetWallSolid(Wall wall)
        {
            // get wall solid
            Options opts = new Options();
            GeometryElement wallGeom = wall.get_Geometry(opts);
            Solid wallSolid = wallGeom.First() as Solid;

            return wallSolid;
        }

        internal void InsertWallOpening(Document doc, Element curElem, Wall wall)
        {
            Solid wallSolid = GetWallSolid(wall);
            IList<Reference> wallFaces = HostObjectUtils.GetSideFaces(wall, ShellLayerType.Interior);
            Face wallFace = doc.GetElement(wallFaces[0]).GetGeometryObjectFromReference(wallFaces[0]) as Face;
            
            // get family solid
            Solid famSolid = GetFamilySolid(curElem);

            // get intersect solid between wall and family instance
            Solid intersect = BooleanOperationsUtils.ExecuteBooleanOperation(famSolid, wallSolid, BooleanOperationsType.Intersect);

            // get face and edge array from intersect solid
            Face faceInt = null;
            foreach (Face curFace in intersect.Faces)
            {
                // get non-intersecting face - this is the outside face of the intersecting solid
                if (wallFace.Intersect(curFace) == FaceIntersectionFaceResult.NonIntersecting)
                    faceInt = curFace;
            }

            EdgeArray edgeArray = faceInt.EdgeLoops.get_Item(0);

            List<Curve> curves = new List<Curve>();
            foreach (Edge e in edgeArray)
            {
                curves.Add(e.AsCurve());
            }

            List<XYZ> pointList = GetPointsFromCurves(curves);
            XYZ minPoint = GetMinPoint(pointList);
            XYZ maxPoint = GetMaxPoint(pointList);

            // create opening and delete generic family
            Opening opening = doc.Create.NewOpening(wall, minPoint, maxPoint);
            doc.Delete(curElem.Id);
        }

        private static Solid GetFamilySolid(Element curElem)
        {
            // get generic family geometry
            Options opts = new Options();
            GeometryElement objGeom = curElem.get_Geometry(opts);
            GeometryInstance geomInst = objGeom.First() as GeometryInstance;
            GeometryElement geomElem = geomInst.GetInstanceGeometry();

            Solid solid = null;
            IEnumerator<GeometryObject> enumer = geomElem.GetEnumerator();
            while (enumer.MoveNext())
            {
                Solid curSolid = enumer.Current as Solid;
                if (curSolid.SurfaceArea > 0)
                {
                    solid = curSolid;
                }
            }

            return solid;
        }

        public static XYZ GetMinPoint(List<XYZ> pointList)
        {
            //get minimum point from a list of points
            double minX = 0;double minY = 0;double minZ = 0;

            for(int i=0; i<pointList.Count; i++)
            {
                if(i==0)
                {
                    minX = pointList[i].X; minY= pointList[i].Y; minZ = pointList[i].Z;
                }
                else 
                {
                    if (pointList[i].X < minX)
                    {
                        minX = pointList[i].X;
                    }
                    if (pointList[i].Y < minY)
                    {
                        minY = pointList[i].Y;
                    }
                    if (pointList[i].Z < minZ)
                    {
                        minZ = pointList[i].Z;
                    }
                }
                
            }
            XYZ minPoint = new XYZ(minX, minY, minZ);
            return minPoint;
        }

        internal List<XYZ> GetPointsFromCurves(List<Curve> curves)
        {
            List<XYZ> returnList = new List<XYZ>();

            foreach(Curve curve in curves)
            {
                returnList.Add(curve.GetEndPoint(0));
            }

            return returnList;
        }

        public static XYZ GetMaxPoint(List<XYZ> pointList)
        {
            //get maximum point from a list of points
            double maxX = 0; double maxY = 0; double maxZ = 0;

            for (int i = 0; i < pointList.Count; i++)
            {
                if (i == 0)
                {
                    maxX = pointList[i].X; maxY = pointList[i].Y; maxZ = pointList[i].Z;
                }
                else
                {
                    if (pointList[i].X > maxX)
                    {
                        maxX = pointList[i].X;
                    }
                    if (pointList[i].Y > maxY)
                    {
                        maxY = pointList[i].Y;
                    }
                    if (pointList[i].Z > maxZ)
                    {
                        maxZ = pointList[i].Z;
                    }
                }

            }
            XYZ maxPoint = new XYZ(maxX, maxY, maxZ);
            return maxPoint;
        }

        public static String GetMethod()
        {
            var method = MethodBase.GetCurrentMethod().DeclaringType?.FullName;
            return method;
        }
    }
}
