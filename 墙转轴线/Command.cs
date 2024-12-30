using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;

namespace 墙转轴线
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand

    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //ui应用程序
            UIApplication uiapp = commandData.Application;
            //应用程序
            Application app = uiapp.Application;
            //ui文档
            UIDocument uidoc = uiapp.ActiveUIDocument;
            //文档
            Document doc = uidoc.Document;
            WallFilter wallFilter = new WallFilter();
            List<Reference> walls = uidoc.Selection.PickObjects(ObjectType.Element, wallFilter).ToList();
            List<Curve> curves = new List<Curve>();
            walls.ForEach(i =>
            {
                LocationCurve location = doc.GetElement(i).Location as LocationCurve;
                curves.Add(location.Curve);
            });

            using (Transaction trans = new Transaction(doc, "transaction"))
            {
                trans.Start();
                curves.ForEach(i =>
                {
                    double distance = 2000;
                    if (i as Line != null)
                    {
                        distance = UnitUtils.ConvertToInternalUnits(distance, DisplayUnitType.DUT_MILLIMETERS);
                        Line line = i as Line;
                        XYZ vec = line.Direction.Normalize() * distance;
                        XYZ point0 = line.GetEndPoint(0) - vec;
                        XYZ point1 = line.GetEndPoint(1) + vec;
                        line = Line.CreateBound(point0, point1);
                        Grid.Create(doc, line);
                    }
                    if (i as Arc != null)
                    {
                        distance = UnitUtils.ConvertToInternalUnits(distance, DisplayUnitType.DUT_MILLIMETERS);
                        Arc arc = i as Arc;
                        GetStartAndEndPoint(arc, out XYZ StartPoint, out XYZ EndPoint);

                        SetStartAndEndAngle(StartPoint, EndPoint, arc, out double StartAngle, out double EndAngle);
                        arc = Arc.Create(arc.Center, arc.Radius, StartAngle - distance / arc.Radius, EndAngle + distance / arc.Radius, XYZ.BasisX, XYZ.BasisY);
                        Grid.Create(doc, arc);
                    }
                });
                trans.Commit();
            }

            return Result.Succeeded;
        }

        private void SetStartAndEndAngle(XYZ StartPoint, XYZ EndPoint, Arc GridArc, out double StartAngle, out double EndAngle)
        {
            double StartToEndAngle = GetAngleBetweenVectors(StartPoint - GridArc.Center, EndPoint - GridArc.Center);
            Transform t = GridArc.ComputeDerivatives(0.5, true);

            double dis = GetDistanceFromPointToCurve(t.Origin, Line.CreateBound(StartPoint, EndPoint));
            if (dis > GridArc.Radius)
            {
                StartToEndAngle = 2 * Math.PI - StartToEndAngle;
            }
            var vector = StartPoint - GridArc.Center;
            if (vector.Y > 0)
            {
                StartAngle = GetAngleBetweenVectors(vector, XYZ.BasisX);
            }
            else
            {
                StartAngle = -GetAngleBetweenVectors(vector, XYZ.BasisX);
            }

            EndAngle = StartAngle + StartToEndAngle;
        }

        /// <summary>
        /// 计算两个向量的夹角
        /// </summary>
        /// <param name="vector1"></param>
        /// <param name="vector2"></param>
        /// <returns></returns>
        public double GetAngleBetweenVectors(XYZ vector1, XYZ vector2)
        { // 计算点积
            double dotProduct = vector1.DotProduct(vector2);
            //计算两个向量的长度
            double magnitude1 = vector1.GetLength();
            double magnitude2 = vector2.GetLength();
            //计算夹角的余弦值
            double cosTheta = dotProduct / (magnitude1 * magnitude2);
            //确保余弦值在有效范围内
            cosTheta = Math.Max(-1.0, Math.Min(1.0, cosTheta));
            //计算夹角的弧度
            double thetaInRadians = Math.Acos(cosTheta);
            return thetaInRadians;
        }

        /// <summary>
        /// 点到曲线的距离
        /// </summary>
        /// <param name="point"></param>
        /// <param name="line"></param>
        /// <returns></returns>
        public double GetDistanceFromPointToCurve(XYZ point, Curve curve)
        { // 获取曲线在点附近的最近点
            IntersectionResult result = curve.Project(point);
            if (result != null)
            {
                XYZ nearestPoint = result.XYZPoint;
                // 计算点到最近点的向量长度
                double distance = point.DistanceTo(nearestPoint);
                return distance;
            }
            else
            {
                // 处理曲线和点不相交的情况
                return 0;
            }
        }

        public void GetStartAndEndPoint(Arc arc, out XYZ StartPoint, out XYZ EndPoint)
        {
            XYZ Point0 = arc.GetEndPoint(0);
            XYZ Point1 = arc.GetEndPoint(1);

            XYZ vector0 = (Point0 - arc.Center);
            XYZ vector1 = (Point1 - arc.Center);

            XYZ vectorZ = vector0.CrossProduct(vector1);
            if (vectorZ.Z > 0)
            {
                StartPoint = Point0;
                EndPoint = Point1;
            }
            else
            {
                StartPoint = Point1;
                EndPoint = Point0;
            }
        }
    }

    internal class WallFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            if (elem.Category?.Name == "墙")
            {
                return true;
            }
            return false;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }
}