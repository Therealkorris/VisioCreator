using System;
using System.Collections.Generic;

namespace VisioPlugin
{
    public class ShapeInfo
    {
        public string ShapeId { get; set; }
        public string ShapeType { get; set; }
        public string ShapeColor { get; set; }
        public string Text { get; set; }
        public double PosX { get; set; }
        public double PosY { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public double Angle { get; set; }
        public int ZOrder { get; set; }
        public Dictionary<string, string> CustomProperties { get; set; }
        
        // Connection specific properties
        public bool IsConnector { get; set; }
        public string ConnectorType { get; set; }
        public string SourceShapeId { get; set; }
        public string TargetShapeId { get; set; }
        public double BeginX { get; set; }
        public double BeginY { get; set; }
        public double EndX { get; set; }
        public double EndY { get; set; }

        // Additional connector routing properties
        public List<Point> RoutingPoints { get; set; } = new List<Point>();
        public string ConnectorPattern { get; set; }  // Line pattern (straight, curved, etc)
        public double ConnectorWeight { get; set; }   // Line weight/thickness
        public string ConnectorRounding { get; set; } // Rounding for corners
        public List<ControlPoint> ControlPoints { get; set; } = new List<ControlPoint>();

        public ShapeInfo()
        {
            CustomProperties = new Dictionary<string, string>();
            RoutingPoints = new List<Point>();
            ControlPoints = new List<ControlPoint>();
        }

        public override string ToString()
        {
            if (IsConnector)
            {
                return $"{ShapeType} (ID: {ShapeId}) - Connects {SourceShapeId} to {TargetShapeId}";
            }
            return $"{ShapeType} (ID: {ShapeId})" + (!string.IsNullOrEmpty(ShapeColor) ? $" - {ShapeColor}" : "");
        }
    }

    public class Point
    {
        public double X { get; set; }
        public double Y { get; set; }

        public Point(double x, double y)
        {
            X = x;
            Y = y;
        }
    }

    public class ControlPoint
    {
        public double X { get; set; }
        public double Y { get; set; }
        public string Type { get; set; }  // "Bezier", "NURBS", etc.
        public double? Weight { get; set; }  // For NURBS control points

        public ControlPoint(double x, double y, string type, double? weight = null)
        {
            X = x;
            Y = y;
            Type = type;
            Weight = weight;
        }
    }
} 