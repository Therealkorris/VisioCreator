using System;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;

namespace VisioPlugin
{
    public class ShapeInfo
    {
        public string ShapeId { get; set; }
        public string ShapeType { get; set; }
        public string ShapeColor { get; set; }
        public string Text { get; set; }
        public string Category { get; set; }
        private double _posX;
        private double _posY;
        private double _width;
        private double _height;
        public double Angle { get; set; }
        public int ZOrder { get; set; }
        public Dictionary<string, string> CustomProperties { get; set; }
        
        // Store page dimensions for percentage calculations
        private static double PageWidth { get; set; }
        private static double PageHeight { get; set; }

        // Position and size in inches (internal Visio units)
        public double PosX 
        { 
            get => _posX;
            set 
            {
                _posX = value;
                PosXPercent = PageWidth > 0 ? (value / PageWidth) * 100 : value;
            }
        }
        public double PosY 
        { 
            get => _posY;
            set 
            {
                _posY = value;
                PosYPercent = PageHeight > 0 ? (value / PageHeight) * 100 : value;
            }
        }
        public double Width 
        { 
            get => _width;
            set 
            {
                _width = value;
                WidthPercent = PageWidth > 0 ? (value / PageWidth) * 100 : value;
            }
        }
        public double Height 
        { 
            get => _height;
            set 
            {
                _height = value;
                HeightPercent = PageHeight > 0 ? (value / PageHeight) * 100 : value;
            }
        }

        // Position and size in percentages (for API)
        public double PosXPercent { get; private set; }
        public double PosYPercent { get; private set; }
        public double WidthPercent { get; private set; }
        public double HeightPercent { get; private set; }
        
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

        // Method to set page dimensions (call this when getting page info)
        public static void SetPageDimensions(double width, double height)
        {
            PageWidth = width;
            PageHeight = height;
        }

        public override string ToString()
        {
            if (IsConnector)
            {
                return $"{ShapeType} (ID: {ShapeId}) - Connects {SourceShapeId} to {TargetShapeId}";
            }
            return $"{ShapeType} (ID: {ShapeId})" + (!string.IsNullOrEmpty(ShapeColor) ? $" - {ShapeColor}" : "");
        }

        // Method to get position and size as percentages for API
        public JObject ToApiFormat()
        {
            return new JObject
            {
                ["category"] = Category,
                ["shape_id"] = ShapeId,
                ["shape_type"] = ShapeType,
                ["pos_x"] = PosXPercent,
                ["pos_y"] = PosYPercent,
                ["width"] = WidthPercent,
                ["height"] = HeightPercent,
                ["shape_color"] = ShapeColor,
                ["text"] = Text,
                ["angle"] = Angle,
                ["z_order"] = ZOrder,
                ["is_connector"] = IsConnector,
                ["connector_type"] = ConnectorType,
                ["source_shape_id"] = SourceShapeId,
                ["target_shape_id"] = TargetShapeId,
                ["begin_x"] = BeginX,
                ["begin_y"] = BeginY,
                ["end_x"] = EndX,
                ["end_y"] = EndY,
                ["routing_points"] = JToken.FromObject(RoutingPoints),
                ["connector_pattern"] = ConnectorPattern,
                ["connector_weight"] = ConnectorWeight,
                ["connector_rounding"] = ConnectorRounding,
                ["control_points"] = JToken.FromObject(ControlPoints),
                ["custom_properties"] = JToken.FromObject(CustomProperties)
            };
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