using System;

namespace VisioPlugin
{
    public class ShapeInfo
    {
        public string ShapeId { get; set; }
        public string ShapeType { get; set; }
        public string ShapeColor { get; set; }

        public override string ToString()
        {
            return $"{ShapeType} (ID: {ShapeId})" + (!string.IsNullOrEmpty(ShapeColor) ? $" - {ShapeColor}" : "");
        }
    }
} 