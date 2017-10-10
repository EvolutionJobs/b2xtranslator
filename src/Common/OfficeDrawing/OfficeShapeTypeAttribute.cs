using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing
{
    /// <summary>
    /// Used for mapping Office shape types to the classes implementing them.
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    public class OfficeShapeTypeAttribute : Attribute
    {
        public OfficeShapeTypeAttribute() { }

        public OfficeShapeTypeAttribute(uint typecode)
        {
            this.TypeCode = typecode;
        }

        public uint TypeCode;
    }
}
