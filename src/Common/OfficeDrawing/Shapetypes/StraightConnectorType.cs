using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(32)]
    public class StraightConnectorType : ShapeType
    {
        public StraightConnectorType()
        {
            this.Path = "m,l21600,21600e";
            this.ConnectorLocations = "0,0;21600,21600";
        }
    }
}
