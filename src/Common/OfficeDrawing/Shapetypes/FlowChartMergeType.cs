using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeTypeAttribute(128)]
    class FlowChartMergeType : ShapeType
    {
        public FlowChartMergeType()
        {
            this.ShapeConcentricFill = true;
            this.Joins = JoinStyle.miter;
            this.Path = "m,l21600,,10800,21600xe";
            this.ConnectorLocations = "10800,0;5400,10800;10800,21600;16200,10800";
            this.TextboxRectangle = "5400,0,16200,10800";
        }
    }
}


