using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeTypeAttribute(126)]
    class FlowChartSortType : ShapeType
    {
        public FlowChartSortType()
        {
            this.ShapeConcentricFill = true;
            this.Joins = JoinStyle.miter;
            this.Path = "m10800,l,10800,10800,21600,21600,10800xem,10800nfl21600,10800e";
            this.ConnectorLocations = "Rectangle";
            this.TextboxRectangle = "5400,5400,16200,16200";
        }
    }
}


