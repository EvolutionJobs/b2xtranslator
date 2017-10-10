using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(9)]
    public class HexagonType : ShapeType
    {
        public HexagonType()
        {
            this.ShapeConcentricFill = true;

            this.Joins = JoinStyle.miter;

            this.Path = "m@0,l,10800@0,21600@1,21600,21600,10800@1,xe";

            this.Formulas = new List<string>();
            Formulas.Add("val #0");
            Formulas.Add("sum width 0 #0");
            Formulas.Add("sum height 0 #0");
            Formulas.Add("prod @0 2929 10000");
            Formulas.Add("sum width 0 @3");
            Formulas.Add("sum height 0 @3");

            this.AdjustmentValues = "5400";
            
            this.ConnectorLocations = "Rectangle";

            this.TextboxRectangle = "1800,1800,19800,19800;3600,3600,18000,18000;6300,6300,15300,15300";

            this.Handles = new List<Handle>();
            var HandleOne = new Handle
            {
                position = "#0,topLeft",
                xrange = "0,10800"
            };
            Handles.Add(HandleOne);

        }
    }
}
