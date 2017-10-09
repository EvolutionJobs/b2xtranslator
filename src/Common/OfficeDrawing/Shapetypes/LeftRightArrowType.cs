using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeTypeAttribute(69)]
    public class LeftRightArrowType :ShapeType
    {
        public LeftRightArrowType()
        {
            this.ShapeConcentricFill = false;

            this.Joins = JoinStyle.miter;

            this.Path = "m,10800l@0,21600@0@3@2@3@2,21600,21600,10800@2,0@2@1@0@1@0,xe";

            this.Formulas = new List<string>();
            this.Formulas.Add("val #0");
            this.Formulas.Add("val #1");
            this.Formulas.Add("sum 21600 0 #0");
            this.Formulas.Add("sum 21600 0 #1");
            this.Formulas.Add("prod #0 #1 10800");
            this.Formulas.Add("sum #0 0 @4");
            this.Formulas.Add("sum 21600 0 @5");

            this.AdjustmentValues="4320,5400";

            this.ConnectorLocations="@2,0;10800,@1;@0,0;0,10800;@0,21600;10800,@3;@2,21600;21600,10800";

            this.ConnectorAngles="270,270,270,180,90,90,90,0";

            this.TextboxRectangle="@5,@1,@6,@3";

            this.Handles = new List<Handle>();
            Handle HandleOne = new Handle();
            HandleOne.position="#0,#1";
            HandleOne.xrange="0,10800";
            HandleOne.yrange="0,10800";
            this.Handles.Add(HandleOne);

        }
    }
}
