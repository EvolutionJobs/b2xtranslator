using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing.Shapetypes
{
    /// <summary>
    /// interim solution<br/>
    /// OOX uses an additional attribute: arcsize
    /// </summary>
    [OfficeShapeTypeAttribute(2)]
    public class RoundedRectangleType : RectangleType
    {
        public RoundedRectangleType() : base()
        {
            this.Joins = JoinStyle.round;
            this.AdjustmentValues = "5400";
        }
    }
}
