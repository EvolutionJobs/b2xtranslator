using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeTypeAttribute(166)]
    public class TextDeflateInflate : ShapeType
    {
        public TextDeflateInflate()
        {
            this.TextPath = true;
            this.Joins = JoinStyle.none;
            this.AdjustmentValues = "6054";
            this.Path = "m,l21600,m,10125c7200@1,14400@1,21600,10125m,11475c7200@2,14400@2,21600,11475m,21600r21600,e";
            this.ConnectorType = "rect";

            this.ExtrusionOk = true;
            this.Lock = new ProtectionBooleans();
            this.Lock.fUsefLockText = true;
            this.Lock.fLockText = true;
            this.LockShapeType = true;

            this.Formulas = new List<string>();
            this.Formulas.Add("prod #0 4 3");
            this.Formulas.Add("sum @0 0 4275");
            this.Formulas.Add("sum @0 0 2925");

            this.Handles = new List<Handle>();
            Handle h1 = new Handle();
            h1.position="center,#0";
            h1.yrange="1308,20292";
            this.Handles.Add(h1);

        }
    }
}
