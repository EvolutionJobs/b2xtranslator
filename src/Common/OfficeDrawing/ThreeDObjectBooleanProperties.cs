using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing
{
    public class ThreeDObjectProperties
    {
        public bool fc3DLightFace;
        public bool fc3UseExtrusionColor;
        public bool fc3DMetallic;
        public bool fc3D;

        public bool fUsefc3DLightFace;
        public bool fUsefc3DUseExtrusionColor;
        public bool fUsefc3DMetallic;
        public bool fUsefc3D;

        public ThreeDObjectProperties(uint entryOperand)
        {

            fc3DLightFace = Utils.BitmaskToBool(entryOperand, 0x1 << 0);
            fc3UseExtrusionColor = Utils.BitmaskToBool(entryOperand, 0x1 << 1);
            fc3DMetallic = Utils.BitmaskToBool(entryOperand, 0x1 << 2);
            fc3D= Utils.BitmaskToBool(entryOperand, 0x1 << 3);
            //12 unused
            fUsefc3DLightFace = Utils.BitmaskToBool(entryOperand, 0x1 << 16);
            fUsefc3DUseExtrusionColor = Utils.BitmaskToBool(entryOperand, 0x1 << 17);
            fUsefc3DMetallic = Utils.BitmaskToBool(entryOperand, 0x1 << 18);
            fUsefc3D = Utils.BitmaskToBool(entryOperand, 0x1 << 19);
            //12 unused
        }
    }
}
