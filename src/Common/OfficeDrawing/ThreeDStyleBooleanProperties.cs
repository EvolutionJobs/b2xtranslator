using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing
{
    public class ThreeDStyleProperties
    {
        public bool fc3DFillHarsh;
        public bool fc3DKeyHarsh;
        public bool fc3DParallel;
        public bool fc3DRotationCenterAuto;
        public bool fc3DConstrainRotation;

        public bool fUsefc3DFillHarsh;
        public bool fUsefc3DKeyHarsh;
        public bool fUsefc3DParallel;
        public bool fUsefc3DRotationCenterAuto;
        public bool fUsefc3DConstrainRotation;

        public ThreeDStyleProperties(UInt32 entryOperand)
        {

            fc3DFillHarsh = Utils.BitmaskToBool(entryOperand, 0x1 << 0);
            fc3DKeyHarsh = Utils.BitmaskToBool(entryOperand, 0x1 << 1);
            fc3DParallel = Utils.BitmaskToBool(entryOperand, 0x1 << 2);
            fc3DRotationCenterAuto = Utils.BitmaskToBool(entryOperand, 0x1 << 3);
            fc3DConstrainRotation = Utils.BitmaskToBool(entryOperand, 0x1 << 4);
            //11 unused
            fUsefc3DFillHarsh = Utils.BitmaskToBool(entryOperand, 0x1 << 16);
            fUsefc3DKeyHarsh = Utils.BitmaskToBool(entryOperand, 0x1 << 17);
            fUsefc3DParallel = Utils.BitmaskToBool(entryOperand, 0x1 << 18);
            fUsefc3DRotationCenterAuto = Utils.BitmaskToBool(entryOperand, 0x1 << 19);
            fUsefc3DConstrainRotation = Utils.BitmaskToBool(entryOperand, 0x1 << 20);
            //11 unused
        }
    }
}
