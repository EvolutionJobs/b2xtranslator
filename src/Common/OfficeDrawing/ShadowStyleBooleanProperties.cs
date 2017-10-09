using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing
{
    public class ShadowStyleBooleanProperties
    {
        public bool fShadowObscured;
        public bool fShadow;
        public bool fUsefshadowObscured;
        public bool fUsefShadow;

        public ShadowStyleBooleanProperties(UInt32 entryOperand)
        {
            fShadowObscured = Utils.BitmaskToBool(entryOperand, 0x1);
            fShadow = Utils.BitmaskToBool(entryOperand, 0x2);
            fUsefshadowObscured = Utils.BitmaskToBool(entryOperand, 0x10000);
            fUsefShadow = Utils.BitmaskToBool(entryOperand, 0x20000);
        }
    }
}
