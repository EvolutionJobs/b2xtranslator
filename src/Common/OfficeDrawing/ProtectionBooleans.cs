using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing
{
    public class ProtectionBooleans
    {
        public bool fLockAgainstGrouping;
        public bool fLockAdjustHandles;
        public bool fLockText;
        public bool fLockVertices;
        public bool fLockCropping;
        public bool fLockAgainstSelect;
        public bool fLockPosition;
        public bool fLockAspectRatio;
        public bool fLockRotation;
        public bool fLockAgainstUngrouping;

        public bool fUsefLockAgainstGrouping;
        public bool fUsefLockAdjustHandles;
        public bool fUsefLockText;
        public bool fUsefLockVertices;
        public bool fUsefLockCropping;
        public bool fUsefLockAgainstSelect;
        public bool fUsefLockPosition;
        public bool fUsefLockAspectRatio;
        public bool fUsefLockRotation;
        public bool fUsefLockAgainstUngrouping;

        public ProtectionBooleans()
        {
        }

        public ProtectionBooleans(uint entryOperand)
        {
            fLockAgainstGrouping = Utils.BitmaskToBool(entryOperand, 0x1);
            fLockAdjustHandles = Utils.BitmaskToBool(entryOperand, 0x2);
            fLockText = Utils.BitmaskToBool(entryOperand, 0x4);
            fLockVertices = Utils.BitmaskToBool(entryOperand, 0x8);

            fLockCropping = Utils.BitmaskToBool(entryOperand, 0x10);
            fLockAgainstSelect = Utils.BitmaskToBool(entryOperand, 0x20);
            fLockPosition = Utils.BitmaskToBool(entryOperand, 0x30);
            fLockAspectRatio = Utils.BitmaskToBool(entryOperand, 0x40);

            fLockRotation = Utils.BitmaskToBool(entryOperand, 0x100);
            fLockAgainstUngrouping = Utils.BitmaskToBool(entryOperand, 0x200);

            //unused 0x400 0x800 0x1000 0x2000 0x4000 0x8000

            fUsefLockAgainstGrouping = Utils.BitmaskToBool(entryOperand, 0x10000);
            fUsefLockAdjustHandles = Utils.BitmaskToBool(entryOperand, 0x20000);
            fUsefLockText = Utils.BitmaskToBool(entryOperand, 0x40000);
            fUsefLockVertices = Utils.BitmaskToBool(entryOperand, 0x80000);

            fUsefLockCropping = Utils.BitmaskToBool(entryOperand, 0x100000);
            fUsefLockAgainstSelect = Utils.BitmaskToBool(entryOperand, 0x200000);
            fUsefLockPosition = Utils.BitmaskToBool(entryOperand, 0x400000);
            fUsefLockAspectRatio = Utils.BitmaskToBool(entryOperand, 0x800000);

            fUsefLockRotation = Utils.BitmaskToBool(entryOperand, 0x1000000);
            fUsefLockAgainstUngrouping = Utils.BitmaskToBool(entryOperand, 0x2000000);
        }
    }
}
