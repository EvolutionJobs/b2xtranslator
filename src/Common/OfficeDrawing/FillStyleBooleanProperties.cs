using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing
{
    public class FillStyleBooleanProperties
    {
        public bool fNoFillHitTest;
        public bool fillUseRect;
        public bool fillShape;
        public bool fHitTestFill;
        public bool fFilled;
        public bool fUseShapeAnchor;
        public bool fRecolorFillAsPicture;
        public bool fUsefNoFillHitTest;
        public bool fUsefillUseRect;
        public bool fUsefillShape;
        public bool fUseHitTestFill;
        public bool fUsefFilled;
        public bool fUsefUseShapeAnchor;
        public bool fUsefRecolorFillAsPicture;

        public FillStyleBooleanProperties(uint entryOperand)
        {
            fNoFillHitTest = Utils.BitmaskToBool(entryOperand, 0x1);
            fillUseRect = Utils.BitmaskToBool(entryOperand, 0x1 << 1);
            fillShape = Utils.BitmaskToBool(entryOperand, 0x1 << 2);
            fHitTestFill = Utils.BitmaskToBool(entryOperand, 0x1 << 3);
            fFilled = Utils.BitmaskToBool(entryOperand, 0x1 << 4);
            fUseShapeAnchor = Utils.BitmaskToBool(entryOperand, 0x1 << 5);
            fRecolorFillAsPicture = Utils.BitmaskToBool(entryOperand, 0x1 << 6);
            // 0x1 << 7-15 is ununsed
            fUsefNoFillHitTest = Utils.BitmaskToBool(entryOperand, 0x1 << 16);
            fUsefillUseRect = Utils.BitmaskToBool(entryOperand, 0x1 << 17);
            fUsefillShape = Utils.BitmaskToBool(entryOperand, 0x1 << 18);
            fUseHitTestFill = Utils.BitmaskToBool(entryOperand, 0x1 << 19);
            fUsefFilled = Utils.BitmaskToBool(entryOperand, 0x1 << 20);
            fUsefUseShapeAnchor = Utils.BitmaskToBool(entryOperand, 0x1 << 21);
            fUsefRecolorFillAsPicture = Utils.BitmaskToBool(entryOperand, 0x1 << 22);

        }
    }
}
