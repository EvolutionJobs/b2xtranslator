using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing
{
    public class GroupShapeBooleans
    {
        public bool fPrint;
        public bool fHidden;
        public bool fOneD;
        public bool fIsButton;

        public bool fOnDblClickNotify;
        public bool fBehindDocument;
        public bool fEditedWrap;
        public bool fScriptAnchor;

        public bool fReallyHidden;
        public bool fAllowOverlap;
        public bool fUserDrawn;
        public bool fHorizRule;

        public bool fNoshadeHR;
        public bool fStandardHR;
        public bool fIsBullet;
        public bool fLayoutInCell;

        public bool fUsefPrint;
        public bool fUsefHidden;
        public bool fUsefOneD;
        public bool fUsefIsButton;

        public bool fUsefOnDblClickNotify;
        public bool fUsefBehindDocument;
        public bool fUsefEditedWrap;
        public bool fUsefScriptAnchor;

        public bool fUsefReallyHidden;
        public bool fUsefAllowOverlap;
        public bool fUsefUserDrawn;
        public bool fUsefHorizRule;

        public bool fUsefNoshadeHR;
        public bool fUsefStandardHR;
        public bool fUsefIsBullet;
        public bool fUsefLayoutInCell;

        public GroupShapeBooleans(uint entryOperand)
        {
            fPrint = Utils.BitmaskToBool(entryOperand, 0x1);
            fHidden = Utils.BitmaskToBool(entryOperand, 0x2);
            fOneD = Utils.BitmaskToBool(entryOperand, 0x4);
            fIsButton = Utils.BitmaskToBool(entryOperand, 0x8);

            fOnDblClickNotify = Utils.BitmaskToBool(entryOperand, 0x10);
            fBehindDocument = Utils.BitmaskToBool(entryOperand, 0x20);
            fEditedWrap = Utils.BitmaskToBool(entryOperand, 0x40);
            fScriptAnchor = Utils.BitmaskToBool(entryOperand, 0x80);

            fReallyHidden = Utils.BitmaskToBool(entryOperand, 0x100);
            fAllowOverlap = Utils.BitmaskToBool(entryOperand, 0x200);
            fUserDrawn = Utils.BitmaskToBool(entryOperand, 0x400);
            fHorizRule = Utils.BitmaskToBool(entryOperand, 0x800);

            fNoshadeHR = Utils.BitmaskToBool(entryOperand, 0x1000);
            fStandardHR = Utils.BitmaskToBool(entryOperand, 0x2000);
            fIsBullet = Utils.BitmaskToBool(entryOperand, 0x4000);
            fLayoutInCell = Utils.BitmaskToBool(entryOperand, 0x8000);

            fUsefPrint = Utils.BitmaskToBool(entryOperand, 0x10000);
            fUsefHidden = Utils.BitmaskToBool(entryOperand, 0x20000);
            fUsefOneD = Utils.BitmaskToBool(entryOperand, 0x40000);
            fUsefIsButton = Utils.BitmaskToBool(entryOperand, 0x80000);

            fUsefOnDblClickNotify = Utils.BitmaskToBool(entryOperand, 0x100000);
            fUsefBehindDocument = Utils.BitmaskToBool(entryOperand, 0x200000);
            fUsefEditedWrap = Utils.BitmaskToBool(entryOperand, 0x400000);
            fUsefScriptAnchor = Utils.BitmaskToBool(entryOperand, 0x800000);

            fUsefReallyHidden = Utils.BitmaskToBool(entryOperand, 0x1000000);
            fUsefAllowOverlap = Utils.BitmaskToBool(entryOperand, 0x2000000);
            fUsefUserDrawn = Utils.BitmaskToBool(entryOperand, 0x4000000);
            fUsefHorizRule = Utils.BitmaskToBool(entryOperand, 0x8000000);

            fUsefNoshadeHR = Utils.BitmaskToBool(entryOperand, 0x10000000);
            fUsefStandardHR = Utils.BitmaskToBool(entryOperand, 0x20000000);
            fUsefIsBullet = Utils.BitmaskToBool(entryOperand, 0x40000000);
            fUsefLayoutInCell = Utils.BitmaskToBool(entryOperand, 0x80000000);
        }
    }
}
