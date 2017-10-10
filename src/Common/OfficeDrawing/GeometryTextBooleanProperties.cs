using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing
{
    public class GeometryTextBooleanProperties
    {
        public bool gtextFStrikethrough;
        public bool gtextFSmallcaps;
        public bool gtextFShadow;
        public bool gtextFUnderline;
        public bool gtextFItalic;
        public bool gtextFBold;
        public bool gtextFDxMeasure;
        public bool gtextFNormalize;
        public bool gtextFBestFit;
        public bool gtextFShrinkFit;
        public bool gtextFStretch;
        public bool gtextFTight;
        public bool gtextFKern;
        public bool gtextFVertical;
        public bool fGtext;
        public bool gtextFReverseRows;

        public bool fUsegtextFSStrikeThrough;
        public bool fUsegtextFSmallcaps;
        public bool fUsegtextFShadow;
        public bool fUsegtextFUnderline;
        public bool fUsegtextFItalic;
        public bool fUsegtextFBold;
        public bool fUsegtextFDxMeasure;
        public bool fUsegtextFNormalize;
        public bool fUsegtextFBestFit;
        public bool fUsegtextFShrinkFit;
        public bool fUsegtextFStretch;
        public bool fUsegtextFTight;
        public bool fUsegtextFKern;
        public bool fUsegtextFVertical;
        public bool fUsefGtext;
        public bool fUsegtextFReverseRows;

        public GeometryTextBooleanProperties(uint entryOperand)
        {
            gtextFStrikethrough = Utils.BitmaskToBool(entryOperand, 0x1);
            gtextFSmallcaps = Utils.BitmaskToBool(entryOperand, 0x1 << 1);
            gtextFShadow = Utils.BitmaskToBool(entryOperand, 0x1 << 2);
            gtextFUnderline = Utils.BitmaskToBool(entryOperand, 0x1 << 3);
            gtextFItalic = Utils.BitmaskToBool(entryOperand, 0x1 << 4);
            gtextFBold = Utils.BitmaskToBool(entryOperand, 0x1 << 5);
            gtextFDxMeasure = Utils.BitmaskToBool(entryOperand, 0x1 << 6);
            gtextFNormalize = Utils.BitmaskToBool(entryOperand, 0x1 << 7);
            gtextFBestFit = Utils.BitmaskToBool(entryOperand, 0x1 << 8);
            gtextFShrinkFit = Utils.BitmaskToBool(entryOperand, 0x1 << 9);
            gtextFStretch = Utils.BitmaskToBool(entryOperand, 0x1 << 10);
            gtextFTight = Utils.BitmaskToBool(entryOperand, 0x1 << 11);
            gtextFKern = Utils.BitmaskToBool(entryOperand, 0x1 << 12);
            gtextFVertical = Utils.BitmaskToBool(entryOperand, 0x1 << 13);
            fGtext = Utils.BitmaskToBool(entryOperand, 0x1 << 14);
            gtextFReverseRows = Utils.BitmaskToBool(entryOperand, 0x1 << 15);

            fUsegtextFSStrikeThrough = Utils.BitmaskToBool(entryOperand, 0x1 << 16);
            fUsegtextFSmallcaps = Utils.BitmaskToBool(entryOperand, 0x1 << 17);
            fUsegtextFShadow = Utils.BitmaskToBool(entryOperand, 0x1 << 18);
            fUsegtextFUnderline = Utils.BitmaskToBool(entryOperand, 0x1 << 19);
            fUsegtextFItalic = Utils.BitmaskToBool(entryOperand, 0x1 << 20);
            fUsegtextFBold = Utils.BitmaskToBool(entryOperand, 0x1 << 21);
            fUsegtextFDxMeasure = Utils.BitmaskToBool(entryOperand, 0x1 << 22);
            fUsegtextFNormalize = Utils.BitmaskToBool(entryOperand, 0x1 << 23);
            fUsegtextFBestFit = Utils.BitmaskToBool(entryOperand, 0x1 << 24);
            fUsegtextFShrinkFit = Utils.BitmaskToBool(entryOperand, 0x1 << 25);
            fUsegtextFStretch = Utils.BitmaskToBool(entryOperand, 0x1 << 26);
            fUsegtextFTight = Utils.BitmaskToBool(entryOperand, 0x1 << 27);
            fUsegtextFKern = Utils.BitmaskToBool(entryOperand, 0x1 << 28);
            fUsegtextFVertical = Utils.BitmaskToBool(entryOperand, 0x1 << 29);
            fUsefGtext = Utils.BitmaskToBool(entryOperand, 0x1 << 30);
            fUsegtextFReverseRows = Utils.BitmaskToBool(entryOperand, 0x40000000);
        }
    }
}
