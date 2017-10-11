using System;
using System.Collections.Generic;
using b2xtranslator.Tools;

namespace b2xtranslator.DocFileFormat
{
    public class CharacterProperties
    {
        public bool fBold;
        public bool fItalic;
        public bool fRMarkDel;
        public bool fOutline;
        public bool fFldVanish;
        public bool fSmallCaps;
        public bool fCaps;
        public bool fVanish;
        public bool fRMark;
        public bool fSpec;
        public bool fStrike;
        public bool fObj;
        public bool fShadow;
        public bool fLowerCase;
        public bool fData;
        public bool fOle2;
        public bool fEmboss;
        public bool fImprint;
        public bool fDStrike;
        public bool fBoldBi;
        public bool fComplexScripts;
        public bool fItalicBi;
        public bool fBiDi;
        public bool fIcoBi;
        public bool fNonGlyph;
        public bool fBoldOther;
        public bool fItalicOther;
        public bool fNoProof;
        public bool fWebHidden;
        public bool fFitText;
        public bool fCalc;
        public bool fFmtLineProp;
        public ushort hps;
        //ftc;
        //ftcAsci;
        //ftcFE;
        //ftcOther;
        //ftcBi;
        public int dxaSpace;
        public RGBColor cv;
        public Global.ColorIdentifier ico;
        public ushort pctCharWidth;
        public short lid;
        public short lidDefault;
        public short lidFE;
        public short lidBi;
        public byte kcd;
        public bool fUndetermine;
        public byte iss;
        public bool fSpecSymbol;
        public byte idct;
        public byte idctHint;
        public Global.UnderlineCode kul;
        public byte hres;
        public byte chHres;
        public ushort hpsKern;
        public ushort hpsPos;
        public RGBColor cvUl;
        public ShadingDescriptor shd;
        public BorderCode brc;
        public short ibstRMark;
        public Global.TextAnimation sfxtText;
        public bool fDblBdr;
        public bool fBorderWS;
        public ushort ufel;
        public Global.FarEastLayout itypFELayout;
        public bool fTNY;
        public bool fWarichu;
        public bool fKumimoji;
        public bool fRuby;
        public bool fLSFitText;
        public Global.WarichuBracket iWarichuBracket;
        public bool fWarichuNoOpenBracket;
        public bool fTNYCompress;
        public bool fTNYFetchTxm;
        public bool fCellFitText;
        public ushort hpsAsci;
        public ushort hpsFE;
        public ushort hpsBi;
        //ftcSym;
        public char xchSym;
        public bool fNumRunBi;
        public bool fSysVanish;
        public bool fDiacRunBi;
        public bool fBoldPresent;
        public bool fItalicPresent;
        public int fcPic;
        public int fcObj;
        public uint lTagObj;
        public int fcData;
        public bool fDirty;
        public Global.HyphenationRule hresOld;
        public uint chHresOld;
        public int dxpKashida;
        public int dxpSpace;
        public short ibstRMarkDel;
        public DateAndTime dttmRMark;
        public DateAndTime dttmRMarkDel;
        public ushort istd;
        public ushort idslRMReason;
        public ushort idslRMReasonDel;
        public ushort cpg;
        public ushort iatrUndetType;
        public bool fUlGap;
        public Global.ColorIdentifier icoHighlight;
        public bool fHighlight;
        public bool fScriptAnchor;
        public bool fFixedObj;
        public bool fNavHighlight;
        public bool fChsDiff;
        public bool fMacChs;
        public bool fFtcAsciSym;
        public bool fFtcReq;
        public bool fLangApplied;
        public bool fSpareLangApplied;
        public bool fForcedCvAuto;
        public bool fPropRMark;
        public short ibstPropRMark;
        public DateAndTime dttmPropRMark;
        public bool fAnmPropRMark;
        public bool fConflictOrig;
        public bool fConflictOtherDel;
        public ushort wConflict;
        public short ibstConflict;
        public DateAndTime dttmConflict;
        public bool fDispFldRMark;
        public short ibstDispFldRMark;
        public DateAndTime dttmDispFldRMark;
        public string xstDispFldRMark;
        public bool fcObjp;
        public int dxaFitText;
        public int lFitTextID;
        public byte lbrCRJ;
        public int rsidProp;
        public int rsidText;
        public int rsidRMDel;
        public bool fSpecVanish;
        public bool fHasOldProps;
        public PictureBulletInformation pbi;
        public int hplcnf;
        public byte ffm;
        public bool fSdtVanish;
        public FontFamilyName FontAscii;
        public b2xtranslator.DocFileFormat.Global.UnderlineCode UnderlineStyle;

        /// <summary>
        /// Creates a CHP with default properties
        /// </summary>
        public CharacterProperties()
        {
            setDefaultValues();
        }

        /// <summary>
        /// Builds a CHP based on a CHPX
        /// </summary>
        /// <param name="styles">The stylesheet</param>
        /// <param name="chpx">The CHPX</param>
        public CharacterProperties(CharacterPropertyExceptions chpx, ParagraphPropertyExceptions parentPapx, WordDocument parentDocument)
        {
            setDefaultValues();

            //get all CHPX in the hierarchy
            var chpxHierarchy = new List<CharacterPropertyExceptions>();
            chpxHierarchy.Add(chpx);

            //add parent character styles
            buildHierarchy(chpxHierarchy, parentDocument.Styles, (ushort)getIsdt(chpx));

            //add parent paragraph styles
            buildHierarchy(chpxHierarchy, parentDocument.Styles, parentPapx.istd);

            chpxHierarchy.Reverse();

            //apply the CHPX hierarchy to this CHP
            foreach(var c in chpxHierarchy)
            {
                applyChpx(c, parentDocument);
            }
        }

        private void applyChpx(PropertyExceptions chpx, WordDocument parentDocument)
        {
            foreach (var sprm in chpx.grpprl)
            {
                switch (sprm.OpCode)
                {
                    //style id 
                    case SinglePropertyModifier.OperationCode.sprmCIstd:
                        this.istd = System.BitConverter.ToUInt16(sprm.Arguments, 0);
                        break;
                    //font name ASCII
                    case SinglePropertyModifier.OperationCode.sprmCRgFtc0:
                        this.FontAscii = (FontFamilyName)parentDocument.FontTable.Data[System.BitConverter.ToUInt16(sprm.Arguments, 0)];
                        break;
                    //font size
                    case SinglePropertyModifier.OperationCode.sprmCHps:
                        this.hps = sprm.Arguments[0];
                        break;
                    // color
                    case SinglePropertyModifier.OperationCode.sprmCCv:
                        this.cv = new RGBColor(System.BitConverter.ToInt32(sprm.Arguments, 0), RGBColor.ByteOrder.RedFirst);
                        break;
                    //bold
                    case SinglePropertyModifier.OperationCode.sprmCFBold:
                        this.fBold = handleToogleValue(this.fBold, sprm.Arguments[0]);
                        break;
                    //italic
                    case SinglePropertyModifier.OperationCode.sprmCFItalic:
                        this.fItalic = handleToogleValue(this.fItalic, sprm.Arguments[0]);
                        break;
                    //outline
                    case SinglePropertyModifier.OperationCode.sprmCFOutline:
                        this.fOutline = Utils.ByteToBool(sprm.Arguments[0]);
                        break;
                    //shadow
                    case SinglePropertyModifier.OperationCode.sprmCFShadow:
                        this.fShadow = Utils.ByteToBool(sprm.Arguments[0]);
                        break;
                    //strike through
                    case SinglePropertyModifier.OperationCode.sprmCFStrike:
                        this.fStrike = Utils.ByteToBool(sprm.Arguments[0]);
                        break;
                        // underline
                    case SinglePropertyModifier.OperationCode.sprmCKul:
                        this.UnderlineStyle = (Global.UnderlineCode)sprm.Arguments[0];
                        break;
                }
            }
        }

        private void buildHierarchy(List<CharacterPropertyExceptions> hierarchy, StyleSheet styleSheet, ushort istdStart)
        {
            int istd = (int)istdStart;
            bool goOn = true;
            while (goOn)
            {
                try
                {
                    var baseChpx = styleSheet.Styles[istd].chpx;
                    if (baseChpx != null)
                    {
                        hierarchy.Add(baseChpx);
                        istd = (int)styleSheet.Styles[istd].istdBase;
                    }
                    else
                    {
                        goOn = false;
                    }
                }
                catch (Exception)
                {
                    goOn = false;
                }
            }
        }

        private bool handleToogleValue(bool currentValue, byte toggle)
        {
            if (toggle == 1)
                return true;
            else if (toggle == 129)
                //invert the current value
                if (currentValue)
                    return false;
                else
                    return true;
            else if (toggle == 128)
                //use the current value
                return currentValue;
            else
                return false;
        }

        private void setDefaultValues()
        {
            this.hps = 20;
            this.fcPic = -1;
            this.istd = 10;
            this.lidDefault = 0x0400;
            this.lidFE = 0x0400;
        }

        private int getIsdt(CharacterPropertyExceptions chpx)
        {
            int ret = 10; //default value for istd
            foreach (var sprm in chpx.grpprl)
            {
                if (sprm.OpCode == SinglePropertyModifier.OperationCode.sprmCIstd)
                {
                    ret = (int)System.BitConverter.ToInt16(sprm.Arguments, 0);
                    break;
                }
            }
            return ret;
        }
    }
}
