using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
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
        public UInt16 hps;
        //ftc;
        //ftcAsci;
        //ftcFE;
        //ftcOther;
        //ftcBi;
        public Int32 dxaSpace;
        public RGBColor cv;
        public Global.ColorIdentifier ico;
        public UInt16 pctCharWidth;
        public Int16 lid;
        public Int16 lidDefault;
        public Int16 lidFE;
        public Int16 lidBi;
        public byte kcd;
        public bool fUndetermine;
        public byte iss;
        public bool fSpecSymbol;
        public byte idct;
        public byte idctHint;
        public Global.UnderlineCode kul;
        public byte hres;
        public byte chHres;
        public UInt16 hpsKern;
        public UInt16 hpsPos;
        public RGBColor cvUl;
        public ShadingDescriptor shd;
        public BorderCode brc;
        public Int16 ibstRMark;
        public Global.TextAnimation sfxtText;
        public bool fDblBdr;
        public bool fBorderWS;
        public UInt16 ufel;
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
        public UInt16 hpsAsci;
        public UInt16 hpsFE;
        public UInt16 hpsBi;
        //ftcSym;
        public char xchSym;
        public bool fNumRunBi;
        public bool fSysVanish;
        public bool fDiacRunBi;
        public bool fBoldPresent;
        public bool fItalicPresent;
        public Int32 fcPic;
        public Int32 fcObj;
        public UInt32 lTagObj;
        public Int32 fcData;
        public bool fDirty;
        public Global.HyphenationRule hresOld;
        public UInt32 chHresOld;
        public Int32 dxpKashida;
        public Int32 dxpSpace;
        public Int16 ibstRMarkDel;
        public DateAndTime dttmRMark;
        public DateAndTime dttmRMarkDel;
        public UInt16 istd;
        public UInt16 idslRMReason;
        public UInt16 idslRMReasonDel;
        public UInt16 cpg;
        public UInt16 iatrUndetType;
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
        public Int16 ibstPropRMark;
        public DateAndTime dttmPropRMark;
        public bool fAnmPropRMark;
        public bool fConflictOrig;
        public bool fConflictOtherDel;
        public UInt16 wConflict;
        public Int16 ibstConflict;
        public DateAndTime dttmConflict;
        public bool fDispFldRMark;
        public Int16 ibstDispFldRMark;
        public DateAndTime dttmDispFldRMark;
        public string xstDispFldRMark;
        public bool fcObjp;
        public Int32 dxaFitText;
        public Int32 lFitTextID;
        public byte lbrCRJ;
        public Int32 rsidProp;
        public Int32 rsidText;
        public Int32 rsidRMDel;
        public bool fSpecVanish;
        public bool fHasOldProps;
        public PictureBulletInformation pbi;
        public Int32 hplcnf;
        public byte ffm;
        public bool fSdtVanish;
        public FontFamilyName FontAscii;
        public DIaLOGIKa.b2xtranslator.DocFileFormat.Global.UnderlineCode UnderlineStyle;

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
            List<CharacterPropertyExceptions> chpxHierarchy = new List<CharacterPropertyExceptions>();
            chpxHierarchy.Add(chpx);

            //add parent character styles
            buildHierarchy(chpxHierarchy, parentDocument.Styles, (UInt16)getIsdt(chpx));

            //add parent paragraph styles
            buildHierarchy(chpxHierarchy, parentDocument.Styles, parentPapx.istd);

            chpxHierarchy.Reverse();

            //apply the CHPX hierarchy to this CHP
            foreach(CharacterPropertyExceptions c in chpxHierarchy)
            {
                applyChpx(c, parentDocument);
            }
        }

        private void applyChpx(PropertyExceptions chpx, WordDocument parentDocument)
        {
            foreach (SinglePropertyModifier sprm in chpx.grpprl)
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

        private void buildHierarchy(List<CharacterPropertyExceptions> hierarchy, StyleSheet styleSheet, UInt16 istdStart)
        {
            int istd = (int)istdStart;
            bool goOn = true;
            while (goOn)
            {
                try
                {
                    CharacterPropertyExceptions baseChpx = styleSheet.Styles[istd].chpx;
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
            foreach (SinglePropertyModifier sprm in chpx.grpprl)
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
