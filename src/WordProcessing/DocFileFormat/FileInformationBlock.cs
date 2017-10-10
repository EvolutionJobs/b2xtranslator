using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.DocFileFormat
{
    public class FileInformationBlock
    {
        public enum FibVersion
        {
            Fib1997Beta = 0x00C0,
            Fib1997 = 0x00C1,
            Fib2000 = 0x00D9,
            Fib2002 = 0x0101,
            Fib2003 = 0x010C,
            Fib2007 = 0x0112
        }

        #region FibBase
        public ushort wIdent;
        public FibVersion nFib;
        public ushort lid;
        public short pnNext;
        public bool fDot;
        public bool fGlsy;
        public bool fComplex;
        public bool fHasPic;
        public ushort cQuickSaves;
        public bool fEncrypted;
        public bool fWhichTblStm;
        public bool fReadOnlyRecommended;
        public bool fWriteReservation;
        public bool fExtChar;
        public bool fLoadOverwrite;
        public bool fFarEast;
        public bool fCrypto;
        public ushort nFibBack;
        public int lKey;
        public byte envr;
        public bool fMac;
        public bool fEmptySpecial;
        public bool fLoadOverridePage;
        public bool fFutureSavedUndo;
        public bool fWord97Saved;
        public int fcMin;
        public int fcMac;
        #endregion

        #region RgW97
        public short lidFE;
        #endregion

        #region RgLw97
        public int cbMac;
        public int ccpText;
        public int ccpFtn;
        public int ccpHdr;
        public int ccpAtn;
        public int ccpEdn;
        public int ccpTxbx;
        public int ccpHdrTxbx;
        #endregion

        #region FibWord97
        public uint fcStshfOrig;
        public uint lcbStshfOrig;
        public uint fcStshf;
        public uint lcbStshf;
        public uint fcPlcffndRef;
        public uint lcbPlcffndRef;
        public uint fcPlcffndTxt;
        public uint lcbPlcffndTxt;
        public uint fcPlcfandRef;
        public uint lcbPlcfandRef;
        public uint fcPlcfandTxt;
        public uint lcbPlcfandTxt;
        public uint fcPlcfSed;
        public uint lcbPlcfSed;
        public uint fcPlcPad;
        public uint lcbPlcPad;
        public uint fcPlcfPhe;
        public uint lcbPlcfPhe;
        public uint fcSttbfGlsy;
        public uint lcbSttbfGlsy;
        public uint fcPlcfGlsy;
        public uint lcbPlcfGlsy;
        public uint fcPlcfHdd;
        public uint lcbPlcfHdd;
        public uint fcPlcfBteChpx;
        public uint lcbPlcfBteChpx;
        public uint fcPlcfBtePapx;
        public uint lcbPlcfBtePapx;
        public uint fcPlcfSea;
        public uint lcbPlcfSea;
        public uint fcSttbfFfn;
        public uint lcbSttbfFfn;
        public uint fcPlcfFldMom;
        public uint lcbPlcfFldMom;
        public uint fcPlcfFldHdr;
        public uint lcbPlcfFldHdr;
        public uint fcPlcfFldFtn;
        public uint lcbPlcfFldFtn;
        public uint fcPlcfFldAtn;
        public uint lcbPlcfFldAtn;
        public uint fcPlcfFldMcr;
        public uint lcbPlcfFldMcr;
        public uint fcSttbfBkmk;
        public uint lcbSttbfBkmk;
        public uint fcPlcfBkf;
        public uint lcbPlcfBkf;
        public uint fcPlcfBkl;
        public uint lcbPlcfBkl;
        public uint fcCmds;
        public uint lcbCmds;
        public uint fcSttbfMcr;
        public uint lcbSttbfMcr;
        public uint fcPrDrvr;
        public uint lcbPrDrvr;
        public uint fcPrEnvPort;
        public uint lcbPrEnvPort;
        public uint fcPrEnvLand;
        public uint lcbPrEnvLand;
        public uint fcWss;
        public uint lcbWss;
        public uint fcDop;
        public uint lcbDop;
        public uint fcSttbfAssoc;
        public uint lcbSttbfAssoc;
        public uint fcClx;
        public uint lcbClx;
        public uint fcPlcfPgdFtn;
        public uint lcbPlcfPgdFtn;
        public uint fcAutosaveSource;
        public uint lcbAutosaveSource;
        public uint fcGrpXstAtnOwners;
        public uint lcbGrpXstAtnOwners;
        public uint fcSttbfAtnBkmk;
        public uint lcbSttbfAtnBkmk;
        public uint fcPlcSpaMom;
        public uint lcbPlcSpaMom;
        public uint fcPlcSpaHdr;
        public uint lcbPlcSpaHdr;
        public uint fcPlcfAtnBkf;
        public uint lcbPlcfAtnBkf;
        public uint fcPlcfAtnBkl;
        public uint lcbPlcfAtnBkl;
        public uint fcPms;
        public uint lcbPms;
        public uint fcFormFldSttbs;
        public uint lcbFormFldSttbs;
        public uint fcPlcfendRef;
        public uint lcbPlcfendRef;
        public uint fcPlcfendTxt;
        public uint lcbPlcfendTxt;
        public uint fcPlcfFldEdn;
        public uint lcbPlcfFldEdn;
        public uint fcDggInfo;
        public uint lcbDggInfo;
        public uint fcSttbfRMark;
        public uint lcbSttbfRMark;
        public uint fcSttbfCaption;
        public uint lcbSttbfCaption;
        public uint fcSttbfAutoCaption;
        public uint lcbSttbfAutoCaption;
        public uint fcPlcfWkb;
        public uint lcbPlcfWkb;
        public uint fcPlcfSpl;
        public uint lcbPlcfSpl;
        public uint fcPlcftxbxTxt;
        public uint lcbPlcftxbxTxt;
        public uint fcPlcfFldTxbx;
        public uint lcbPlcfFldTxbx;
        public uint fcPlcfHdrtxbxTxt;
        public uint lcbPlcfHdrtxbxTxt;
        public uint fcPlcffldHdrTxbx;
        public uint lcbPlcffldHdrTxbx;
        public uint fcStwUser;
        public uint lcbStwUser;
        public uint fcSttbTtmbd;
        public uint lcbSttbTtmbd;
        public uint fcCookieData;
        public uint lcbCookieData;
        public uint fcPgdMotherOldOld;
        public uint lcbPgdMotherOldOld;
        public uint fcBkdMotherOldOld;
        public uint lcbBkdMotherOldOld;
        public uint fcPgdFtnOldOld;
        public uint lcbPgdFtnOldOld;
        public uint fcBkdFtnOldOld;
        public uint lcbBkdFtnOldOld;
        public uint fcPgdEdnOldOld;
        public uint lcbPgdEdnOldOld;
        public uint fcBkdEdnOldOld;
        public uint lcbBkdEdnOldOld;
        public uint fcSttbfIntlFld;
        public uint lcbSttbfIntlFld;
        public uint fcRouteSlip;
        public uint lcbRouteSlip;
        public uint fcSttbSavedBy;
        public uint lcbSttbSavedBy;
        public uint fcSttbFnm;
        public uint lcbSttbFnm;
        public uint fcPlfLst;
        public uint lcbPlfLst;
        public uint fcPlfLfo;
        public uint lcbPlfLfo;
        public uint fcPlcfTxbxBkd;
        public uint lcbPlcfTxbxBkd;
        public uint fcPlcfTxbxHdrBkd;
        public uint lcbPlcfTxbxHdrBkd;
        public uint fcDocUndoWord9;
        public uint lcbDocUndoWord9;
        public uint fcRgbUse;
        public uint lcbRgbUse;
        public uint fcUsp;
        public uint lcbUsp;
        public uint fcUskf;
        public uint lcbUskf;
        public uint fcPlcupcRgbUse;
        public uint lcbPlcupcRgbUse;
        public uint fcPlcupcUsp;
        public uint lcbPlcupcUsp;
        public uint fcSttbGlsyStyle;
        public uint lcbSttbGlsyStyle;
        public uint fcPlgosl;
        public uint lcbPlgosl;
        public uint fcPlcocx;
        public uint lcbPlcocx;
        public uint fcPlcfBteLvc;
        public uint lcbPlcfBteLvc;
        public uint dwLowDateTime;
        public uint dwHighDateTime;
        public uint fcPlcfLvcPre10;
        public uint lcbPlcfLvcPre10;
        public uint fcPlcfAsumy;
        public uint lcbPlcfAsumy;
        public uint fcPlcfGram;
        public uint lcbPlcfGram;
        public uint fcSttbListNames;
        public uint lcbSttbListNames;
        public uint fcSttbfUssr;
        public uint lcbSttbfUssr;
        #endregion

        #region FibWord2000
        public uint fcPlcfTch;
        public uint lcbPlcfTch;
        public uint fcRmdThreading;
        public uint lcbRmdThreading;
        public uint fcMid;
        public uint lcbMid;
        public uint fcSttbRgtplc;
        public uint lcbSttbRgtplc;
        public uint fcMsoEnvelope;
        public uint lcbMsoEnvelope;
        public uint fcPlcfLad;
        public uint lcbPlcfLad;
        public uint fcRgDofr;
        public uint lcbRgDofr;
        public uint fcPlcosl;
        public uint lcbPlcosl;
        public uint fcPlcfCookieOld;
        public uint lcbPlcfCookieOld;
        public uint fcPgdMotherOld;
        public uint lcbPgdMotherOld;
        public uint fcBkdMotherOld;
        public uint lcbBkdMotherOld;
        public uint fcPgdFtnOld;
        public uint lcbPgdFtnOld;
        public uint fcBkdFtnOld;
        public uint lcbBkdFtnOld;
        public uint fcPgdEdnOld;
        public uint lcbPgdEdnOld;
        public uint fcBkdEdnOld;
        public uint lcbBkdEdnOld;
        #endregion

        #region Fib2002
        public uint fcPlcfPgp;
        public uint lcbPlcfPgp;
        public uint fcPlcfuim;
        public uint lcbPlcfuim;
        public uint fcPlfguidUim;
        public uint lcbPlfguidUim;
        public uint fcAtrdExtra;
        public uint lcbAtrdExtra;
        public uint fcPlrsid;
        public uint lcbPlrsid;
        public uint fcSttbfBkmkFactoid;
        public uint lcbSttbfBkmkFactoid;
        public uint fcPlcfBkfFactoid;
        public uint lcbPlcfBkfFactoid;
        public uint fcPlcfcookie;
        public uint lcbPlcfcookie;
        public uint fcPlcfBklFactoid;
        public uint lcbPlcfBklFactoid;
        public uint fcFactoidData;
        public uint lcbFactoidData;
        public uint fcDocUndo;
        public uint lcbDocUndo;
        public uint fcSttbfBkmkFcc;
        public uint lcbSttbfBkmkFcc;
        public uint fcPlcfBkfFcc;
        public uint lcbPlcfBkfFcc;
        public uint fcPlcfBklFcc;
        public uint lcbPlcfBklFcc;
        public uint fcSttbfbkmkBPRepairs;
        public uint lcbSttbfbkmkBPRepairs;
        public uint fcPlcfbkfBPRepairs;
        public uint lcbPlcfbkfBPRepairs;
        public uint fcPlcfbklBPRepairs;
        public uint lcbPlcfbklBPRepairs;
        public uint fcPmsNew;
        public uint lcbPmsNew;
        public uint fcODSO;
        public uint lcbODSO;
        public uint fcPlcfpmiOldXP;
        public uint lcbPlcfpmiOldXP;
        public uint fcPlcfpmiNewXP;
        public uint lcbPlcfpmiNewXP;
        public uint fcPlcfpmiMixedXP;
        public uint lcbPlcfpmiMixedXP;
        public uint fcPlcffactoid;
        public uint lcbPlcffactoid;
        public uint fcPlcflvcOldXP;
        public uint lcbPlcflvcOldXP;
        public uint fcPlcflvcNewXP;
        public uint lcbPlcflvcNewXP;
        public uint fcPlcflvcMixedXP;
        public uint lcbPlcflvcMixedXP;
        #endregion

        #region Fib2003
        public uint fcHplxsdr;
        public uint lcbHplxsdr;
        public uint fcSttbfBkmkSdt;
        public uint lcbSttbfBkmkSdt;
        public uint fcPlcfBkfSdt;
        public uint lcbPlcfBkfSdt;
        public uint fcPlcfBklSdt;
        public uint lcbPlcfBklSdt;
        public uint fcCustomXForm;
        public uint lcbCustomXForm;
        public uint fcSttbfBkmkProt;
        public uint lcbSttbfBkmkProt;
        public uint fcPlcfBkfProt;
        public uint lcbPlcfBkfProt;
        public uint fcPlcfBklProt;
        public uint lcbPlcfBklProt;
        public uint fcSttbProtUser;
        public uint lcbSttbProtUser;
        public uint fcPlcfpmiOld;
        public uint lcbPlcfpmiOld;
        public uint fcPlcfpmiOldInline;
        public uint lcbPlcfpmiOldInline;
        public uint fcPlcfpmiNew;
        public uint lcbPlcfpmiNew;
        public uint fcPlcfpmiNewInline;
        public uint lcbPlcfpmiNewInline;
        public uint fcPlcflvcOld;
        public uint lcbPlcflvcOld;
        public uint fcPlcflvcOldInline;
        public uint lcbPlcflvcOldInline;
        public uint fcPlcflvcNew;
        public uint lcbPlcflvcNew;
        public uint fcPlcflvcNewInline;
        public uint lcbPlcflvcNewInline;
        public uint fcPgdMother;
        public uint lcbPgdMother;
        public uint fcBkdMother;
        public uint lcbBkdMother;
        public uint fcAfdMother;
        public uint lcbAfdMother;
        public uint fcPgdFtn;
        public uint lcbPgdFtn;
        public uint fcBkdFtn;
        public uint lcbBkdFtn;
        public uint fcAfdFtn;
        public uint lcbAfdFtn;
        public uint fcPgdEdn;
        public uint lcbPgdEdn;
        public uint fcBkdEdn;
        public uint lcbBkdEdn;
        public uint fcAfdEdn;
        public uint lcbAfdEdn;
        public uint fcAfd;
        public uint lcbAfd;
        #endregion

        #region Fib2007
        public uint fcPlcfmthd;
        public uint lcbPlcfmthd;
        public uint fcSttbfBkmkMoveFrom;
        public uint lcbSttbfBkmkMoveFrom;
        public uint fcPlcfBkfMoveFrom;
        public uint lcbPlcfBkfMoveFrom;
        public uint fcPlcfBklMoveFrom;
        public uint lcbPlcfBklMoveFrom;
        public uint fcSttbfBkmkMoveTo;
        public uint lcbSttbfBkmkMoveTo;
        public uint fcPlcfBkfMoveTo;
        public uint lcbPlcfBkfMoveTo;
        public uint fcPlcfBklMoveTo;
        public uint lcbPlcfBklMoveTo;
        public uint fcSttbfBkmkArto;
        public uint lcbSttbfBkmkArto;
        public uint fcPlcfBkfArto;
        public uint lcbPlcfBkfArto;
        public uint fcPlcfBklArto;
        public uint lcbPlcfBklArto;
        public uint fcArtoData;
        public uint lcbArtoData;
        public uint fcOssTheme;
        public uint lcbOssTheme;
        public uint fcColorSchemeMapping;
        public uint lcbColorSchemeMapping;
        #endregion

        #region FibNew
        public FibVersion nFibNew;
        public ushort cQuickSavesNew;
        #endregion

        #region others
        public ushort csw;
        public ushort cslw;
        public ushort cbRgFcLcb;
        public ushort cswNew;
        #endregion

        //*****************************************************************************************
        //                                                                              CONSTRUCTOR
        //*****************************************************************************************

        public FileInformationBlock(VirtualStreamReader reader)
        {
            ushort flag16 = 0;
            byte flag8 = 0;

            //read the FIB base
            this.wIdent = reader.ReadUInt16();
            this.nFib = (FibVersion)reader.ReadUInt16();
            reader.ReadBytes(2);
            this.lid = reader.ReadUInt16();
            this.pnNext = reader.ReadInt16();
            flag16 = reader.ReadUInt16();
            this.fDot = Utils.BitmaskToBool((int)flag16, 0x0001);
            this.fGlsy = Utils.BitmaskToBool((int)flag16, 0x0002);
            this.fComplex = Utils.BitmaskToBool((int)flag16, 0x0002);
            this.fHasPic = Utils.BitmaskToBool((int)flag16, 0x0008);
            this.cQuickSaves = (ushort)(((int)flag16 & 0x00F0) >> 4);
            this.fEncrypted = Utils.BitmaskToBool((int)flag16, 0x0100);
            this.fWhichTblStm = Utils.BitmaskToBool((int)flag16, 0x0200);
            this.fReadOnlyRecommended = Utils.BitmaskToBool((int)flag16, 0x0400);
            this.fWriteReservation = Utils.BitmaskToBool((int)flag16, 0x0800);
            this.fExtChar = Utils.BitmaskToBool((int)flag16, 0x1000);
            this.fLoadOverwrite = Utils.BitmaskToBool((int)flag16, 0x2000);
            this.fFarEast = Utils.BitmaskToBool((int)flag16, 0x4000);
            this.fCrypto = Utils.BitmaskToBool((int)flag16, 0x8000);
            this.nFibBack = reader.ReadUInt16();
            this.lKey = reader.ReadInt32();
            this.envr = reader.ReadByte();
            flag8 = reader.ReadByte();
            this.fMac = Utils.BitmaskToBool((int)flag8, 0x01);
            this.fEmptySpecial = Utils.BitmaskToBool((int)flag8, 0x02);
            this.fLoadOverridePage = Utils.BitmaskToBool((int)flag8, 0x04);
            this.fFutureSavedUndo = Utils.BitmaskToBool((int)flag8, 0x08);
            this.fWord97Saved = Utils.BitmaskToBool((int)flag8, 0x10);
            reader.ReadBytes(4);
            this.fcMin = reader.ReadInt32();
            this.fcMac = reader.ReadInt32();

            this.csw = reader.ReadUInt16();

            //read the RgW97
            reader.ReadBytes(26);
            this.lidFE = reader.ReadInt16();

            this.cslw = reader.ReadUInt16();

            //read the RgLW97
            this.cbMac = reader.ReadInt32();
            reader.ReadBytes(8);
            this.ccpText = reader.ReadInt32();
            this.ccpFtn = reader.ReadInt32();
            this.ccpHdr = reader.ReadInt32();
            reader.ReadBytes(4);
            this.ccpAtn = reader.ReadInt32();
            this.ccpEdn = reader.ReadInt32();
            this.ccpTxbx = reader.ReadInt32();
            this.ccpHdrTxbx = reader.ReadInt32();
            reader.ReadBytes(44);

            this.cbRgFcLcb = reader.ReadUInt16();

            if (this.nFib >= FibVersion.Fib1997Beta)
            {
                //Read the FibRgFcLcb97
                this.fcStshfOrig = reader.ReadUInt32();
                this.lcbStshfOrig = reader.ReadUInt32();
                this.fcStshf = reader.ReadUInt32();
                this.lcbStshf = reader.ReadUInt32();
                this.fcPlcffndRef = reader.ReadUInt32();
                this.lcbPlcffndRef = reader.ReadUInt32();
                this.fcPlcffndTxt = reader.ReadUInt32();
                this.lcbPlcffndTxt = reader.ReadUInt32();
                this.fcPlcfandRef = reader.ReadUInt32();
                this.lcbPlcfandRef = reader.ReadUInt32();
                this.fcPlcfandTxt = reader.ReadUInt32();
                this.lcbPlcfandTxt = reader.ReadUInt32();
                this.fcPlcfSed = reader.ReadUInt32();
                this.lcbPlcfSed = reader.ReadUInt32();
                this.fcPlcPad = reader.ReadUInt32();
                this.lcbPlcPad = reader.ReadUInt32();
                this.fcPlcfPhe = reader.ReadUInt32();
                this.lcbPlcfPhe = reader.ReadUInt32();
                this.fcSttbfGlsy = reader.ReadUInt32();
                this.lcbSttbfGlsy = reader.ReadUInt32();
                this.fcPlcfGlsy = reader.ReadUInt32();
                this.lcbPlcfGlsy = reader.ReadUInt32();
                this.fcPlcfHdd = reader.ReadUInt32();
                this.lcbPlcfHdd = reader.ReadUInt32();
                this.fcPlcfBteChpx = reader.ReadUInt32();
                this.lcbPlcfBteChpx = reader.ReadUInt32();
                this.fcPlcfBtePapx = reader.ReadUInt32();
                this.lcbPlcfBtePapx = reader.ReadUInt32();
                this.fcPlcfSea = reader.ReadUInt32();
                this.lcbPlcfSea = reader.ReadUInt32();
                this.fcSttbfFfn = reader.ReadUInt32();
                this.lcbSttbfFfn = reader.ReadUInt32();
                this.fcPlcfFldMom = reader.ReadUInt32();
                this.lcbPlcfFldMom = reader.ReadUInt32();
                this.fcPlcfFldHdr = reader.ReadUInt32();
                this.lcbPlcfFldHdr = reader.ReadUInt32();
                this.fcPlcfFldFtn = reader.ReadUInt32();
                this.lcbPlcfFldFtn = reader.ReadUInt32();
                this.fcPlcfFldAtn = reader.ReadUInt32();
                this.lcbPlcfFldAtn = reader.ReadUInt32();
                this.fcPlcfFldMcr = reader.ReadUInt32();
                this.lcbPlcfFldMcr = reader.ReadUInt32();
                this.fcSttbfBkmk = reader.ReadUInt32();
                this.lcbSttbfBkmk = reader.ReadUInt32();
                this.fcPlcfBkf = reader.ReadUInt32();
                this.lcbPlcfBkf = reader.ReadUInt32();
                this.fcPlcfBkl = reader.ReadUInt32();
                this.lcbPlcfBkl = reader.ReadUInt32();
                this.fcCmds = reader.ReadUInt32();
                this.lcbCmds = reader.ReadUInt32();
                reader.ReadUInt32();
                reader.ReadUInt32();
                this.fcSttbfMcr = reader.ReadUInt32();
                this.lcbSttbfMcr = reader.ReadUInt32();
                this.fcPrDrvr = reader.ReadUInt32();
                this.lcbPrDrvr = reader.ReadUInt32();
                this.fcPrEnvPort = reader.ReadUInt32();
                this.lcbPrEnvPort = reader.ReadUInt32();
                this.fcPrEnvLand = reader.ReadUInt32();
                this.lcbPrEnvLand = reader.ReadUInt32();
                this.fcWss = reader.ReadUInt32();
                this.lcbWss = reader.ReadUInt32();
                this.fcDop = reader.ReadUInt32();
                this.lcbDop = reader.ReadUInt32();
                this.fcSttbfAssoc = reader.ReadUInt32();
                this.lcbSttbfAssoc = reader.ReadUInt32();
                this.fcClx = reader.ReadUInt32();
                this.lcbClx = reader.ReadUInt32();
                this.fcPlcfPgdFtn = reader.ReadUInt32();
                this.lcbPlcfPgdFtn = reader.ReadUInt32();
                this.fcAutosaveSource = reader.ReadUInt32();
                this.lcbAutosaveSource = reader.ReadUInt32();
                this.fcGrpXstAtnOwners = reader.ReadUInt32();
                this.lcbGrpXstAtnOwners = reader.ReadUInt32();
                this.fcSttbfAtnBkmk = reader.ReadUInt32();
                this.lcbSttbfAtnBkmk = reader.ReadUInt32();
                reader.ReadUInt32();
                reader.ReadUInt32();
                reader.ReadUInt32();
                reader.ReadUInt32();
                this.fcPlcSpaMom = reader.ReadUInt32();
                this.lcbPlcSpaMom = reader.ReadUInt32();
                this.fcPlcSpaHdr = reader.ReadUInt32();
                this.lcbPlcSpaHdr = reader.ReadUInt32();
                this.fcPlcfAtnBkf = reader.ReadUInt32();
                this.lcbPlcfAtnBkf = reader.ReadUInt32();
                this.fcPlcfAtnBkl = reader.ReadUInt32();
                this.lcbPlcfAtnBkl = reader.ReadUInt32();
                this.fcPms = reader.ReadUInt32();
                this.lcbPms = reader.ReadUInt32();
                this.fcFormFldSttbs = reader.ReadUInt32();
                this.lcbFormFldSttbs = reader.ReadUInt32();
                this.fcPlcfendRef = reader.ReadUInt32();
                this.lcbPlcfendRef = reader.ReadUInt32();
                this.fcPlcfendTxt = reader.ReadUInt32();
                this.lcbPlcfendTxt = reader.ReadUInt32();
                this.fcPlcfFldEdn = reader.ReadUInt32();
                this.lcbPlcfFldEdn = reader.ReadUInt32();
                reader.ReadUInt32();
                reader.ReadUInt32();
                this.fcDggInfo = reader.ReadUInt32();
                this.lcbDggInfo = reader.ReadUInt32();
                this.fcSttbfRMark = reader.ReadUInt32();
                this.lcbSttbfRMark = reader.ReadUInt32();
                this.fcSttbfCaption = reader.ReadUInt32();
                this.lcbSttbfCaption = reader.ReadUInt32();
                this.fcSttbfAutoCaption = reader.ReadUInt32();
                this.lcbSttbfAutoCaption = reader.ReadUInt32();
                this.fcPlcfWkb = reader.ReadUInt32();
                this.lcbPlcfWkb = reader.ReadUInt32();
                this.fcPlcfSpl = reader.ReadUInt32();
                this.lcbPlcfSpl = reader.ReadUInt32();
                this.fcPlcftxbxTxt = reader.ReadUInt32();
                this.lcbPlcftxbxTxt = reader.ReadUInt32();
                this.fcPlcfFldTxbx = reader.ReadUInt32();
                this.lcbPlcfFldTxbx = reader.ReadUInt32();
                this.fcPlcfHdrtxbxTxt = reader.ReadUInt32();
                this.lcbPlcfHdrtxbxTxt = reader.ReadUInt32();
                this.fcPlcffldHdrTxbx = reader.ReadUInt32();
                this.lcbPlcffldHdrTxbx = reader.ReadUInt32();
                this.fcStwUser = reader.ReadUInt32();
                this.lcbStwUser = reader.ReadUInt32();
                this.fcSttbTtmbd = reader.ReadUInt32();
                this.lcbSttbTtmbd = reader.ReadUInt32();
                this.fcCookieData = reader.ReadUInt32();
                this.lcbCookieData = reader.ReadUInt32();
                this.fcPgdMotherOldOld = reader.ReadUInt32();
                this.lcbPgdMotherOldOld = reader.ReadUInt32();
                this.fcBkdMotherOldOld = reader.ReadUInt32();
                this.lcbBkdMotherOldOld = reader.ReadUInt32();
                this.fcPgdFtnOldOld = reader.ReadUInt32();
                this.lcbPgdFtnOldOld = reader.ReadUInt32();
                this.fcBkdFtnOldOld = reader.ReadUInt32();
                this.lcbBkdFtnOldOld = reader.ReadUInt32();
                this.fcPgdEdnOldOld = reader.ReadUInt32();
                this.lcbPgdEdnOldOld = reader.ReadUInt32();
                this.fcBkdEdnOldOld = reader.ReadUInt32();
                this.lcbBkdEdnOldOld = reader.ReadUInt32();
                this.fcSttbfIntlFld = reader.ReadUInt32();
                this.lcbSttbfIntlFld = reader.ReadUInt32();
                this.fcRouteSlip = reader.ReadUInt32();
                this.lcbRouteSlip = reader.ReadUInt32();
                this.fcSttbSavedBy = reader.ReadUInt32();
                this.lcbSttbSavedBy = reader.ReadUInt32();
                this.fcSttbFnm = reader.ReadUInt32();
                this.lcbSttbFnm = reader.ReadUInt32();
                this.fcPlfLst = reader.ReadUInt32();
                this.lcbPlfLst = reader.ReadUInt32();
                this.fcPlfLfo = reader.ReadUInt32();
                this.lcbPlfLfo = reader.ReadUInt32();
                this.fcPlcfTxbxBkd = reader.ReadUInt32();
                this.lcbPlcfTxbxBkd = reader.ReadUInt32();
                this.fcPlcfTxbxHdrBkd = reader.ReadUInt32();
                this.lcbPlcfTxbxHdrBkd = reader.ReadUInt32();
                this.fcDocUndoWord9 = reader.ReadUInt32();
                this.lcbDocUndoWord9 = reader.ReadUInt32();
                this.fcRgbUse = reader.ReadUInt32();
                this.lcbRgbUse = reader.ReadUInt32();
                this.fcUsp = reader.ReadUInt32();
                this.lcbUsp = reader.ReadUInt32();
                this.fcUskf = reader.ReadUInt32();
                this.lcbUskf = reader.ReadUInt32();
                this.fcPlcupcRgbUse = reader.ReadUInt32();
                this.lcbPlcupcRgbUse = reader.ReadUInt32();
                this.fcPlcupcUsp = reader.ReadUInt32();
                this.lcbPlcupcUsp = reader.ReadUInt32();
                this.fcSttbGlsyStyle = reader.ReadUInt32();
                this.lcbSttbGlsyStyle = reader.ReadUInt32();
                this.fcPlgosl = reader.ReadUInt32();
                this.lcbPlgosl = reader.ReadUInt32();
                this.fcPlcocx = reader.ReadUInt32();
                this.lcbPlcocx = reader.ReadUInt32();
                this.fcPlcfBteLvc = reader.ReadUInt32();
                this.lcbPlcfBteLvc = reader.ReadUInt32();
                this.dwLowDateTime = reader.ReadUInt32();
                this.dwHighDateTime = reader.ReadUInt32();
                this.fcPlcfLvcPre10 = reader.ReadUInt32();
                this.lcbPlcfLvcPre10 = reader.ReadUInt32();
                this.fcPlcfAsumy = reader.ReadUInt32();
                this.lcbPlcfAsumy = reader.ReadUInt32();
                this.fcPlcfGram = reader.ReadUInt32();
                this.lcbPlcfGram = reader.ReadUInt32();
                this.fcSttbListNames = reader.ReadUInt32();
                this.lcbSttbListNames = reader.ReadUInt32();
                this.fcSttbfUssr = reader.ReadUInt32();
                this.lcbSttbfUssr = reader.ReadUInt32();
            }
            if (this.nFib >= FibVersion.Fib2000)
            {
                //Read also the FibRgFcLcb2000
                this.fcPlcfTch = reader.ReadUInt32();
                this.lcbPlcfTch = reader.ReadUInt32();
                this.fcRmdThreading = reader.ReadUInt32();
                this.lcbRmdThreading = reader.ReadUInt32();
                this.fcMid = reader.ReadUInt32();
                this.lcbMid = reader.ReadUInt32();
                this.fcSttbRgtplc = reader.ReadUInt32();
                this.lcbSttbRgtplc = reader.ReadUInt32();
                this.fcMsoEnvelope = reader.ReadUInt32();
                this.lcbMsoEnvelope = reader.ReadUInt32();
                this.fcPlcfLad = reader.ReadUInt32();
                this.lcbPlcfLad = reader.ReadUInt32();
                this.fcRgDofr = reader.ReadUInt32();
                this.lcbRgDofr = reader.ReadUInt32();
                this.fcPlcosl = reader.ReadUInt32();
                this.lcbPlcosl = reader.ReadUInt32();
                this.fcPlcfCookieOld = reader.ReadUInt32();
                this.lcbPlcfCookieOld = reader.ReadUInt32();
                this.fcPgdMotherOld = reader.ReadUInt32();
                this.lcbPgdMotherOld = reader.ReadUInt32();
                this.fcBkdMotherOld = reader.ReadUInt32();
                this.lcbBkdMotherOld = reader.ReadUInt32();
                this.fcPgdFtnOld = reader.ReadUInt32();
                this.lcbPgdFtnOld = reader.ReadUInt32();
                this.fcBkdFtnOld = reader.ReadUInt32();
                this.lcbBkdFtnOld = reader.ReadUInt32();
                this.fcPgdEdnOld = reader.ReadUInt32();
                this.lcbPgdEdnOld = reader.ReadUInt32();
                this.fcBkdEdnOld = reader.ReadUInt32();
                this.lcbBkdEdnOld = reader.ReadUInt32();
            }
            if (this.nFib >= FibVersion.Fib2002)
            {
                //Read also the fibRgFcLcb2002
                reader.ReadUInt32();
                reader.ReadUInt32();
                this.fcPlcfPgp = reader.ReadUInt32();
                this.lcbPlcfPgp = reader.ReadUInt32();
                this.fcPlcfuim = reader.ReadUInt32();
                this.lcbPlcfuim = reader.ReadUInt32();
                this.fcPlfguidUim = reader.ReadUInt32();
                this.lcbPlfguidUim = reader.ReadUInt32();
                this.fcAtrdExtra = reader.ReadUInt32();
                this.lcbAtrdExtra = reader.ReadUInt32();
                this.fcPlrsid = reader.ReadUInt32();
                this.lcbPlrsid = reader.ReadUInt32();
                this.fcSttbfBkmkFactoid = reader.ReadUInt32();
                this.lcbSttbfBkmkFactoid = reader.ReadUInt32();
                this.fcPlcfBkfFactoid = reader.ReadUInt32();
                this.lcbPlcfBkfFactoid = reader.ReadUInt32();
                this.fcPlcfcookie = reader.ReadUInt32();
                this.lcbPlcfcookie = reader.ReadUInt32();
                this.fcPlcfBklFactoid = reader.ReadUInt32();
                this.lcbPlcfBklFactoid = reader.ReadUInt32();
                this.fcFactoidData = reader.ReadUInt32();
                this.lcbFactoidData = reader.ReadUInt32();
                this.fcDocUndo = reader.ReadUInt32();
                this.lcbDocUndo = reader.ReadUInt32();
                this.fcSttbfBkmkFcc = reader.ReadUInt32();
                this.lcbSttbfBkmkFcc = reader.ReadUInt32();
                this.fcPlcfBkfFcc = reader.ReadUInt32();
                this.lcbPlcfBkfFcc = reader.ReadUInt32();
                this.fcPlcfBklFcc = reader.ReadUInt32();
                this.lcbPlcfBklFcc = reader.ReadUInt32();
                this.fcSttbfbkmkBPRepairs = reader.ReadUInt32();
                this.lcbSttbfbkmkBPRepairs = reader.ReadUInt32();
                this.fcPlcfbkfBPRepairs = reader.ReadUInt32();
                this.lcbPlcfbkfBPRepairs = reader.ReadUInt32();
                this.fcPlcfbklBPRepairs = reader.ReadUInt32();
                this.lcbPlcfbklBPRepairs = reader.ReadUInt32();
                this.fcPmsNew = reader.ReadUInt32();
                this.lcbPmsNew = reader.ReadUInt32();
                this.fcODSO = reader.ReadUInt32();
                this.lcbODSO = reader.ReadUInt32();
                this.fcPlcfpmiOldXP = reader.ReadUInt32();
                this.lcbPlcfpmiOldXP = reader.ReadUInt32();
                this.fcPlcfpmiNewXP = reader.ReadUInt32();
                this.lcbPlcfpmiNewXP = reader.ReadUInt32();
                this.fcPlcfpmiMixedXP = reader.ReadUInt32();
                this.lcbPlcfpmiMixedXP = reader.ReadUInt32();
                reader.ReadUInt32();
                reader.ReadUInt32();
                this.fcPlcffactoid = reader.ReadUInt32();
                this.lcbPlcffactoid = reader.ReadUInt32();
                this.fcPlcflvcOldXP = reader.ReadUInt32();
                this.lcbPlcflvcOldXP = reader.ReadUInt32();
                this.fcPlcflvcNewXP = reader.ReadUInt32();
                this.lcbPlcflvcNewXP = reader.ReadUInt32();
                this.fcPlcflvcMixedXP = reader.ReadUInt32();
                this.lcbPlcflvcMixedXP = reader.ReadUInt32();
            }
            if (this.nFib >= FibVersion.Fib2003)
            {
                //Read also the fibRgFcLcb2003
                this.fcHplxsdr = reader.ReadUInt32();
                this.lcbHplxsdr = reader.ReadUInt32();
                this.fcSttbfBkmkSdt = reader.ReadUInt32();
                this.lcbSttbfBkmkSdt = reader.ReadUInt32();
                this.fcPlcfBkfSdt = reader.ReadUInt32();
                this.lcbPlcfBkfSdt = reader.ReadUInt32();
                this.fcPlcfBklSdt = reader.ReadUInt32();
                this.lcbPlcfBklSdt = reader.ReadUInt32();
                this.fcCustomXForm = reader.ReadUInt32();
                this.lcbCustomXForm = reader.ReadUInt32();
                this.fcSttbfBkmkProt = reader.ReadUInt32();
                this.lcbSttbfBkmkProt = reader.ReadUInt32();
                this.fcPlcfBkfProt = reader.ReadUInt32();
                this.lcbPlcfBkfProt = reader.ReadUInt32();
                this.fcPlcfBklProt = reader.ReadUInt32();
                this.lcbPlcfBklProt = reader.ReadUInt32();
                this.fcSttbProtUser = reader.ReadUInt32();
                this.lcbSttbProtUser = reader.ReadUInt32();
                reader.ReadUInt32();
                reader.ReadUInt32();
                this.fcPlcfpmiOld = reader.ReadUInt32();
                this.lcbPlcfpmiOld = reader.ReadUInt32();
                this.fcPlcfpmiOldInline = reader.ReadUInt32();
                this.lcbPlcfpmiOldInline = reader.ReadUInt32();
                this.fcPlcfpmiNew = reader.ReadUInt32();
                this.lcbPlcfpmiNew = reader.ReadUInt32();
                this.fcPlcfpmiNewInline = reader.ReadUInt32();
                this.lcbPlcfpmiNewInline = reader.ReadUInt32();
                this.fcPlcflvcOld = reader.ReadUInt32();
                this.lcbPlcflvcOld = reader.ReadUInt32();
                this.fcPlcflvcOldInline = reader.ReadUInt32();
                this.lcbPlcflvcOldInline = reader.ReadUInt32();
                this.fcPlcflvcNew = reader.ReadUInt32();
                this.lcbPlcflvcNew = reader.ReadUInt32();
                this.fcPlcflvcNewInline = reader.ReadUInt32();
                this.lcbPlcflvcNewInline = reader.ReadUInt32();
                this.fcPgdMother = reader.ReadUInt32();
                this.lcbPgdMother = reader.ReadUInt32();
                this.fcBkdMother = reader.ReadUInt32();
                this.lcbBkdMother = reader.ReadUInt32();
                this.fcAfdMother = reader.ReadUInt32();
                this.lcbAfdMother = reader.ReadUInt32();
                this.fcPgdFtn = reader.ReadUInt32();
                this.lcbPgdFtn = reader.ReadUInt32();
                this.fcBkdFtn = reader.ReadUInt32();
                this.lcbBkdFtn = reader.ReadUInt32();
                this.fcAfdFtn = reader.ReadUInt32();
                this.lcbAfdFtn = reader.ReadUInt32();
                this.fcPgdEdn = reader.ReadUInt32();
                this.lcbPgdEdn = reader.ReadUInt32();
                this.fcBkdEdn = reader.ReadUInt32();
                this.lcbBkdEdn = reader.ReadUInt32();
                this.fcAfdEdn = reader.ReadUInt32();
                this.lcbAfdEdn = reader.ReadUInt32();
                this.fcAfd = reader.ReadUInt32();
                this.lcbAfd = reader.ReadUInt32();
            }
            if (this.nFib >= FibVersion.Fib2007)
            {
                //Read also the fibRgFcLcb2007
                this.fcPlcfmthd = reader.ReadUInt32();
                this.lcbPlcfmthd = reader.ReadUInt32();
                this.fcSttbfBkmkMoveFrom = reader.ReadUInt32();
                this.lcbSttbfBkmkMoveFrom = reader.ReadUInt32();
                this.fcPlcfBkfMoveFrom = reader.ReadUInt32();
                this.lcbPlcfBkfMoveFrom = reader.ReadUInt32();
                this.fcPlcfBklMoveFrom = reader.ReadUInt32();
                this.lcbPlcfBklMoveFrom = reader.ReadUInt32();
                this.fcSttbfBkmkMoveTo = reader.ReadUInt32();
                this.lcbSttbfBkmkMoveTo = reader.ReadUInt32();
                this.fcPlcfBkfMoveTo = reader.ReadUInt32();
                this.lcbPlcfBkfMoveTo = reader.ReadUInt32();
                this.fcPlcfBklMoveTo = reader.ReadUInt32();
                this.lcbPlcfBklMoveTo = reader.ReadUInt32();
                reader.ReadUInt32();
                reader.ReadUInt32();
                reader.ReadUInt32();
                reader.ReadUInt32();
                reader.ReadUInt32();
                reader.ReadUInt32();
                this.fcSttbfBkmkArto = reader.ReadUInt32();
                this.lcbSttbfBkmkArto = reader.ReadUInt32();
                this.fcPlcfBkfArto = reader.ReadUInt32();
                this.lcbPlcfBkfArto = reader.ReadUInt32();
                this.fcPlcfBklArto = reader.ReadUInt32();
                this.lcbPlcfBklArto = reader.ReadUInt32();
                this.fcArtoData = reader.ReadUInt32();
                this.lcbArtoData = reader.ReadUInt32();
                reader.ReadUInt32();
                reader.ReadUInt32();
                reader.ReadUInt32();
                reader.ReadUInt32();
                reader.ReadUInt32();
                reader.ReadUInt32();
                this.fcOssTheme = reader.ReadUInt32();
                this.lcbOssTheme = reader.ReadUInt32();
                this.fcColorSchemeMapping = reader.ReadUInt32();
                this.lcbColorSchemeMapping = reader.ReadUInt32();
            }

            this.cswNew = reader.ReadUInt16();

            if (this.cswNew != 0)
            {
                //Read the FibRgCswNew
                this.nFibNew = (FibVersion)reader.ReadUInt16();
                this.cQuickSavesNew = reader.ReadUInt16();
            }
        }
    }
}
