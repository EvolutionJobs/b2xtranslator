/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *        notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright
 *       notice, this list of conditions and the following disclaimer in the
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the name of DIaLOGIKa nor the
 *       names of its contributors may be used to endorse or promote products
 *       derived from this software without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY DIaLOGIKa ''AS IS'' AND ANY
 * EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 * DISCLAIMED. IN NO EVENT SHALL DIaLOGIKa BE LIABLE FOR ANY
 * DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 * (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 * ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 * (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */

using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.DocFileFormat
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
        public UInt16 wIdent;
        public FibVersion nFib;
        public UInt16 lid;
        public Int16 pnNext;
        public bool fDot;
        public bool fGlsy;
        public bool fComplex;
        public bool fHasPic;
        public UInt16 cQuickSaves;
        public bool fEncrypted;
        public bool fWhichTblStm;
        public bool fReadOnlyRecommended;
        public bool fWriteReservation;
        public bool fExtChar;
        public bool fLoadOverwrite;
        public bool fFarEast;
        public bool fCrypto;
        public UInt16 nFibBack;
        public Int32 lKey;
        public byte envr;
        public bool fMac;
        public bool fEmptySpecial;
        public bool fLoadOverridePage;
        public bool fFutureSavedUndo;
        public bool fWord97Saved;
        public Int32 fcMin;
        public Int32 fcMac;
        #endregion

        #region RgW97
        public Int16 lidFE;
        #endregion

        #region RgLw97
        public Int32 cbMac;
        public Int32 ccpText;
        public Int32 ccpFtn;
        public Int32 ccpHdr;
        public Int32 ccpAtn;
        public Int32 ccpEdn;
        public Int32 ccpTxbx;
        public Int32 ccpHdrTxbx;
        #endregion

        #region FibWord97
        public UInt32 fcStshfOrig;
        public UInt32 lcbStshfOrig;
        public UInt32 fcStshf;
        public UInt32 lcbStshf;
        public UInt32 fcPlcffndRef;
        public UInt32 lcbPlcffndRef;
        public UInt32 fcPlcffndTxt;
        public UInt32 lcbPlcffndTxt;
        public UInt32 fcPlcfandRef;
        public UInt32 lcbPlcfandRef;
        public UInt32 fcPlcfandTxt;
        public UInt32 lcbPlcfandTxt;
        public UInt32 fcPlcfSed;
        public UInt32 lcbPlcfSed;
        public UInt32 fcPlcPad;
        public UInt32 lcbPlcPad;
        public UInt32 fcPlcfPhe;
        public UInt32 lcbPlcfPhe;
        public UInt32 fcSttbfGlsy;
        public UInt32 lcbSttbfGlsy;
        public UInt32 fcPlcfGlsy;
        public UInt32 lcbPlcfGlsy;
        public UInt32 fcPlcfHdd;
        public UInt32 lcbPlcfHdd;
        public UInt32 fcPlcfBteChpx;
        public UInt32 lcbPlcfBteChpx;
        public UInt32 fcPlcfBtePapx;
        public UInt32 lcbPlcfBtePapx;
        public UInt32 fcPlcfSea;
        public UInt32 lcbPlcfSea;
        public UInt32 fcSttbfFfn;
        public UInt32 lcbSttbfFfn;
        public UInt32 fcPlcfFldMom;
        public UInt32 lcbPlcfFldMom;
        public UInt32 fcPlcfFldHdr;
        public UInt32 lcbPlcfFldHdr;
        public UInt32 fcPlcfFldFtn;
        public UInt32 lcbPlcfFldFtn;
        public UInt32 fcPlcfFldAtn;
        public UInt32 lcbPlcfFldAtn;
        public UInt32 fcPlcfFldMcr;
        public UInt32 lcbPlcfFldMcr;
        public UInt32 fcSttbfBkmk;
        public UInt32 lcbSttbfBkmk;
        public UInt32 fcPlcfBkf;
        public UInt32 lcbPlcfBkf;
        public UInt32 fcPlcfBkl;
        public UInt32 lcbPlcfBkl;
        public UInt32 fcCmds;
        public UInt32 lcbCmds;
        public UInt32 fcSttbfMcr;
        public UInt32 lcbSttbfMcr;
        public UInt32 fcPrDrvr;
        public UInt32 lcbPrDrvr;
        public UInt32 fcPrEnvPort;
        public UInt32 lcbPrEnvPort;
        public UInt32 fcPrEnvLand;
        public UInt32 lcbPrEnvLand;
        public UInt32 fcWss;
        public UInt32 lcbWss;
        public UInt32 fcDop;
        public UInt32 lcbDop;
        public UInt32 fcSttbfAssoc;
        public UInt32 lcbSttbfAssoc;
        public UInt32 fcClx;
        public UInt32 lcbClx;
        public UInt32 fcPlcfPgdFtn;
        public UInt32 lcbPlcfPgdFtn;
        public UInt32 fcAutosaveSource;
        public UInt32 lcbAutosaveSource;
        public UInt32 fcGrpXstAtnOwners;
        public UInt32 lcbGrpXstAtnOwners;
        public UInt32 fcSttbfAtnBkmk;
        public UInt32 lcbSttbfAtnBkmk;
        public UInt32 fcPlcSpaMom;
        public UInt32 lcbPlcSpaMom;
        public UInt32 fcPlcSpaHdr;
        public UInt32 lcbPlcSpaHdr;
        public UInt32 fcPlcfAtnBkf;
        public UInt32 lcbPlcfAtnBkf;
        public UInt32 fcPlcfAtnBkl;
        public UInt32 lcbPlcfAtnBkl;
        public UInt32 fcPms;
        public UInt32 lcbPms;
        public UInt32 fcFormFldSttbs;
        public UInt32 lcbFormFldSttbs;
        public UInt32 fcPlcfendRef;
        public UInt32 lcbPlcfendRef;
        public UInt32 fcPlcfendTxt;
        public UInt32 lcbPlcfendTxt;
        public UInt32 fcPlcfFldEdn;
        public UInt32 lcbPlcfFldEdn;
        public UInt32 fcDggInfo;
        public UInt32 lcbDggInfo;
        public UInt32 fcSttbfRMark;
        public UInt32 lcbSttbfRMark;
        public UInt32 fcSttbfCaption;
        public UInt32 lcbSttbfCaption;
        public UInt32 fcSttbfAutoCaption;
        public UInt32 lcbSttbfAutoCaption;
        public UInt32 fcPlcfWkb;
        public UInt32 lcbPlcfWkb;
        public UInt32 fcPlcfSpl;
        public UInt32 lcbPlcfSpl;
        public UInt32 fcPlcftxbxTxt;
        public UInt32 lcbPlcftxbxTxt;
        public UInt32 fcPlcfFldTxbx;
        public UInt32 lcbPlcfFldTxbx;
        public UInt32 fcPlcfHdrtxbxTxt;
        public UInt32 lcbPlcfHdrtxbxTxt;
        public UInt32 fcPlcffldHdrTxbx;
        public UInt32 lcbPlcffldHdrTxbx;
        public UInt32 fcStwUser;
        public UInt32 lcbStwUser;
        public UInt32 fcSttbTtmbd;
        public UInt32 lcbSttbTtmbd;
        public UInt32 fcCookieData;
        public UInt32 lcbCookieData;
        public UInt32 fcPgdMotherOldOld;
        public UInt32 lcbPgdMotherOldOld;
        public UInt32 fcBkdMotherOldOld;
        public UInt32 lcbBkdMotherOldOld;
        public UInt32 fcPgdFtnOldOld;
        public UInt32 lcbPgdFtnOldOld;
        public UInt32 fcBkdFtnOldOld;
        public UInt32 lcbBkdFtnOldOld;
        public UInt32 fcPgdEdnOldOld;
        public UInt32 lcbPgdEdnOldOld;
        public UInt32 fcBkdEdnOldOld;
        public UInt32 lcbBkdEdnOldOld;
        public UInt32 fcSttbfIntlFld;
        public UInt32 lcbSttbfIntlFld;
        public UInt32 fcRouteSlip;
        public UInt32 lcbRouteSlip;
        public UInt32 fcSttbSavedBy;
        public UInt32 lcbSttbSavedBy;
        public UInt32 fcSttbFnm;
        public UInt32 lcbSttbFnm;
        public UInt32 fcPlfLst;
        public UInt32 lcbPlfLst;
        public UInt32 fcPlfLfo;
        public UInt32 lcbPlfLfo;
        public UInt32 fcPlcfTxbxBkd;
        public UInt32 lcbPlcfTxbxBkd;
        public UInt32 fcPlcfTxbxHdrBkd;
        public UInt32 lcbPlcfTxbxHdrBkd;
        public UInt32 fcDocUndoWord9;
        public UInt32 lcbDocUndoWord9;
        public UInt32 fcRgbUse;
        public UInt32 lcbRgbUse;
        public UInt32 fcUsp;
        public UInt32 lcbUsp;
        public UInt32 fcUskf;
        public UInt32 lcbUskf;
        public UInt32 fcPlcupcRgbUse;
        public UInt32 lcbPlcupcRgbUse;
        public UInt32 fcPlcupcUsp;
        public UInt32 lcbPlcupcUsp;
        public UInt32 fcSttbGlsyStyle;
        public UInt32 lcbSttbGlsyStyle;
        public UInt32 fcPlgosl;
        public UInt32 lcbPlgosl;
        public UInt32 fcPlcocx;
        public UInt32 lcbPlcocx;
        public UInt32 fcPlcfBteLvc;
        public UInt32 lcbPlcfBteLvc;
        public UInt32 dwLowDateTime;
        public UInt32 dwHighDateTime;
        public UInt32 fcPlcfLvcPre10;
        public UInt32 lcbPlcfLvcPre10;
        public UInt32 fcPlcfAsumy;
        public UInt32 lcbPlcfAsumy;
        public UInt32 fcPlcfGram;
        public UInt32 lcbPlcfGram;
        public UInt32 fcSttbListNames;
        public UInt32 lcbSttbListNames;
        public UInt32 fcSttbfUssr;
        public UInt32 lcbSttbfUssr;
        #endregion

        #region FibWord2000
        public UInt32 fcPlcfTch;
        public UInt32 lcbPlcfTch;
        public UInt32 fcRmdThreading;
        public UInt32 lcbRmdThreading;
        public UInt32 fcMid;
        public UInt32 lcbMid;
        public UInt32 fcSttbRgtplc;
        public UInt32 lcbSttbRgtplc;
        public UInt32 fcMsoEnvelope;
        public UInt32 lcbMsoEnvelope;
        public UInt32 fcPlcfLad;
        public UInt32 lcbPlcfLad;
        public UInt32 fcRgDofr;
        public UInt32 lcbRgDofr;
        public UInt32 fcPlcosl;
        public UInt32 lcbPlcosl;
        public UInt32 fcPlcfCookieOld;
        public UInt32 lcbPlcfCookieOld;
        public UInt32 fcPgdMotherOld;
        public UInt32 lcbPgdMotherOld;
        public UInt32 fcBkdMotherOld;
        public UInt32 lcbBkdMotherOld;
        public UInt32 fcPgdFtnOld;
        public UInt32 lcbPgdFtnOld;
        public UInt32 fcBkdFtnOld;
        public UInt32 lcbBkdFtnOld;
        public UInt32 fcPgdEdnOld;
        public UInt32 lcbPgdEdnOld;
        public UInt32 fcBkdEdnOld;
        public UInt32 lcbBkdEdnOld;
        #endregion

        #region Fib2002
        public UInt32 fcPlcfPgp;
        public UInt32 lcbPlcfPgp;
        public UInt32 fcPlcfuim;
        public UInt32 lcbPlcfuim;
        public UInt32 fcPlfguidUim;
        public UInt32 lcbPlfguidUim;
        public UInt32 fcAtrdExtra;
        public UInt32 lcbAtrdExtra;
        public UInt32 fcPlrsid;
        public UInt32 lcbPlrsid;
        public UInt32 fcSttbfBkmkFactoid;
        public UInt32 lcbSttbfBkmkFactoid;
        public UInt32 fcPlcfBkfFactoid;
        public UInt32 lcbPlcfBkfFactoid;
        public UInt32 fcPlcfcookie;
        public UInt32 lcbPlcfcookie;
        public UInt32 fcPlcfBklFactoid;
        public UInt32 lcbPlcfBklFactoid;
        public UInt32 fcFactoidData;
        public UInt32 lcbFactoidData;
        public UInt32 fcDocUndo;
        public UInt32 lcbDocUndo;
        public UInt32 fcSttbfBkmkFcc;
        public UInt32 lcbSttbfBkmkFcc;
        public UInt32 fcPlcfBkfFcc;
        public UInt32 lcbPlcfBkfFcc;
        public UInt32 fcPlcfBklFcc;
        public UInt32 lcbPlcfBklFcc;
        public UInt32 fcSttbfbkmkBPRepairs;
        public UInt32 lcbSttbfbkmkBPRepairs;
        public UInt32 fcPlcfbkfBPRepairs;
        public UInt32 lcbPlcfbkfBPRepairs;
        public UInt32 fcPlcfbklBPRepairs;
        public UInt32 lcbPlcfbklBPRepairs;
        public UInt32 fcPmsNew;
        public UInt32 lcbPmsNew;
        public UInt32 fcODSO;
        public UInt32 lcbODSO;
        public UInt32 fcPlcfpmiOldXP;
        public UInt32 lcbPlcfpmiOldXP;
        public UInt32 fcPlcfpmiNewXP;
        public UInt32 lcbPlcfpmiNewXP;
        public UInt32 fcPlcfpmiMixedXP;
        public UInt32 lcbPlcfpmiMixedXP;
        public UInt32 fcPlcffactoid;
        public UInt32 lcbPlcffactoid;
        public UInt32 fcPlcflvcOldXP;
        public UInt32 lcbPlcflvcOldXP;
        public UInt32 fcPlcflvcNewXP;
        public UInt32 lcbPlcflvcNewXP;
        public UInt32 fcPlcflvcMixedXP;
        public UInt32 lcbPlcflvcMixedXP;
        #endregion

        #region Fib2003
        public UInt32 fcHplxsdr;
        public UInt32 lcbHplxsdr;
        public UInt32 fcSttbfBkmkSdt;
        public UInt32 lcbSttbfBkmkSdt;
        public UInt32 fcPlcfBkfSdt;
        public UInt32 lcbPlcfBkfSdt;
        public UInt32 fcPlcfBklSdt;
        public UInt32 lcbPlcfBklSdt;
        public UInt32 fcCustomXForm;
        public UInt32 lcbCustomXForm;
        public UInt32 fcSttbfBkmkProt;
        public UInt32 lcbSttbfBkmkProt;
        public UInt32 fcPlcfBkfProt;
        public UInt32 lcbPlcfBkfProt;
        public UInt32 fcPlcfBklProt;
        public UInt32 lcbPlcfBklProt;
        public UInt32 fcSttbProtUser;
        public UInt32 lcbSttbProtUser;
        public UInt32 fcPlcfpmiOld;
        public UInt32 lcbPlcfpmiOld;
        public UInt32 fcPlcfpmiOldInline;
        public UInt32 lcbPlcfpmiOldInline;
        public UInt32 fcPlcfpmiNew;
        public UInt32 lcbPlcfpmiNew;
        public UInt32 fcPlcfpmiNewInline;
        public UInt32 lcbPlcfpmiNewInline;
        public UInt32 fcPlcflvcOld;
        public UInt32 lcbPlcflvcOld;
        public UInt32 fcPlcflvcOldInline;
        public UInt32 lcbPlcflvcOldInline;
        public UInt32 fcPlcflvcNew;
        public UInt32 lcbPlcflvcNew;
        public UInt32 fcPlcflvcNewInline;
        public UInt32 lcbPlcflvcNewInline;
        public UInt32 fcPgdMother;
        public UInt32 lcbPgdMother;
        public UInt32 fcBkdMother;
        public UInt32 lcbBkdMother;
        public UInt32 fcAfdMother;
        public UInt32 lcbAfdMother;
        public UInt32 fcPgdFtn;
        public UInt32 lcbPgdFtn;
        public UInt32 fcBkdFtn;
        public UInt32 lcbBkdFtn;
        public UInt32 fcAfdFtn;
        public UInt32 lcbAfdFtn;
        public UInt32 fcPgdEdn;
        public UInt32 lcbPgdEdn;
        public UInt32 fcBkdEdn;
        public UInt32 lcbBkdEdn;
        public UInt32 fcAfdEdn;
        public UInt32 lcbAfdEdn;
        public UInt32 fcAfd;
        public UInt32 lcbAfd;
        #endregion

        #region Fib2007
        public UInt32 fcPlcfmthd;
        public UInt32 lcbPlcfmthd;
        public UInt32 fcSttbfBkmkMoveFrom;
        public UInt32 lcbSttbfBkmkMoveFrom;
        public UInt32 fcPlcfBkfMoveFrom;
        public UInt32 lcbPlcfBkfMoveFrom;
        public UInt32 fcPlcfBklMoveFrom;
        public UInt32 lcbPlcfBklMoveFrom;
        public UInt32 fcSttbfBkmkMoveTo;
        public UInt32 lcbSttbfBkmkMoveTo;
        public UInt32 fcPlcfBkfMoveTo;
        public UInt32 lcbPlcfBkfMoveTo;
        public UInt32 fcPlcfBklMoveTo;
        public UInt32 lcbPlcfBklMoveTo;
        public UInt32 fcSttbfBkmkArto;
        public UInt32 lcbSttbfBkmkArto;
        public UInt32 fcPlcfBkfArto;
        public UInt32 lcbPlcfBkfArto;
        public UInt32 fcPlcfBklArto;
        public UInt32 lcbPlcfBklArto;
        public UInt32 fcArtoData;
        public UInt32 lcbArtoData;
        public UInt32 fcOssTheme;
        public UInt32 lcbOssTheme;
        public UInt32 fcColorSchemeMapping;
        public UInt32 lcbColorSchemeMapping;
        #endregion

        #region FibNew
        public FibVersion nFibNew;
        public UInt16 cQuickSavesNew;
        #endregion

        #region others
        public UInt16 csw;
        public UInt16 cslw;
        public UInt16 cbRgFcLcb;
        public UInt16 cswNew;
        #endregion

        //*****************************************************************************************
        //                                                                              CONSTRUCTOR
        //*****************************************************************************************

        public FileInformationBlock(VirtualStreamReader reader)
        {
            UInt16 flag16 = 0;
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
            this.cQuickSaves = (UInt16)(((int)flag16 & 0x00F0) >> 4);
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
