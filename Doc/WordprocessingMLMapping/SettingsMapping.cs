using System;
using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.DocFileFormat;
using System.Xml;
using b2xtranslator.OpenXmlLib;

namespace b2xtranslator.WordprocessingMLMapping
{
    public class SettingsMapping : PropertiesMapping,
          IMapping<DocumentProperties>
    {
        private enum FootnotePosition
        {
            beneathText,
            docEnd,
            pageBottom,
            sectEnd
        }

        private enum ZoomType
        {
            none,
            fullPage,
            bestFit
        }

        private ConversionContext _ctx;

        public SettingsMapping(ConversionContext ctx, OpenXmlPart targetPart)
            : base(XmlWriter.Create(targetPart.GetStream(), ctx.WriterSettings))
        {
            this._ctx = ctx;
        }

        public void Apply(DocumentProperties dop)
        {
            //start w:settings
            this._writer.WriteStartElement("w", "settings", OpenXmlNamespaces.WordprocessingML);

            //zoom
            this._writer.WriteStartElement("w", "zoom", OpenXmlNamespaces.WordprocessingML);
            this._writer.WriteAttributeString("w", "percent", OpenXmlNamespaces.WordprocessingML, dop.wScaleSaved.ToString());
            var zoom = (ZoomType)dop.zkSaved;
            if (zoom != ZoomType.none)
            {
                this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, zoom.ToString());
            }
            this._writer.WriteEndElement();

            //doc protection
            //<w:documentProtection w:edit="forms" w:enforcement="1"/>
           

            //embed system fonts
            if(!dop.fDoNotEmbedSystemFont)
            {
                this._writer.WriteElementString("w", "embedSystemFonts", OpenXmlNamespaces.WordprocessingML, "");
            }

            //mirror margins
            if (dop.fMirrorMargins)
            {
                this._writer.WriteElementString("w", "mirrorMargins", OpenXmlNamespaces.WordprocessingML, "");
            }

            //evenAndOddHeaders 
            if (dop.fFacingPages)
            {
                this._writer.WriteElementString("w", "evenAndOddHeaders", OpenXmlNamespaces.WordprocessingML, "");
            }

            //proof state
            var proofState = this._nodeFactory.CreateElement("w", "proofState", OpenXmlNamespaces.WordprocessingML);
            if (dop.fGramAllClean)
                appendValueAttribute(proofState, "grammar", "clean");
            if (proofState.Attributes.Count > 0)
                proofState.WriteTo(this._writer);

            //stylePaneFormatFilter
            if (dop.grfFmtFilter != 0)
            {
                this._writer.WriteStartElement("w", "stylePaneFormatFilter", OpenXmlNamespaces.WordprocessingML);
                this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, string.Format("{0:x4}", dop.grfFmtFilter));
                this._writer.WriteEndElement();
            }

            //default tab stop
            this._writer.WriteStartElement("w", "defaultTabStop", OpenXmlNamespaces.WordprocessingML);
            this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, dop.dxaTab.ToString());
            this._writer.WriteEndElement();

            //drawing grid
            if(dop.dogrid != null)
            {
                this._writer.WriteStartElement("w", "displayHorizontalDrawingGridEvery", OpenXmlNamespaces.WordprocessingML);
                this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, dop.dogrid.dxGridDisplay.ToString());
                this._writer.WriteEndElement();

                this._writer.WriteStartElement("w", "displayVerticalDrawingGridEvery", OpenXmlNamespaces.WordprocessingML);
                this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, dop.dogrid.dyGridDisplay.ToString());
                this._writer.WriteEndElement();

                if (dop.dogrid.fFollowMargins == false)
                {
                    this._writer.WriteElementString("w", "doNotUseMarginsForDrawingGridOrigin", OpenXmlNamespaces.WordprocessingML, "");
                }
            }

            //typography
            if (dop.doptypography != null)
            {
                if (dop.doptypography.fKerningPunct == false)
                {
                    this._writer.WriteElementString("w", "noPunctuationKerning", OpenXmlNamespaces.WordprocessingML, "");
                }
            }

            //footnote properties
            var footnotePr = this._nodeFactory.CreateElement("w", "footnotePr", OpenXmlNamespaces.WordprocessingML);
            if (dop.nFtn != 0)
                appendValueAttribute(footnotePr, "numStart", dop.nFtn.ToString());
            if (dop.rncFtn != 0)
                appendValueAttribute(footnotePr, "numRestart", dop.rncFtn.ToString());
            if (dop.Fpc != 0)
                appendValueAttribute(footnotePr, "pos", ((FootnotePosition)dop.Fpc).ToString());
            if (footnotePr.Attributes.Count > 0)
                footnotePr.WriteTo(this._writer);


            writeCompatibilitySettings(dop);

            writeRsidList();

            //close w:settings
            this._writer.WriteEndElement();

            this._writer.Flush();
        }

        private void writeRsidList()
        {
            //convert the rsid list
            this._writer.WriteStartElement("w", "rsids", OpenXmlNamespaces.WordprocessingML);
            this._ctx.AllRsids.Sort();
            foreach (string rsid in this._ctx.AllRsids)
            {
                this._writer.WriteStartElement("w", "rsid", OpenXmlNamespaces.WordprocessingML);
                this._writer.WriteAttributeString("w", "val", OpenXmlNamespaces.WordprocessingML, rsid);
                this._writer.WriteEndElement();
            }
            this._writer.WriteEndElement();
        }

        private void writeCompatibilitySettings(DocumentProperties dop)
        {
            //compatibility settings
            this._writer.WriteStartElement("w", "compat", OpenXmlNamespaces.WordprocessingML);

            //some settings must always be written
            this._writer.WriteElementString("w", "useNormalStyleForList", OpenXmlNamespaces.WordprocessingML, "");
            this._writer.WriteElementString("w", "doNotUseIndentAsNumberingTabStop", OpenXmlNamespaces.WordprocessingML, "");
            this._writer.WriteElementString("w", "useAltKinsokuLineBreakRules", OpenXmlNamespaces.WordprocessingML, "");
            this._writer.WriteElementString("w", "allowSpaceOfSameStyleInTable", OpenXmlNamespaces.WordprocessingML, "");
            this._writer.WriteElementString("w", "doNotSuppressIndentation", OpenXmlNamespaces.WordprocessingML, "");
            this._writer.WriteElementString("w", "doNotAutofitConstrainedTables", OpenXmlNamespaces.WordprocessingML, "");
            this._writer.WriteElementString("w", "autofitToFirstFixedWidthCell", OpenXmlNamespaces.WordprocessingML, "");
            this._writer.WriteElementString("w", "displayHangulFixedWidth", OpenXmlNamespaces.WordprocessingML, "");
            this._writer.WriteElementString("w", "splitPgBreakAndParaMark", OpenXmlNamespaces.WordprocessingML, "");
            this._writer.WriteElementString("w", "doNotVertAlignCellWithSp", OpenXmlNamespaces.WordprocessingML, "");
            this._writer.WriteElementString("w", "doNotBreakConstrainedForcedTable", OpenXmlNamespaces.WordprocessingML, "");
            this._writer.WriteElementString("w", "doNotVertAlignInTxbx", OpenXmlNamespaces.WordprocessingML, "");
            this._writer.WriteElementString("w", "useAnsiKerningPairs", OpenXmlNamespaces.WordprocessingML, "");
            this._writer.WriteElementString("w", "cachedColBalance", OpenXmlNamespaces.WordprocessingML, "");
            
            //others are saved in the file
            if(!dop.fDontAdjustLineHeightInTable)
                this._writer.WriteElementString("w", "adjustLineHeightInTable", OpenXmlNamespaces.WordprocessingML, "");
            if (dop.fAlignTablesRowByRow)
                this._writer.WriteElementString("w", "alignTablesRowByRow", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fApplyBreakingRules)
                this._writer.WriteElementString("w", "applyBreakingRules", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fUseAutoSpaceForFullWidthAlpha)
                this._writer.WriteElementString("w", "autoSpaceLikeWord95", OpenXmlNamespaces.WordprocessingML, "");
            if(!dop.fDntBlnSbDbWid)
                this._writer.WriteElementString("w", "balanceSingleByteDoubleByteWidth", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fConvMailMergeEsc)
                this._writer.WriteElementString("w", "convMailMergeEsc", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fDontBreakWrappedTables)
                this._writer.WriteElementString("w", "doNotBreakWrappedTables", OpenXmlNamespaces.WordprocessingML, "");
            if(!dop.fExpShRtn)
                this._writer.WriteElementString("w", "doNotExpandShiftReturn", OpenXmlNamespaces.WordprocessingML, "");
            if(!dop.fLeaveBackslashAlone)
                this._writer.WriteElementString("w", "doNotLeaveBackslashAlone", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fDontSnapToGridInCell)
                this._writer.WriteElementString("w", "doNotSnapToGridInCell", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fDontUseAsianBreakRules)
                this._writer.WriteElementString("w", "doNotUseEastAsianBreakRules", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fDontUseHTMLParagraphAutoSpacing)
                this._writer.WriteElementString("w", "doNotUseHTMLParagraphAutoSpacing", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fDontWrapTextWithPunct)
                this._writer.WriteElementString("w", "doNotWrapTextWithPunct", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fFtnLayoutLikeWW8)
                this._writer.WriteElementString("w", "footnoteLayoutLikeWW8", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fForgetLastTabAlign)
                this._writer.WriteElementString("w", "forgetLastTabAlignment", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fGrowAutofit)
                this._writer.WriteElementString("w", "growAutofit", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fLayoutRawTableWidth)
                this._writer.WriteElementString("w", "layoutRawTableWidth", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fLayoutTableRowsApart)
                this._writer.WriteElementString("w", "layoutTableRowsApart", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fLineWrapLikeWord6)
                this._writer.WriteElementString("w", "lineWrapLikeWord6", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fMWSmallCaps)
                this._writer.WriteElementString("w", "mwSmallCaps", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fNoColumnBalance)
                this._writer.WriteElementString("w", "noColumnBalance", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fNoLeading)
                this._writer.WriteElementString("w", "noLeading", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fNoSpaceRaiseLower)
                this._writer.WriteElementString("w", "noSpaceRaiseLower", OpenXmlNamespaces.WordprocessingML, "");
            if (dop.fNoTabForInd)
                this._writer.WriteElementString("w", "noTabHangInd", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fPrintBodyBeforeHdr)
                this._writer.WriteElementString("w", "printBodyTextBeforeHeader", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fMapPrintTextColor)
                this._writer.WriteElementString("w", "printColBlack", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fDontAllowFieldEndSelect)
                this._writer.WriteElementString("w", "selectFldWithFirstOrLastChar", OpenXmlNamespaces.WordprocessingML, "");
            if (dop.fSpLayoutLikeWW8)
                this._writer.WriteElementString("w", "shapeLayoutLikeWW8", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fShowBreaksInFrames)
                this._writer.WriteElementString("w", "showBreaksInFrames", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fMakeSpaceForUL)
                this._writer.WriteElementString("w", "spaceForUL", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fTruncDxaExpand)
                this._writer.WriteElementString("w", "spacingInWholePoints", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fSubOnSize)
                this._writer.WriteElementString("w", "subFontBySize", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fSuppressSpbfAfterPageBreak)
                this._writer.WriteElementString("w", "suppressSpBfAfterPgBrk", OpenXmlNamespaces.WordprocessingML, "");
            if (dop.fSuppressTopSpacing)
                this._writer.WriteElementString("w", "suppressTopSpacing", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fSwapBordersFacingPgs)
                this._writer.WriteElementString("w", "swapBordersFacingPages", OpenXmlNamespaces.WordprocessingML, "");
            if(!dop.fDntULTrlSpc)
                this._writer.WriteElementString("w", "ulTrailSpace", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fPrintMet)
                this._writer.WriteElementString("w", "usePrinterMetrics", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fUseWord2002TableStyleRules)
                this._writer.WriteElementString("w", "useWord2002TableStyleRules", OpenXmlNamespaces.WordprocessingML, "");
            if (dop.fUserWord97LineBreakingRules)
                this._writer.WriteElementString("w", "useWord97LineBreakRules", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fWPJust)
                this._writer.WriteElementString("w", "wpJustification", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fWPSpace)
                this._writer.WriteElementString("w", "wpSpaceWidth", OpenXmlNamespaces.WordprocessingML, "");
            if(dop.fWrapTrailSpaces)
                this._writer.WriteElementString("w", "wrapTrailSpaces", OpenXmlNamespaces.WordprocessingML, "");

            this._writer.WriteEndElement();
        }
    }
}
