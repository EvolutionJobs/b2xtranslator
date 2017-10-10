/*
 * Copyright (c) 2009, DIaLOGIKa
 *
 * All rights reserved.
 * 
 * Redistribution and use in source and binary forms, with or without 
 * modification, are permitted provided that the following conditions are met:
 * 
 *     * Redistributions of source code must retain the above copyright 
 *       notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright 
 *       notice, this list of conditions and the following disclaimer in the 
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the names of copyright holders, nor the names of its contributors 
 *       may be used to endorse or promote products derived from this software 
 *       without specific prior written permission.
 * 
 * THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" 
 * AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED 
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. 
 * IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, 
 * INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, 
 * BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, 
 * DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF 
 * LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE 
 * OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF 
 * ADVISED OF THE POSSIBILITY OF SUCH DAMAGE. 
 */

using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.DrawingML;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.StyleData;
using System.Globalization;

namespace DIaLOGIKa.b2xtranslator.SpreadsheetMLMapping
{
    public class TextBodyMapping : AbstractChartMapping,
          IMapping<AttachedLabelSequence>
    {
        public TextBodyMapping(ExcelContext workbookContext, ChartContext chartContext)
            : base(workbookContext, chartContext)
        {
        }

        #region IMapping<AttachedLabelSequence> Members

        public void Apply(AttachedLabelSequence attachedLabelSequence)
        {
            // c:txPr
            _writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElTxPr, Dml.Chart.Ns);
            {
                // a:bodyPr (is empty for legends)
                _writer.WriteElementString(Dml.Prefix, Dml.Text.ElBodyPr, Dml.Ns, string.Empty);

                // a:lstStyle (is empty for legends)
                _writer.WriteElementString(Dml.Prefix, Dml.Text.ElLstStyle, Dml.Ns, string.Empty);

                // a:p
                _writer.WriteStartElement(Dml.Prefix, Dml.Text.ElP, Dml.Ns);
                {
                    // a:pPr
                    _writer.WriteStartElement(Dml.Prefix, Dml.Text.ElPPr, Dml.Ns);
                    {
                        int fontIndex = 0;

                        if (attachedLabelSequence.FontX != null && attachedLabelSequence.FontX.iFont <= this.WorkbookContext.XlsDoc.WorkBookData.styleData.FontDataList.Count)
                        {
                            // FontX.iFont is a 1-based index
                            fontIndex = attachedLabelSequence.FontX.iFont - 1;
                        }
                        var fontData = this.WorkbookContext.XlsDoc.WorkBookData.styleData.FontDataList[fontIndex];

                        // a:defRPr
                        _writer.WriteStartElement(Dml.Prefix, Dml.TextParagraph.ElDefRPr, Dml.Ns);

                        //_writer.WriteAttributeString(Dml.TextCharacter.AttrKumimoji, );
                        //_writer.WriteAttributeString(Dml.TextCharacter.AttrLang,  );
                        //_writer.WriteAttributeString(Dml.TextCharacter.AttrAltLang,  );
                        _writer.WriteAttributeString(Dml.TextCharacter.AttrSz, (fontData.size.ToPoints() * 100).ToString(CultureInfo.InvariantCulture));
                        _writer.WriteAttributeString(Dml.TextCharacter.AttrB, fontData.isBold ? "1" : "0");
                        _writer.WriteAttributeString(Dml.TextCharacter.AttrI, fontData.isItalic ? "1" : "0");
                        _writer.WriteAttributeString(Dml.TextCharacter.AttrU, mapUnderlineStyle(fontData.uStyle));
                        _writer.WriteAttributeString(Dml.TextCharacter.AttrStrike, fontData.isStrike ? "sngStrike" : "noStrike");
                        //_writer.WriteAttributeString(Dml.TextCharacter.AttrKern,  );
                        //_writer.WriteAttributeString(Dml.TextCharacter.AttrCap,  );
                        //_writer.WriteAttributeString(Dml.TextCharacter.AttrSpc,  );
                        //_writer.WriteAttributeString(Dml.TextCharacter.AttrNormalizeH,  );
                        if (fontData.vertAlign != SuperSubScriptStyle.none)
                        {
                            _writer.WriteAttributeString(Dml.TextCharacter.AttrBaseline, fontData.vertAlign == SuperSubScriptStyle.superscript ? "30000" : "-25000");
                        }
                        //_writer.WriteAttributeString(Dml.TextCharacter.AttrNoProof,  );
                        //_writer.WriteAttributeString(Dml.TextCharacter.AttrDirty,  );
                        //_writer.WriteAttributeString(Dml.TextCharacter.AttrErr,  );
                        //_writer.WriteAttributeString(Dml.TextCharacter.AttrSmtClean,  );
                        //_writer.WriteAttributeString(Dml.TextCharacter.AttrSmtId,  );
                        //_writer.WriteAttributeString(Dml.TextCharacter.AttrBmk,  );
                        {
                            // a:latin
                            _writer.WriteStartElement(Dml.Prefix, Dml.TextCharacter.ElLatin, Dml.Ns);
                            _writer.WriteAttributeString(Dml.TextCharacter.AttrTypeface, fontData.fontName);
                            _writer.WriteEndElement(); // a:latin

                            // a:ea
                            _writer.WriteStartElement(Dml.Prefix, Dml.TextCharacter.ElEa, Dml.Ns);
                            _writer.WriteAttributeString(Dml.TextCharacter.AttrTypeface, fontData.fontName);
                            _writer.WriteEndElement(); // a:ea

                            // a:cs
                            _writer.WriteStartElement(Dml.Prefix, Dml.TextCharacter.ElCs, Dml.Ns);
                            _writer.WriteAttributeString(Dml.TextCharacter.AttrTypeface, fontData.fontName);
                            _writer.WriteEndElement(); // a:cs
                        }
                        _writer.WriteEndElement(); // a:defRPr
                    }
                    _writer.WriteEndElement(); // a:pPr
                }
                _writer.WriteEndElement(); // a:p
            }
            _writer.WriteEndElement(); // c:txPr
        }
        #endregion

        private string mapUnderlineStyle(UnderlineStyle uStyle)
        {
            // TODO: map "accounting" variants
            switch (uStyle)
            {
                case UnderlineStyle.none:
                    return "none";
                case UnderlineStyle.singleLine:
                    return "sng";
                case UnderlineStyle.singleAccounting:
                    return "sng";
                case UnderlineStyle.doubleLine:
                    return "dbl";
                case UnderlineStyle.doubleAccounting:
                    return "dbl";
                default:
                    return "none";
            }
        }
    }
}
