/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *       notice, this list of conditions and the following disclaimer.
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
using System.Xml;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.StyleData;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.SpreadsheetMLMapping
{
    public class StylesMapping : AbstractOpenXmlMapping,
          IMapping<StyleData>
    {
        ExcelContext xlsContext;

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="xlsContext">The excel context object</param>
        public StylesMapping(ExcelContext xlsContext)
            : base(XmlWriter.Create(xlsContext.SpreadDoc.WorkbookPart.AddStylesPart().GetStream(), xlsContext.WriterSettings))
        {
            this.xlsContext = xlsContext;
        }

        /// <summary>
        /// The overload apply method 
        /// Creates the Styles xml document 
        /// </summary>
        /// <param name="sd">StyleData Object</param>
        public void Apply(StyleData sd)
        {
            _writer.WriteStartDocument();
            _writer.WriteStartElement("styleSheet", OpenXmlNamespaces.SpreadsheetML);


            // Format mapping 
            _writer.WriteStartElement("numFmts");
            _writer.WriteAttributeString("count", sd.FormatDataList.Count.ToString());
            foreach (FormatData format in sd.FormatDataList)
            {
                _writer.WriteStartElement("numFmt");
                _writer.WriteAttributeString("numFmtId", format.ifmt.ToString());
                _writer.WriteAttributeString("formatCode", format.formatString);
                _writer.WriteEndElement();
            }
            _writer.WriteEndElement(); 

             


            /// Font Mapping
            //<fonts count="1">
            //<font>
            //<sz val="10"/>
            //<name val="Arial"/>
            //</font>
            //</fonts>
            _writer.WriteStartElement("fonts");
            _writer.WriteAttributeString("count", sd.FontDataList.Count.ToString());
            foreach (FontData font in sd.FontDataList)
            {
                ///
                StyleMappingHelper.addFontElement(_writer, font, FontElementType.NormalStyle); 


            }
            // write fonts end element 
            _writer.WriteEndElement(); 
            
            /// Fill Mapping 
            //<fills count="2">
            //<fill>
            //<patternFill patternType="none"/>
            //</fill>           
            _writer.WriteStartElement("fills");
            _writer.WriteAttributeString("count", sd.FillDataList.Count.ToString());
            foreach (FillData fd in sd.FillDataList)
            {
                _writer.WriteStartElement("fill");
                _writer.WriteStartElement("patternFill");
                _writer.WriteAttributeString("patternType", StyleMappingHelper.getStringFromFillPatern(fd.Fillpatern));

                // foreground color 
                WriteRgbForegroundColor(_writer, StyleMappingHelper.convertColorIdToRGB(fd.IcvFore)); 

                // background color 
                WriteRgbBackgroundColor(_writer, StyleMappingHelper.convertColorIdToRGB(fd.IcvBack));

                _writer.WriteEndElement();
                _writer.WriteEndElement();
            }

            _writer.WriteEndElement();


            /// Border Mapping 
            //<borders count="1">
            //  <border>
            //      <left/>
            //      <right/>
            //      <top/>
            //      <bottom/>
            //      <diagonal/>
            //  </border>
            //</borders>
            _writer.WriteStartElement("borders");
            _writer.WriteAttributeString("count", sd.BorderDataList.Count.ToString());
            foreach (BorderData borderData in sd.BorderDataList)
            {
                _writer.WriteStartElement("border");

                // write diagonal settings 
                if (borderData.diagonalValue == 1)
                {
                    _writer.WriteAttributeString("diagonalDown", "1");
                }
                else if (borderData.diagonalValue == 2)
                {
                    _writer.WriteAttributeString("diagonalUp", "1");
                }
                else if (borderData.diagonalValue == 3)
                {
                    _writer.WriteAttributeString("diagonalDown", "1");
                    _writer.WriteAttributeString("diagonalUp", "1");
                }
                else
                {
                    // do nothing !
                }

               
                string borderStyle = ""; 

                // left border 
                _writer.WriteStartElement("left");
                borderStyle = StyleMappingHelper.convertBorderStyle(borderData.left.style); 
                if (!borderStyle.Equals("none"))
                {
                    _writer.WriteAttributeString("style", borderStyle);
                    WriteRgbColor(_writer, StyleMappingHelper.convertColorIdToRGB(borderData.left.colorId));
                }
                _writer.WriteEndElement();

                // right border 
                _writer.WriteStartElement("right");
                borderStyle = StyleMappingHelper.convertBorderStyle(borderData.right.style);
                if (!borderStyle.Equals("none"))
                {
                    _writer.WriteAttributeString("style", borderStyle);
                    WriteRgbColor(_writer, StyleMappingHelper.convertColorIdToRGB(borderData.right.colorId));
                }
                _writer.WriteEndElement();

                // top border 
                _writer.WriteStartElement("top");
                borderStyle = StyleMappingHelper.convertBorderStyle(borderData.top.style);
                if (!borderStyle.Equals("none"))
                {
                    _writer.WriteAttributeString("style", borderStyle);
                    WriteRgbColor(_writer, StyleMappingHelper.convertColorIdToRGB(borderData.top.colorId));
                }
                _writer.WriteEndElement();

                // bottom border 
                _writer.WriteStartElement("bottom");
                borderStyle = StyleMappingHelper.convertBorderStyle(borderData.bottom.style);
                if (!borderStyle.Equals("none"))
                {
                    _writer.WriteAttributeString("style", borderStyle);
                    WriteRgbColor(_writer, StyleMappingHelper.convertColorIdToRGB(borderData.bottom.colorId));
                }
                _writer.WriteEndElement();

                // diagonal border 
                _writer.WriteStartElement("diagonal");
                borderStyle = StyleMappingHelper.convertBorderStyle(borderData.diagonal.style);
                if (!borderStyle.Equals("none"))
                {
                    _writer.WriteAttributeString("style", borderStyle);
                    WriteRgbColor(_writer, StyleMappingHelper.convertColorIdToRGB(borderData.diagonal.colorId));
                }
                _writer.WriteEndElement();

                _writer.WriteEndElement(); // end border 
            }
            _writer.WriteEndElement(); // end borders 

            ///<cellStyleXfs count="1">
            ///<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
            ///</cellStyleXfs> 
            // xfcellstyle mapping 
            _writer.WriteStartElement("cellStyleXfs");
            _writer.WriteAttributeString("count", sd.XFCellStyleDataList.Count.ToString());
            foreach (XFData xfcellstyle in sd.XFCellStyleDataList)
            {
                _writer.WriteStartElement("xf");
                _writer.WriteAttributeString("numFmtId", xfcellstyle.ifmt.ToString());
                _writer.WriteAttributeString("fontId", xfcellstyle.fontId.ToString());
                _writer.WriteAttributeString("fillId", xfcellstyle.fillId.ToString());
                _writer.WriteAttributeString("borderId", xfcellstyle.borderId.ToString());

                if (xfcellstyle.hasAlignment)
                {
                    StylesMapping.WriteCellAlignment(_writer, xfcellstyle);
                }

                _writer.WriteEndElement();
            }

            _writer.WriteEndElement();




            ///<cellXfs count="6">
            ///<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
            // xfcell mapping 
            _writer.WriteStartElement("cellXfs");
            _writer.WriteAttributeString("count", sd.XFCellDataList.Count.ToString());
            foreach (XFData xfcell in sd.XFCellDataList)
            {
                _writer.WriteStartElement("xf");
                _writer.WriteAttributeString("numFmtId", xfcell.ifmt.ToString());
                _writer.WriteAttributeString("fontId", xfcell.fontId.ToString());
                _writer.WriteAttributeString("fillId", xfcell.fillId.ToString());
                _writer.WriteAttributeString("borderId", xfcell.borderId.ToString());
                _writer.WriteAttributeString("xfId", xfcell.ixfParent.ToString());

                // applyNumberFormat="1"
                if (xfcell.ifmt != 0)
                {
                    _writer.WriteAttributeString("applyNumberFormat", "1");
                }

                // applyBorder="1"
                if (xfcell.borderId != 0)
                {
                    _writer.WriteAttributeString("applyBorder", "1");
                }

                // applyFill="1"
                if (xfcell.fillId != 0)
                {
                    _writer.WriteAttributeString("applyFill", "1");
                }

                // applyFont="1"
                if (xfcell.fontId != 0)
                {
                    _writer.WriteAttributeString("applyFont", "1");
                }
                if (xfcell.hasAlignment)
                {
                    StylesMapping.WriteCellAlignment(_writer, xfcell); 
                }

                _writer.WriteEndElement();
            }

            _writer.WriteEndElement(); 




            /// write cell styles 
            /// <cellStyles count="1">
            /// <cellStyle name="Normal" xfId="0" builtinId="0"/>
            /// </cellStyles>
            /// 
            _writer.WriteStartElement("cellStyles");
            //_writer.WriteAttributeString("count", sd.StyleList.Count.ToString());
            foreach (Style style in sd.StyleList)
            {
                _writer.WriteStartElement("cellStyle");

                if (style.rgch != null)
                { 
                    _writer.WriteAttributeString("name", style.rgch); 
                }
                // theres a bug with the zero based reading from the referenz id 
                // so the style.ixfe value is reduzed by one
                if (style.ixfe != 0)
                {
                    _writer.WriteAttributeString("xfId", (style.ixfe - 1).ToString());
                }
                else
                {
                    _writer.WriteAttributeString("xfId", (style.ixfe).ToString());
                }
                _writer.WriteAttributeString("builtinId", style.istyBuiltIn.ToString());

                _writer.WriteEndElement();
            }

            _writer.WriteEndElement(); 
            
            // close tags 


            // write color table !!

            if (sd.ColorDataList != null && sd.ColorDataList.Count > 0)
            {
                _writer.WriteStartElement("colors");

                _writer.WriteStartElement("indexedColors");

                // <rgbColor rgb="00000000"/>
                foreach (RGBColor item in sd.ColorDataList)
                {
                    _writer.WriteStartElement("rgbColor");
                    _writer.WriteAttributeString("rgb", String.Format("{0:x2}", item.Alpha).ToString() + item.SixDigitHexCode); 

                    _writer.WriteEndElement(); 

                }


                _writer.WriteEndElement(); 
                _writer.WriteEndElement();
            }
            // end color 

            _writer.WriteEndElement();      // close 
            _writer.WriteEndDocument();

            // close writer 
            _writer.Flush();
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="writer"></param>
        /// <param name="color"></param>

        public static void WriteRgbColor(XmlWriter writer, string color)
        {
            if (!String.IsNullOrEmpty(color) && color != "Auto")
            {
                writer.WriteStartElement("color");
                
                writer.WriteAttributeString("rgb", "FF" + color);
                writer.WriteEndElement();
            }
        }

        // <color indexed="63"/>

        public static void WriteRgbForegroundColor(XmlWriter writer, string color)
        {
            if (!String.IsNullOrEmpty(color) && color != "Auto")
            {
                writer.WriteStartElement("fgColor");
                writer.WriteAttributeString("rgb", "FF" + color);
                writer.WriteEndElement();
            }
        }

        public static void WriteRgbBackgroundColor(XmlWriter writer, string color)
        {
            if (!String.IsNullOrEmpty(color) && color != "Auto")
            {
                writer.WriteStartElement("bgColor");
                writer.WriteAttributeString("rgb", "FF" + color);
                writer.WriteEndElement();
            }
        }

        public static void WriteCellAlignment(XmlWriter _writer, XFData xfcell)
        {
            _writer.WriteStartElement("alignment");
            if (xfcell.wrapText)
            {
                _writer.WriteAttributeString("wrapText", "1");
            }
            if (xfcell.horizontalAlignment != 0xFF)
            {
                _writer.WriteAttributeString("horizontal", StyleMappingHelper.getHorAlignmentValue(xfcell.horizontalAlignment));
            }
            if (xfcell.verticalAlignment != 0x02)
            {
                _writer.WriteAttributeString("vertical", StyleMappingHelper.getVerAlignmentValue(xfcell.verticalAlignment));
            }
            if (xfcell.justifyLastLine)
            {
                _writer.WriteAttributeString("justifyLastLine", "1");
            }
            if (xfcell.shrinkToFit)
            {
                _writer.WriteAttributeString("shrinkToFit", "1");
            }
            if (xfcell.textRotation != 0x00)
            {
                _writer.WriteAttributeString("textRotation", xfcell.textRotation.ToString());
            }
            if (xfcell.indent != 0x00)
            {
                _writer.WriteAttributeString("indent", xfcell.indent.ToString());
            }
            if (xfcell.readingOrder != 0x00)
            {
                _writer.WriteAttributeString("readingOrder", xfcell.readingOrder.ToString());
            }

            _writer.WriteEndElement();
        }
    }
}