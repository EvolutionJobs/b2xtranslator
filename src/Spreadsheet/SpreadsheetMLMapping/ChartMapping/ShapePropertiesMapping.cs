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
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using System;
using System.Globalization;
using DIaLOGIKa.b2xtranslator.Tools;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using System.Collections.Generic;

namespace DIaLOGIKa.b2xtranslator.SpreadsheetMLMapping
{
    /// <summary>
    /// A class for generating shape properties as defined by CT_ShapeProperties
    /// 
    ///     <xsd:complexType name="CT_ShapeProperties">
    ///         <xsd:sequence>
    ///             <xsd:element name="xfrm" type="CT_Transform2D" minOccurs="0" maxOccurs="1">
    ///                 <xsd:annotation>
    ///                     <xsd:documentation>2D Transform for Individual Objects</xsd:documentation>
    ///                 </xsd:annotation>
    ///             </xsd:element>
    ///             <xsd:group ref="EG_Geometry" minOccurs="0" maxOccurs="1" />
    ///             <xsd:group ref="EG_FillProperties" minOccurs="0" maxOccurs="1" />
    ///             <xsd:element name="ln" type="CT_LineProperties" minOccurs="0" maxOccurs="1" />
    ///             <xsd:group ref="EG_EffectProperties" minOccurs="0" maxOccurs="1" />
    ///             <xsd:element name="scene3d" type="CT_Scene3D" minOccurs="0" maxOccurs="1" />      
    ///             <xsd:element name="sp3d" type="CT_Shape3D" minOccurs="0" maxOccurs="1" />
    ///             <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"></xsd:element>
    ///         </xsd:sequence>
    ///         <xsd:attribute name="bwMode" type="ST_BlackWhiteMode" use="optional">
    ///             <xsd:annotation>
    ///                 <xsd:documentation>Black and White Mode</xsd:documentation>
    ///             </xsd:annotation>
    ///         </xsd:attribute>
    ///     </xsd:complexType>
    /// </summary>
    public class ShapePropertiesMapping : AbstractChartMapping,
        IMapping<SsSequence>,
        IMapping<FrameSequence>
    {
        public ShapePropertiesMapping(ExcelContext workbookContext, ChartContext chartContext)
            : base(workbookContext, chartContext)
        {
        }


        #region IMapping<SsSequence> Members

        public void Apply(SsSequence ssSequence)
        {
            // CT_ShapeProperties

            // c:spPr
            _writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElSpPr, Dml.Chart.Ns);
            {
                // a:xfrm

                // EG_Geometry

                // EG_FillProperties
                insertFillProperties(ssSequence.AreaFormat, ssSequence.GelFrameSequence);

                // a:ln
                insertLineProperties(ssSequence.LineFormat, ssSequence.GelFrameSequence);

                // EG_EffectProperties

                // a:scene3d

                // a:sp3d
            }
            _writer.WriteEndElement(); // c:spPr
        }

        #endregion

        #region IMapping<FrameSequence> Members

        public void Apply(FrameSequence frameSequence)
        {
            // CT_ShapeProperties

            // c:spPr
            _writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElSpPr, Dml.Chart.Ns);
            {
                // a:xfrm

                // EG_Geometry

                // EG_FillProperties
                insertFillProperties(frameSequence.AreaFormat, frameSequence.GelFrameSequence);

                // a:ln
                insertLineProperties(frameSequence.LineFormat, frameSequence.GelFrameSequence);

                // EG_EffectProperties

                // a:scene3d

                // a:sp3d
            }
            _writer.WriteEndElement(); // c:spPr
        }

        #endregion

        private void insertFillProperties(AreaFormat areaFormat, GelFrameSequence gelFrameSequence)
        {
            // EG_FillProperties (from AreaFormat)
            if (areaFormat != null)
            {
                if (gelFrameSequence != null && gelFrameSequence.GelFrames.Count > 0)
                {
                    insertDrawingFillProperties(areaFormat, (ShapeOptions)gelFrameSequence.GelFrames[0].OPT1);
                }
                else
                {
                    if (areaFormat.fls == 0x0000)
                    {
                        // a:noFill (CT_NoFillProperties) 
                        _writer.WriteElementString(Dml.Prefix, Dml.ShapeEffects.ElNoFill, Dml.Ns, string.Empty);
                    }
                    else if (areaFormat.fls == 0x0001)
                    {
                        RGBColor fillColor;
                        if (this.ChartSheetContentSequence.Palette != null && areaFormat.icvFore >= 0x0000 && areaFormat.icvFore <= 0x0041)
                        {
                            // there is a valid palette color set
                            fillColor = this.ChartSheetContentSequence.Palette.rgbColorList[areaFormat.icvFore];
                        }
                        else
                        {
                            fillColor = areaFormat.rgbFore;
                        }
                        // a:solidFill (CT_SolidColorFillProperties)
                        _writer.WriteStartElement(Dml.Prefix, Dml.ShapeEffects.ElSolidFill, Dml.Ns);
                        writeValueElement(Dml.Prefix, Dml.BaseTypes.ElSrgbClr, Dml.Ns, fillColor.SixDigitHexCode.ToUpper());
                        _writer.WriteEndElement(); // a:solidFill 
                    }

                    // a:gradFill (CT_GradientFillProperties)

                    // a:blipFill (CT_BlipFillProperties)

                    // a:pattFill (CT_PatternFillProperties)

                    // a:grpFill (CT_GroupFillProperties)
                }
            }
        }

        private void insertLineProperties(LineFormat lineFormat, GelFrameSequence gelFrameSequence)
        {
            if (lineFormat != null)
            {
                // line style mapping
                string prstDash = "solid";
                string pattFillPrst = string.Empty;
                switch (lineFormat.lns)
                {
                    case LineFormat.LineStyle.Dash:
                        prstDash = "lgDash";
                        break;
                    case LineFormat.LineStyle.Dot:
                        prstDash = "sysDash";
                        break;
                    case LineFormat.LineStyle.DashDot:
                        prstDash = "lgDashDot";
                        break;
                    case LineFormat.LineStyle.DashDotDot:
                        prstDash = "lgDashDotDot";
                        break;
                    case LineFormat.LineStyle.None:
                        prstDash = "";
                        break;
                    case LineFormat.LineStyle.DarkGrayPattern:
                        pattFillPrst = "pct75";
                        break;
                    case LineFormat.LineStyle.MediumGrayPattern:
                        pattFillPrst = "pct50";
                        break;
                    case LineFormat.LineStyle.LightGrayPattern:
                        pattFillPrst = "pct25";
                        break;
                }

                // CT_LineProperties
                _writer.WriteStartElement(Dml.Prefix, Dml.ShapeProperties.ElLn, Dml.Ns);

                // w (line width)
                // map to the values used by the Compatibility Pack
                int lineWidth = 0;
                switch (lineFormat.we)
                {
                    case LineFormat.LineWeight.Hairline:
                        lineWidth = 3175;
                        break;
                    case LineFormat.LineWeight.Narrow:
                        lineWidth = 12700;
                        break;
                    case LineFormat.LineWeight.Medium:
                        lineWidth = 25400;
                        break;
                    case LineFormat.LineWeight.Wide:
                        lineWidth = 38100;
                        break;
                }
                if (lineWidth != 0)
                {
                    _writer.WriteAttributeString(Dml.ShapeLineProperties.AttrW, lineWidth.ToString());
                }

                // cap
                // cmpd
                // algn

                {
                    // EG_LineFillProperties
                    if (lineFormat.lns == LineFormat.LineStyle.None)
                    {
                        _writer.WriteElementString(Dml.Prefix, Dml.ShapeEffects.ElNoFill, Dml.Ns, string.Empty);
                    }
                    else
                    {
                        if (string.IsNullOrEmpty(pattFillPrst))
                        {
                            // a:solidFill (CT_SolidColorFillProperties)
                            _writer.WriteStartElement(Dml.Prefix, Dml.ShapeEffects.ElSolidFill, Dml.Ns);
                            writeValueElement(Dml.Prefix, Dml.BaseTypes.ElSrgbClr, Dml.Ns, lineFormat.rgb.SixDigitHexCode.ToUpper());
                            _writer.WriteEndElement(); // a:solidFill 
                        }
                        else
                        {
                            // a:pattFill
                            _writer.WriteStartElement(Dml.Prefix, Dml.ShapeEffects.ElPattFill, Dml.Ns);
                            _writer.WriteAttributeString(Dml.ShapeEffects.AttrPrst, pattFillPrst);

                            _writer.WriteStartElement(Dml.Prefix, Dml.ShapeEffects.ElFgClr, Dml.Ns);
                            writeValueElement(Dml.Prefix, Dml.BaseTypes.ElSrgbClr, Dml.Ns, lineFormat.rgb.SixDigitHexCode.ToUpper());
                            _writer.WriteEndElement();
                            _writer.WriteStartElement(Dml.Prefix, Dml.ShapeEffects.ElBgClr, Dml.Ns);
                            writeValueElement(Dml.Prefix, Dml.BaseTypes.ElSrgbClr, Dml.Ns, "FFFFFF");
                            _writer.WriteEndElement();
                            _writer.WriteEndElement(); // a:pattFill 
                        }
                    }

                    // EG_LineDashProperties
                    if (!string.IsNullOrEmpty(prstDash))
                    {
                        writeValueElement(Dml.Prefix, Dml.ShapeLineProperties.ElPrstDash, Dml.Ns, prstDash);
                    }

                    // EG_LineJoinProperties
                    // (not supported by Excel 2003)

                    // a:headEnd
                    // (not supported by Excel 2003)

                    // a:tailEnd
                    // (not supported by Excel 2003)
                }
                _writer.WriteEndElement(); // a:ln
            }
        }

        private void insertDrawingFillProperties(AreaFormat areaFormat, ShapeOptions so)
        {
            var rgbFore = areaFormat.rgbFore;
            var rgbBack = areaFormat.rgbBack;

            uint fillType = 0;
            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillType))
            {
                fillType = so.OptionsByID[ShapeOptions.PropertyId.fillType].op;
            }
            switch (fillType)
            {
                case 0x0: //solid
                    // a:solidFill (CT_SolidColorFillProperties)
                    _writer.WriteStartElement(Dml.Prefix, Dml.ShapeEffects.ElSolidFill, Dml.Ns);
                    {
                        // a:srgbColor
                        _writer.WriteStartElement(Dml.Prefix, Dml.BaseTypes.ElSrgbClr, Dml.Ns);
                        _writer.WriteAttributeString(Dml.BaseTypes.AttrVal, rgbFore.SixDigitHexCode.ToUpper());
                        _writer.WriteEndElement(); // a:srgbColor
                    }
                    _writer.WriteEndElement(); // a:solidFill
                    break;
                case 0x1: //pattern
                    //uint blipIndex1 = so.OptionsByID[ShapeOptions.PropertyId.fillBlip].op;
                    //DrawingGroup gr1 = (DrawingGroup)this._ctx.Ppt.DocumentRecord.FirstChildWithType<PPDrawingGroup>().Children[0];
                    //BlipStoreEntry bse1 = (BlipStoreEntry)gr1.FirstChildWithType<BlipStoreContainer>().Children[(int)blipIndex1 - 1];
                    //BitmapBlip b1 = (BitmapBlip)_ctx.Ppt.PicturesContainer._pictures[bse1.foDelay];

                    //_writer.WriteStartElement(Dml.Prefix, "pattFill", Dml.Ns);

                    //_writer.WriteAttributeString("prst", Utils.getPrstForPatternCode(b1.m_bTag)); //Utils.getPrstForPattern(blipNamePattern));

                    //_writer.WriteStartElement(Dml.Prefix, "fgClr", Dml.Ns);
                    //_writer.WriteStartElement(Dml.Prefix, "srgbClr", Dml.Ns);
                    //_writer.WriteAttributeString("val", Utils.getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[ShapeOptions.PropertyId.fillColor].op, slide, so));
                    //_writer.WriteEndElement();
                    //_writer.WriteEndElement();

                    //_writer.WriteStartElement(Dml.Prefix, "bgClr", Dml.Ns);
                    //_writer.WriteStartElement(Dml.Prefix, "srgbClr", Dml.Ns);
                    //if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillBackColor))
                    //{
                    //    colorval = Utils.getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[ShapeOptions.PropertyId.fillBackColor].op, slide, so);
                    //}
                    //else
                    //{
                    //    colorval = "ffffff"; //TODO: find out which color to use in this case
                    //}
                    //_writer.WriteAttributeString("val", colorval);
                    //_writer.WriteEndElement();
                    //_writer.WriteEndElement();

                    //_writer.WriteEndElement();

                    break;
                case 0x2: //texture
                case 0x3: //picture
                    //uint blipIndex = 0;
                    //if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillBlip))
                    //{
                    //    blipIndex = so.OptionsByID[ShapeOptions.PropertyId.fillBlip].op;
                    //}
                    //else
                    //{
                    //    blipIndex = so.OptionsByID[ShapeOptions.PropertyId.Pib].op;
                    //}

                    ////string blipName = Encoding.UTF8.GetString(so.OptionsByID[ShapeOptions.PropertyId.fillBlipName].opComplex);
                    //string rId = "";
                    //DrawingGroup gr = (DrawingGroup)this._ctx.Ppt.DocumentRecord.FirstChildWithType<PPDrawingGroup>().Children[0];

                    //if (blipIndex <= gr.FirstChildWithType<BlipStoreContainer>().Children.Count)
                    //{
                    //    BlipStoreEntry bse = (BlipStoreEntry)gr.FirstChildWithType<BlipStoreContainer>().Children[(int)blipIndex - 1];

                    //    if (_ctx.Ppt.PicturesContainer._pictures.ContainsKey(bse.foDelay))
                    //    {
                    //        Record rec = _ctx.Ppt.PicturesContainer._pictures[bse.foDelay];
                    //        ImagePart imgPart = null;
                    //        if (rec is BitmapBlip)
                    //        {
                    //            BitmapBlip b = (BitmapBlip)_ctx.Ppt.PicturesContainer._pictures[bse.foDelay];
                    //            imgPart = _parentSlideMapping.targetPart.AddImagePart(ShapeTreeMapping.getImageType(b.TypeCode));
                    //            imgPart.TargetDirectory = "..\\media";
                    //            System.IO.Stream outStream = imgPart.GetStream();
                    //            outStream.Write(b.m_pvBits, 0, b.m_pvBits.Length);
                    //        }
                    //        else
                    //        {
                    //            MetafilePictBlip b = (MetafilePictBlip)_ctx.Ppt.PicturesContainer._pictures[bse.foDelay];
                    //            imgPart = _parentSlideMapping.targetPart.AddImagePart(ShapeTreeMapping.getImageType(b.TypeCode));
                    //            imgPart.TargetDirectory = "..\\media";
                    //            System.IO.Stream outStream = imgPart.GetStream();
                    //            byte[] decompressed = b.Decrompress();
                    //            outStream.Write(decompressed, 0, decompressed.Length);
                    //        }

                    //        rId = imgPart.RelIdToString;

                    //        _writer.WriteStartElement(Dml.Prefix, "blipFill", Dml.Ns);
                    //        _writer.WriteAttributeString("dpi", "0");
                    //        _writer.WriteAttributeString("rotWithShape", "1");

                    //        _writer.WriteStartElement(Dml.Prefix, "blip", Dml.Ns);
                    //        _writer.WriteAttributeString("r", "embed", OpenXmlNamespaces.Relationships, rId);



                    //        _writer.WriteEndElement();

                    //        _writer.WriteElementString(Dml.Prefix, "srcRect", Dml.Ns, "");

                    //        if (fillType == 0x3)
                    //        {
                    //            _writer.WriteStartElement(Dml.Prefix, "stretch", Dml.Ns);
                    //            _writer.WriteElementString(Dml.Prefix, "fillRect", Dml.Ns, "");
                    //            _writer.WriteEndElement();
                    //        }
                    //        else
                    //        {
                    //            _writer.WriteStartElement(Dml.Prefix, "tile", Dml.Ns);
                    //            _writer.WriteAttributeString("tx", "0");
                    //            _writer.WriteAttributeString("ty", "0");
                    //            _writer.WriteAttributeString("sx", "100000");
                    //            _writer.WriteAttributeString("sy", "100000");
                    //            _writer.WriteAttributeString("flip", "none");
                    //            _writer.WriteAttributeString("algn", "tl");
                    //            _writer.WriteEndElement();
                    //        }

                    //        _writer.WriteEndElement();
                    //    }
                    //}
                    break;
                case 0x4: //shade
                case 0x5: //shadecenter
                case 0x6: //shadeshape
                case 0x7: //shadescale
                    _writer.WriteStartElement(Dml.Prefix, "gradFill", Dml.Ns);
                    _writer.WriteAttributeString("rotWithShape", "1");
                    {
                        _writer.WriteStartElement(Dml.Prefix, "gsLst", Dml.Ns);
                        {
                            _writer.WriteStartElement(Dml.Prefix, "gs", Dml.Ns);
                            _writer.WriteAttributeString("pos", "0");

                            _writer.WriteStartElement(Dml.Prefix, "srgbClr", Dml.Ns);
                            _writer.WriteAttributeString("val", rgbFore.SixDigitHexCode.ToUpper());
                            //if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillOpacity))
                            //{
                            //    _writer.WriteStartElement(Dml.Prefix, "alpha", Dml.Ns);
                            //    _writer.WriteAttributeString("val", Math.Round(((decimal)so.OptionsByID[ShapeOptions.PropertyId.fillOpacity].op / 65536 * 100000)).ToString()); //we need the percentage of the opacity (65536 means 100%)
                            //    _writer.WriteEndElement();
                            //}
                            _writer.WriteEndElement();
                            _writer.WriteEndElement();

                            _writer.WriteStartElement(Dml.Prefix, "gs", Dml.Ns);
                            _writer.WriteAttributeString("pos", "100000");
                            _writer.WriteStartElement(Dml.Prefix, "srgbClr", Dml.Ns);
                            _writer.WriteAttributeString("val", rgbBack.SixDigitHexCode.ToUpper());
                            //if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillBackOpacity))
                            //{
                            //    _writer.WriteStartElement(Dml.Prefix, "alpha", Dml.Ns);
                            //    _writer.WriteAttributeString("val", Math.Round(((decimal)so.OptionsByID[ShapeOptions.PropertyId.fillBackOpacity].op / 65536 * 100000)).ToString()); //we need the percentage of the opacity (65536 means 100%)
                            //    _writer.WriteEndElement();
                            //}
                            _writer.WriteEndElement();
                            _writer.WriteEndElement();
                        }

                        _writer.WriteEndElement(); //gsLst

                        switch (fillType)
                        {
                            case 0x5:
                            case 0x6:
                                _writer.WriteStartElement(Dml.Prefix, "path", Dml.Ns);
                                _writer.WriteAttributeString("path", "shape");
                                _writer.WriteStartElement(Dml.Prefix, "fillToRect", Dml.Ns);
                                _writer.WriteAttributeString("l", "50000");
                                _writer.WriteAttributeString("t", "50000");
                                _writer.WriteAttributeString("r", "50000");
                                _writer.WriteAttributeString("b", "50000");
                                _writer.WriteEndElement();
                                _writer.WriteEndElement(); //path
                                break;
                            case 0x7:
                                decimal angle = 90;
                                if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillAngle))
                                {
                                    if (so.OptionsByID[ShapeOptions.PropertyId.fillAngle].op != 0)
                                    {
                                        var bytes = BitConverter.GetBytes(so.OptionsByID[ShapeOptions.PropertyId.fillAngle].op);
                                        int integral = BitConverter.ToInt16(bytes, 0);
                                        uint fractional = BitConverter.ToUInt16(bytes, 2);
                                        decimal result = integral + ((decimal)fractional / (decimal)65536);
                                        angle = 65536 - fractional; //I have no idea why this works!!                    
                                        angle = angle - 90;
                                        if (angle < 0)
                                        {
                                            angle += 360;
                                        }
                                    }
                                }
                                _writer.WriteStartElement(Dml.Prefix, "lin", Dml.Ns);

                                angle *= 60000;
                                if (angle > 5400000) angle = 5400000;

                                _writer.WriteAttributeString("ang", angle.ToString());
                                _writer.WriteAttributeString("scaled", "1");
                                _writer.WriteEndElement();
                                break;
                            default:
                                _writer.WriteStartElement(Dml.Prefix, "path", Dml.Ns);
                                _writer.WriteAttributeString("path", "rect");
                                _writer.WriteStartElement(Dml.Prefix, "fillToRect", Dml.Ns);
                                _writer.WriteAttributeString("r", "100000");
                                _writer.WriteAttributeString("b", "100000");
                                _writer.WriteEndElement();
                                _writer.WriteEndElement(); //path
                                break;
                        }
                    }
                    _writer.WriteEndElement(); //gradFill

                    break;
                //case 0x7: //shadescale
                    //_writer.WriteStartElement(Dml.Prefix, "gradFill", Dml.Ns);
                    //_writer.WriteAttributeString("rotWithShape", "1");
                    //_writer.WriteStartElement(Dml.Prefix, "gsLst", Dml.Ns);

                    //bool switchColors = false;
                    //decimal angle = 90;
                    //if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillAngle))
                    //{
                    //    if (so.OptionsByID[ShapeOptions.PropertyId.fillAngle].op != 0)
                    //    {
                    //        byte[] bytes = BitConverter.GetBytes(so.OptionsByID[ShapeOptions.PropertyId.fillAngle].op);
                    //        int integral = BitConverter.ToInt16(bytes, 0);
                    //        uint fractional = BitConverter.ToUInt16(bytes, 2);
                    //        decimal result = integral + ((decimal)fractional / (decimal)65536);
                    //        angle = 65536 - fractional; //I have no idea why this works!!                    
                    //        angle = angle - 90;
                    //        if (angle < 0)
                    //        {
                    //            angle += 360;
                    //            switchColors = true;
                    //        }
                    //    }
                    //}

                    //Dictionary<int, string> shadeColorsDic = new Dictionary<int, string>();
                    //List<string> shadeColors = new List<string>();
                    //if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillShadeColors))
                    //{
                    //    uint length = so.OptionsByID[ShapeOptions.PropertyId.fillShadeColors].op;

                    //    //An IMsoArray record that specifies colors and their relative positions. 
                    //    //Each element of the array contains an OfficeArtCOLORREF record color and a FixedPoint, as specified in [MS-OSHARED] 
                    //    //section 2.2.1.6, that specifies its relative position along the gradient vector.
                    //    byte[] data = so.OptionsByID[ShapeOptions.PropertyId.fillShadeColors].opComplex;

                    //    int pos = 0;
                    //    string colval;
                    //    FixedPointNumber fixedpoint;
                    //    UInt16 nElems = BitConverter.ToUInt16(data, pos);
                    //    pos += 2;
                    //    UInt16 nElemsAlloc = BitConverter.ToUInt16(data, pos);
                    //    pos += 2;
                    //    UInt16 cbElem = BitConverter.ToUInt16(data, pos);
                    //    pos += 2;

                    //    if (cbElem == 0xFFF0)
                    //    {
                    //        //If this value is 0xFFF0 then this record is an array of truncated 8 byte elements. Only the 4 low-order bytes are recorded. Each element's 4 high-order bytes equal 0x00000000 and each element's 4 low-order bytes are contained in data.
                    //    }
                    //    else
                    //    {
                    //        while (pos < length)
                    //        {
                    //            colval = Utils.getRGBColorFromOfficeArtCOLORREF(BitConverter.ToUInt32(data, pos), slide, so);

                    //            pos += 4;
                    //            fixedpoint = new FixedPointNumber(BitConverter.ToUInt16(data, pos), BitConverter.ToUInt16(data, pos + 2));
                    //            shadeColors.Insert(0, colval);
                    //            pos += 4;
                    //        }
                    //    }

                    //}
                    //else
                    //{
                    //    bool switchcolors = false;
                    //    if (switchColors & so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillBackColor))
                    //    {
                    //        colorval = Utils.getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[ShapeOptions.PropertyId.fillBackColor].op, slide, so);
                    //    }
                    //    else
                    //    {
                    //        if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillColor))
                    //        {
                    //            colorval = Utils.getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[ShapeOptions.PropertyId.fillColor].op, slide, so);
                    //        }
                    //        else
                    //        {
                    //            colorval = "FFFFFF"; //TODO: find out which color to use in this case
                    //            switchcolors = true;
                    //        }
                    //    }

                    //    if (switchColors | !so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillBackColor))
                    //    {
                    //        colorval2 = Utils.getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[ShapeOptions.PropertyId.fillColor].op, slide, so);
                    //    }
                    //    else
                    //    {
                    //        colorval2 = Utils.getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[ShapeOptions.PropertyId.fillBackColor].op, slide, so);
                    //    }

                    //    if (switchcolors)
                    //    {
                    //        //this is a workaround for a bug. Further analysis necessarry
                    //        string dummy = colorval;
                    //        colorval = colorval2;
                    //        colorval2 = dummy;
                    //    }

                    //    shadeColors.Add(colorval);
                    //    shadeColors.Add(colorval2);
                    //}


                    //int gspos;
                    //string col;
                    //for (int i = 0; i < shadeColors.Count; i++)
                    //{
                    //    col = shadeColors[i];
                    //    if (i == 0)
                    //    {
                    //        gspos = 0;
                    //    }
                    //    else if (i == shadeColors.Count - 1)
                    //    {
                    //        gspos = 100000;
                    //    }
                    //    else
                    //    {
                    //        gspos = i * 100000 / shadeColors.Count;
                    //    }

                    //    _writer.WriteStartElement(Dml.Prefix, "gs", Dml.Ns);
                    //    _writer.WriteAttributeString("pos", gspos.ToString());
                    //    _writer.WriteStartElement(Dml.Prefix, "srgbClr", Dml.Ns);
                    //    _writer.WriteAttributeString("val", col);
                    //    if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillOpacity))
                    //    {
                    //        _writer.WriteStartElement(Dml.Prefix, "alpha", Dml.Ns);
                    //        decimal alpha = Math.Round(((decimal)so.OptionsByID[ShapeOptions.PropertyId.fillOpacity].op / 65536 * 100000)); //we need the percentage of the opacity (65536 means 100%)
                    //        _writer.WriteAttributeString("val", alpha.ToString(CultureInfo.InvariantCulture)); 
                    //        _writer.WriteEndElement();
                    //    }

                    //    if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillShadeType))
                    //    {
                    //        uint flags = so.OptionsByID[ShapeOptions.PropertyId.fillShadeType].op;
                    //        bool none = Tools.Utils.BitmaskToBool(flags, 0x1);
                    //        bool gamma = Tools.Utils.BitmaskToBool(flags, 0x1 << 1);
                    //        bool sigma = Tools.Utils.BitmaskToBool(flags, 0x1 << 2);
                    //        bool band = Tools.Utils.BitmaskToBool(flags, 0x1 << 3);
                    //        bool onecolor = Tools.Utils.BitmaskToBool(flags, 0x1 << 4);

                    //        if (gamma) _writer.WriteElementString(Dml.Prefix, "gamma", Dml.Ns, "");
                    //        if (band)
                    //        {
                    //            _writer.WriteStartElement(Dml.Prefix, "shade", Dml.Ns);
                    //            _writer.WriteAttributeString("val", "37255");
                    //            _writer.WriteEndElement();
                    //        }
                    //        if (gamma) _writer.WriteElementString(Dml.Prefix, "invGamma", Dml.Ns, "");
                    //    }
                    //    _writer.WriteEndElement();
                    //    _writer.WriteEndElement();
                    //}




                    //////new colorval
                    ////_writer.WriteStartElement(Dml.Prefix, "gs", Dml.Ns);
                    ////_writer.WriteAttributeString("pos", "100000");
                    ////_writer.WriteStartElement(Dml.Prefix, "srgbClr", Dml.Ns);
                    ////_writer.WriteAttributeString("val", colorval2);
                    ////if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillBackOpacity))
                    ////{
                    ////    _writer.WriteStartElement(Dml.Prefix, "alpha", Dml.Ns);
                    ////    _writer.WriteAttributeString("val", Math.Round(((decimal)so.OptionsByID[ShapeOptions.PropertyId.fillBackOpacity].op / 65536 * 100000)).ToString()); //we need the percentage of the opacity (65536 means 100%)
                    ////    _writer.WriteEndElement();
                    ////}

                    ////_writer.WriteEndElement();
                    ////_writer.WriteEndElement();

                    //_writer.WriteEndElement(); //gsLst

                    //_writer.WriteStartElement(Dml.Prefix, "lin", Dml.Ns);

                    //angle *= 60000;
                    //if (angle > 5400000) angle = 5400000;

                    //_writer.WriteAttributeString("ang", angle.ToString());
                    //_writer.WriteAttributeString("scaled", "1");
                    //_writer.WriteEndElement();

                    //_writer.WriteEndElement();
                    break;
                case 0x8: //shadetitle
                case 0x9: //background
                    break;

            }
        }

    }
}
