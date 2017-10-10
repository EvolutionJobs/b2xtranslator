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
using DIaLOGIKa.b2xtranslator.PptFileFormat;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.PresentationMLMapping
{
    class ShadowMapping :
        AbstractOpenXmlMapping,
        IMapping<ShapeOptions>
    {
        protected ConversionContext _ctx;
        private ShapeOptions so;

        public ShadowMapping(ConversionContext ctx, XmlWriter writer)
            : base(writer)
        {
            _ctx = ctx;
        }

        private static int counter = 0;
        public void Apply(ShapeOptions pso)
        {
            so = pso;

            counter++;

            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.shadowType))
            {
                switch (so.OptionsByID[ShapeOptions.PropertyId.shadowType].op)
                {
                    case 0: //offset
                        writeOffset();
                        break;
                    case 1: //double
                        _writer.WriteStartElement("a", "effectLst", OpenXmlNamespaces.DrawingML);
                        _writer.WriteStartElement("a", "prstShdw", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("prst", "shdw13");
                        if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.shadowOffsetX))
                            if (so.OptionsByID[ShapeOptions.PropertyId.shadowOffsetX].op != 0)
                                writeDistDir();

                        writeColor();

                        _writer.WriteEndElement();
                        _writer.WriteEndElement();
                        break;
                    case 2: //rich
                        //shadow offset and  a transformation
                         _writer.WriteStartElement("a", "effectLst", OpenXmlNamespaces.DrawingML);
                         _writer.WriteStartElement("a", "outerShdw", OpenXmlNamespaces.DrawingML);

                         if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.shadowOffsetX))
                             if (so.OptionsByID[ShapeOptions.PropertyId.shadowOffsetX].op != 0)
                                 writeDistDir();
                         if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.shadowOriginX))
                         {
                             var bytes = BitConverter.GetBytes(so.OptionsByID[ShapeOptions.PropertyId.shadowOriginX].op);
                             int integral = BitConverter.ToInt16(bytes, 0);
                             uint fractional = BitConverter.ToUInt16(bytes, 2);
                             if (fractional == 0xffff) integral *= -1;
                             Decimal origX = integral; // +((decimal)fractional / (decimal)65536);

                             bytes = BitConverter.GetBytes(so.OptionsByID[ShapeOptions.PropertyId.shadowOriginY].op);
                             integral = BitConverter.ToInt16(bytes, 0);
                             fractional = BitConverter.ToUInt16(bytes, 2);
                             if (fractional == 0xffff) integral *= -1;
                             Decimal origY = integral; // +((decimal)fractional / (decimal)65536);

                             if (origX > 0)
                             {
                                 if (origY > 0)
                                 {
                                     _writer.WriteAttributeString("algn", "tl");
                                 }
                                 else
                                 {
                                     _writer.WriteAttributeString("algn", "b");
                                 }
                             }
                             else
                             {                                 
                                 if (origY > 0)
                                 {
                                     
                                 }
                                 else
                                 {
                                     _writer.WriteAttributeString("algn", "br");
                                 }
                             }
                         }
                         else
                         {
                            _writer.WriteAttributeString("algn", "b");
                         }

                        

                        if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.shadowScaleXToX))
                        {
                            var bytes = BitConverter.GetBytes(so.OptionsByID[ShapeOptions.PropertyId.shadowScaleXToX].op);
                            int integral = -1 * BitConverter.ToInt16(bytes, 0);
                            uint fractional = BitConverter.ToUInt16(bytes, 2);
                            if (fractional == 0xffff) integral *= -1;
                            var result = integral + ((decimal)fractional / (decimal)65536);
                            result = 1 - (result / 65536);
                            _writer.WriteAttributeString("sx", Math.Floor(result * 100000).ToString());
                        }
                        if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.shadowScaleXToY))
                        {
                            int scaleXY = Utils.EMUToMasterCoord((int)so.OptionsByID[ShapeOptions.PropertyId.shadowScaleXToY].op);
                        }
                        if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.shadowScaleYToX))
                        {
                            var scaleYX = (Decimal)(int)so.OptionsByID[ShapeOptions.PropertyId.shadowScaleYToX].op;
                            //_writer.WriteAttributeString("kx", System.Math.Floor(scaleYX / 138790 * 100 * 60000).ToString()); //The 138790 comes from reverse engineering. I can't find a hint in the spec about how to convert this
                            if (scaleYX < 0)
                            {
                                _writer.WriteAttributeString("kx", "-2453606");
                            }
                            else
                            {
                                _writer.WriteAttributeString("kx", "2453606");
                            }
                        }
                        if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.shadowScaleYToY))
                        {
                            var bytes = BitConverter.GetBytes(so.OptionsByID[ShapeOptions.PropertyId.shadowScaleYToY].op);
                            int integral = -1 * BitConverter.ToInt16(bytes, 0);
                            uint fractional = BitConverter.ToUInt16(bytes, 2);
                            if (fractional == 0xffff) integral *= -1;
                            Decimal result = integral; // +((decimal)fractional / (decimal)65536);
                            if (fractional != 0xffff)
                            {
                                result = 1 - (result / 65536);
                            }
                            else
                            {
                                result = (result / 65536);
                            }
                            if (result == 0)
                            {
                                if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.shadowScaleYToX))
                                {
                                    result = (Decimal)(-0.5);
                                }
                                else
                                {
                                    result = (Decimal)(-1);
                                }
                            }
                            _writer.WriteAttributeString("sy", Math.Floor(result * 100000).ToString());
                        }
                        else
                        {
                            _writer.WriteAttributeString("sy","50000");
                        }

                        writeColor();

                        _writer.WriteEndElement();
                        _writer.WriteEndElement();
                        break;
                    case 3: //shape
                        break;
                    case 4: //drawing
                        break;
                    case 5: //embossOrEngrave
                        _writer.WriteStartElement("a", "effectLst", OpenXmlNamespaces.DrawingML);
                        _writer.WriteStartElement("a", "prstShdw", OpenXmlNamespaces.DrawingML);

                        if (so.OptionsByID[ShapeOptions.PropertyId.shadowOffsetX].op == 0x319c)
                        {
                            _writer.WriteAttributeString("prst", "shdw17");
                        }
                        else
                        {
                            _writer.WriteAttributeString("prst", "shdw18");
                        }
                        if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.shadowOffsetX))
                            if (so.OptionsByID[ShapeOptions.PropertyId.shadowOffsetX].op != 0)
                                writeDistDir();

                        _writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                        string colorval = Utils.getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[ShapeOptions.PropertyId.shadowColor].op, so.FirstAncestorWithType<Slide>(), so);
                        _writer.WriteAttributeString("val", colorval);
                        if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.shadowOpacity) && so.OptionsByID[ShapeOptions.PropertyId.shadowOpacity].op != 65536)
                        {
                            _writer.WriteStartElement("a", "alpha", OpenXmlNamespaces.DrawingML);
                            _writer.WriteAttributeString("val", Math.Round(((decimal)so.OptionsByID[ShapeOptions.PropertyId.shadowOpacity].op / 65536 * 100000)).ToString()); //we need the percentage of the opacity (65536 means 100%)
                            _writer.WriteEndElement();
                        }
                        _writer.WriteElementString("a", "gamma", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteStartElement("a", "shade", OpenXmlNamespaces.DrawingML);
                        _writer.WriteAttributeString("val", "60000");
                        _writer.WriteEndElement();
                        _writer.WriteElementString("a", "invGamma", OpenXmlNamespaces.DrawingML, "");
                        _writer.WriteEndElement();

                        _writer.WriteEndElement();
                        _writer.WriteEndElement();
                        break;
                }
            }
            else
            {
                //default is offset
                writeOffset();
            }
            
 
        }

        private void writeOffset()
        {
            _writer.WriteStartElement("a", "effectLst", OpenXmlNamespaces.DrawingML);
            _writer.WriteStartElement("a", "outerShdw", OpenXmlNamespaces.DrawingML);
            writeDistDir();            
            _writer.WriteAttributeString("algn", "ctr");
            writeColor();
            _writer.WriteEndElement();
            _writer.WriteEndElement();
        }

        private void writeColor()
        {
            _writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
            string colorval = "808080";
            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.shadowColor)) colorval = Utils.getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[ShapeOptions.PropertyId.shadowColor].op, so.FirstAncestorWithType<Slide>(), so);
            _writer.WriteAttributeString("val", colorval);
            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.shadowOpacity) && so.OptionsByID[ShapeOptions.PropertyId.shadowOpacity].op != 65536)
            {
                _writer.WriteStartElement("a", "alpha", OpenXmlNamespaces.DrawingML);
                _writer.WriteAttributeString("val", Math.Round(((decimal)so.OptionsByID[ShapeOptions.PropertyId.shadowOpacity].op / 65536 * 100000)).ToString()); //we need the percentage of the opacity (65536 means 100%)
                _writer.WriteEndElement();
            }
            _writer.WriteEndElement();
        }

        private void writeDistDir()
        {
            int distX = Utils.EMUToMasterCoord(0x6338);
            int distY = Utils.EMUToMasterCoord(0x6338);
            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.shadowOffsetX)) distX = Utils.EMUToMasterCoord((int)so.OptionsByID[ShapeOptions.PropertyId.shadowOffsetX].op);
            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.shadowOffsetY)) distY = Utils.EMUToMasterCoord((int)so.OptionsByID[ShapeOptions.PropertyId.shadowOffsetY].op);
            string dir = "18900000";
            if (distX < 0)
            {
                if (distY < 0)
                {
                    dir = "13500000";
                }
                else
                {
                    dir = "8100000";
                }
            }
            else
            {
                if (distY < 0)
                {
                    dir = "18900000";
                }
                else
                {
                    dir = "2700000";
                }
            }
            if (distX < 0)
            {
                distX *= -1;
            }
            if (distY < 0)
            {
                distY *= -1;
            }


            int dist = Utils.MasterCoordToEMU((int)System.Math.Round(System.Math.Sqrt(distX * distX + distY * distY)));

            _writer.WriteAttributeString("dist", dist.ToString());
            _writer.WriteAttributeString("dir", dir);
        }
                
    }
}
