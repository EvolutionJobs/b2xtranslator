

using System;
using b2xtranslator.PptFileFormat;
using b2xtranslator.CommonTranslatorLib;
using System.Xml;
using b2xtranslator.OpenXmlLib;
using b2xtranslator.OfficeDrawing;

namespace b2xtranslator.PresentationMLMapping
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
            this._ctx = ctx;
        }

        private static int counter = 0;
        public void Apply(ShapeOptions pso)
        {
            this.so = pso;

            counter++;

            if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.shadowType))
            {
                switch (this.so.OptionsByID[ShapeOptions.PropertyId.shadowType].op)
                {
                    case 0: //offset
                        writeOffset();
                        break;
                    case 1: //double
                        this._writer.WriteStartElement("a", "effectLst", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteStartElement("a", "prstShdw", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("prst", "shdw13");
                        if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.shadowOffsetX))
                            if (this.so.OptionsByID[ShapeOptions.PropertyId.shadowOffsetX].op != 0)
                                writeDistDir();

                        writeColor();

                        this._writer.WriteEndElement();
                        this._writer.WriteEndElement();
                        break;
                    case 2: //rich
                            //shadow offset and  a transformation
                        this._writer.WriteStartElement("a", "effectLst", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteStartElement("a", "outerShdw", OpenXmlNamespaces.DrawingML);

                         if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.shadowOffsetX))
                             if (this.so.OptionsByID[ShapeOptions.PropertyId.shadowOffsetX].op != 0)
                                 writeDistDir();
                         if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.shadowOriginX))
                         {
                             var bytes = BitConverter.GetBytes(this.so.OptionsByID[ShapeOptions.PropertyId.shadowOriginX].op);
                             int integral = BitConverter.ToInt16(bytes, 0);
                             uint fractional = BitConverter.ToUInt16(bytes, 2);
                             if (fractional == 0xffff) integral *= -1;
                             Decimal origX = integral; // +((decimal)fractional / (decimal)65536);

                             bytes = BitConverter.GetBytes(this.so.OptionsByID[ShapeOptions.PropertyId.shadowOriginY].op);
                             integral = BitConverter.ToInt16(bytes, 0);
                             fractional = BitConverter.ToUInt16(bytes, 2);
                             if (fractional == 0xffff) integral *= -1;
                             Decimal origY = integral; // +((decimal)fractional / (decimal)65536);

                             if (origX > 0)
                             {
                                 if (origY > 0)
                                 {
                                    this._writer.WriteAttributeString("algn", "tl");
                                 }
                                 else
                                 {
                                    this._writer.WriteAttributeString("algn", "b");
                                 }
                             }
                             else
                             {                                 
                                 if (origY > 0)
                                 {
                                     
                                 }
                                 else
                                 {
                                    this._writer.WriteAttributeString("algn", "br");
                                 }
                             }
                         }
                         else
                         {
                            this._writer.WriteAttributeString("algn", "b");
                         }

                        

                        if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.shadowScaleXToX))
                        {
                            var bytes = BitConverter.GetBytes(this.so.OptionsByID[ShapeOptions.PropertyId.shadowScaleXToX].op);
                            int integral = -1 * BitConverter.ToInt16(bytes, 0);
                            uint fractional = BitConverter.ToUInt16(bytes, 2);
                            if (fractional == 0xffff) integral *= -1;
                            decimal result = integral + ((decimal)fractional / (decimal)65536);
                            result = 1 - (result / 65536);
                            this._writer.WriteAttributeString("sx", Math.Floor(result * 100000).ToString());
                        }
                        if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.shadowScaleXToY))
                        {
                            int scaleXY = Utils.EMUToMasterCoord((int)this.so.OptionsByID[ShapeOptions.PropertyId.shadowScaleXToY].op);
                        }
                        if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.shadowScaleYToX))
                        {
                            decimal scaleYX = (Decimal)(int)this.so.OptionsByID[ShapeOptions.PropertyId.shadowScaleYToX].op;
                            //_writer.WriteAttributeString("kx", System.Math.Floor(scaleYX / 138790 * 100 * 60000).ToString()); //The 138790 comes from reverse engineering. I can't find a hint in the spec about how to convert this
                            if (scaleYX < 0)
                            {
                                this._writer.WriteAttributeString("kx", "-2453606");
                            }
                            else
                            {
                                this._writer.WriteAttributeString("kx", "2453606");
                            }
                        }
                        if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.shadowScaleYToY))
                        {
                            var bytes = BitConverter.GetBytes(this.so.OptionsByID[ShapeOptions.PropertyId.shadowScaleYToY].op);
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
                                if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.shadowScaleYToX))
                                {
                                    result = (Decimal)(-0.5);
                                }
                                else
                                {
                                    result = (Decimal)(-1);
                                }
                            }
                            this._writer.WriteAttributeString("sy", Math.Floor(result * 100000).ToString());
                        }
                        else
                        {
                            this._writer.WriteAttributeString("sy","50000");
                        }

                        writeColor();

                        this._writer.WriteEndElement();
                        this._writer.WriteEndElement();
                        break;
                    case 3: //shape
                        break;
                    case 4: //drawing
                        break;
                    case 5: //embossOrEngrave
                        this._writer.WriteStartElement("a", "effectLst", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteStartElement("a", "prstShdw", OpenXmlNamespaces.DrawingML);

                        if (this.so.OptionsByID[ShapeOptions.PropertyId.shadowOffsetX].op == 0x319c)
                        {
                            this._writer.WriteAttributeString("prst", "shdw17");
                        }
                        else
                        {
                            this._writer.WriteAttributeString("prst", "shdw18");
                        }
                        if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.shadowOffsetX))
                            if (this.so.OptionsByID[ShapeOptions.PropertyId.shadowOffsetX].op != 0)
                                writeDistDir();

                        this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                        string colorval = Utils.getRGBColorFromOfficeArtCOLORREF(this.so.OptionsByID[ShapeOptions.PropertyId.shadowColor].op, this.so.FirstAncestorWithType<Slide>(), this.so);
                        this._writer.WriteAttributeString("val", colorval);
                        if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.shadowOpacity) && this.so.OptionsByID[ShapeOptions.PropertyId.shadowOpacity].op != 65536)
                        {
                            this._writer.WriteStartElement("a", "alpha", OpenXmlNamespaces.DrawingML);
                            this._writer.WriteAttributeString("val", Math.Round(((decimal)this.so.OptionsByID[ShapeOptions.PropertyId.shadowOpacity].op / 65536 * 100000)).ToString()); //we need the percentage of the opacity (65536 means 100%)
                            this._writer.WriteEndElement();
                        }
                        this._writer.WriteElementString("a", "gamma", OpenXmlNamespaces.DrawingML, "");
                        this._writer.WriteStartElement("a", "shade", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("val", "60000");
                        this._writer.WriteEndElement();
                        this._writer.WriteElementString("a", "invGamma", OpenXmlNamespaces.DrawingML, "");
                        this._writer.WriteEndElement();

                        this._writer.WriteEndElement();
                        this._writer.WriteEndElement();
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
            this._writer.WriteStartElement("a", "effectLst", OpenXmlNamespaces.DrawingML);
            this._writer.WriteStartElement("a", "outerShdw", OpenXmlNamespaces.DrawingML);
            writeDistDir();
            this._writer.WriteAttributeString("algn", "ctr");
            writeColor();
            this._writer.WriteEndElement();
            this._writer.WriteEndElement();
        }

        private void writeColor()
        {
            this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
            string colorval = "808080";
            if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.shadowColor)) colorval = Utils.getRGBColorFromOfficeArtCOLORREF(this.so.OptionsByID[ShapeOptions.PropertyId.shadowColor].op, this.so.FirstAncestorWithType<Slide>(), this.so);
            this._writer.WriteAttributeString("val", colorval);
            if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.shadowOpacity) && this.so.OptionsByID[ShapeOptions.PropertyId.shadowOpacity].op != 65536)
            {
                this._writer.WriteStartElement("a", "alpha", OpenXmlNamespaces.DrawingML);
                this._writer.WriteAttributeString("val", Math.Round(((decimal)this.so.OptionsByID[ShapeOptions.PropertyId.shadowOpacity].op / 65536 * 100000)).ToString()); //we need the percentage of the opacity (65536 means 100%)
                this._writer.WriteEndElement();
            }
            this._writer.WriteEndElement();
        }

        private void writeDistDir()
        {
            int distX = Utils.EMUToMasterCoord(0x6338);
            int distY = Utils.EMUToMasterCoord(0x6338);
            if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.shadowOffsetX)) distX = Utils.EMUToMasterCoord((int)this.so.OptionsByID[ShapeOptions.PropertyId.shadowOffsetX].op);
            if (this.so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.shadowOffsetY)) distY = Utils.EMUToMasterCoord((int)this.so.OptionsByID[ShapeOptions.PropertyId.shadowOffsetY].op);
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

            this._writer.WriteAttributeString("dist", dist.ToString());
            this._writer.WriteAttributeString("dir", dir);
        }
                
    }
}
