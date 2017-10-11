

using System;
using System.Collections.Generic;
using System.Text;
using b2xtranslator.PptFileFormat;
using b2xtranslator.CommonTranslatorLib;
using System.Xml;
using b2xtranslator.OpenXmlLib;
using b2xtranslator.OfficeDrawing;
using b2xtranslator.Tools;

namespace b2xtranslator.PresentationMLMapping
{
    class FillMapping :
        AbstractOpenXmlMapping
    {
        protected ConversionContext _ctx;
        protected PresentationMapping<RegularContainer> _parentSlideMapping;

        public FillMapping(ConversionContext ctx, XmlWriter writer, PresentationMapping<RegularContainer> parentSlideMapping)
            : base(writer)
        {
            this._ctx = ctx;
            this._parentSlideMapping = parentSlideMapping;
        }

        public void Apply(ShapeOptions so)
        {
            RegularContainer slide = so.FirstAncestorWithType<Slide>();
            if (slide == null) slide = so.FirstAncestorWithType<Note>();
            if (slide == null) slide = so.FirstAncestorWithType<Handout>();
            string colorval = "";
            string colorval2 = "";
            uint fillType = 0;
            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillType)) fillType = so.OptionsByID[ShapeOptions.PropertyId.fillType].op;
            switch (fillType)
            {
                case 0x0: //solid
                    string SchemeType = "";


                    if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillColor))
                    {
                        colorval = Utils.getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[ShapeOptions.PropertyId.fillColor].op, (RegularContainer)slide, so, ref SchemeType);
                    } else {
                        colorval = "FFFFFF"; //TODO: find out which color to use in this case
                    }
                    this._writer.WriteStartElement("a", "solidFill", OpenXmlNamespaces.DrawingML);

                    if (SchemeType.Length == 0)
                    {
                        this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("val", colorval);
                    }
                    else
                    {
                        this._writer.WriteStartElement("a", "schemeClr", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("val", SchemeType);
                    }

                    if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillOpacity) && so.OptionsByID[ShapeOptions.PropertyId.fillOpacity].op != 65536)
                    {
                        this._writer.WriteStartElement("a", "alpha", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("val", Math.Round(((decimal)so.OptionsByID[ShapeOptions.PropertyId.fillOpacity].op / 65536 * 100000)).ToString()); //we need the percentage of the opacity (65536 means 100%)
                        this._writer.WriteEndElement();
                    }
                    this._writer.WriteEndElement();
                    this._writer.WriteEndElement();
                    break;
                case 0x1: //pattern
                    uint blipIndex1 = so.OptionsByID[ShapeOptions.PropertyId.fillBlip].op;
                    var gr1 = (DrawingGroup)this._ctx.Ppt.DocumentRecord.FirstChildWithType<PPDrawingGroup>().Children[0];
                    var bse1 = (BlipStoreEntry)gr1.FirstChildWithType<BlipStoreContainer>().Children[(int)blipIndex1 - 1];
                    var b1 = (BitmapBlip)this._ctx.Ppt.PicturesContainer._pictures[bse1.foDelay];

                    this._writer.WriteStartElement("a", "pattFill", OpenXmlNamespaces.DrawingML);

                    this._writer.WriteAttributeString("prst", Utils.getPrstForPatternCode(b1.m_bTag)); //Utils.getPrstForPattern(blipNamePattern));

                    this._writer.WriteStartElement("a", "fgClr", OpenXmlNamespaces.DrawingML);
                    this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                    this._writer.WriteAttributeString("val", Utils.getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[ShapeOptions.PropertyId.fillColor].op, slide,so));
                    this._writer.WriteEndElement();
                    this._writer.WriteEndElement();

                    this._writer.WriteStartElement("a", "bgClr", OpenXmlNamespaces.DrawingML);
                    this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                    if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillBackColor))
                    {
                        colorval = Utils.getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[ShapeOptions.PropertyId.fillBackColor].op, slide, so);
                    }
                    else
                    {
                        colorval = "ffffff"; //TODO: find out which color to use in this case
                    }
                    this._writer.WriteAttributeString("val", colorval);
                    this._writer.WriteEndElement();
                    this._writer.WriteEndElement();

                    this._writer.WriteEndElement();

                    break;
                case 0x2: //texture
                case 0x3: //picture
                    uint blipIndex = 0;
                    string strUrl = "";

                    if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillBlip))
                    {
                        blipIndex = so.OptionsByID[ShapeOptions.PropertyId.fillBlip].op;
                    }
                    else if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.Pib))
                    {
                        blipIndex = so.OptionsByID[ShapeOptions.PropertyId.Pib].op;
                    }
                    else if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillBlipFlags) && so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillBlipName))
                    {
                        uint flags = so.OptionsByID[ShapeOptions.PropertyId.fillBlipFlags].op;
                        bool comment = !Tools.Utils.BitmaskToBool(flags, 0x1);
                        bool file = Tools.Utils.BitmaskToBool(flags, 0x1);
                        bool url = Tools.Utils.BitmaskToBool(flags, 0x1 << 1);
                        bool DoNotSave = Tools.Utils.BitmaskToBool(flags, 0x1 << 2);
                        bool LinkToFile = Tools.Utils.BitmaskToBool(flags, 0x1 << 3);

                        if (url)
                        {
                            strUrl = ASCIIEncoding.ASCII.GetString(so.OptionsByID[ShapeOptions.PropertyId.fillBlipName].opComplex);
                            strUrl = strUrl.Replace("\0", "");
                        }
                    }
                    else
                    {
                        break;
                    }

                    //string blipName = Encoding.UTF8.GetString(so.OptionsByID[ShapeOptions.PropertyId.fillBlipName].opComplex);
                    string rId = "";
                    var gr = (DrawingGroup)this._ctx.Ppt.DocumentRecord.FirstChildWithType<PPDrawingGroup>().Children[0];
                    ImagePart imgPart = null;


                    if (strUrl.Length > 0)
                    {
                        var er = this._parentSlideMapping.targetPart.AddExternalRelationship(OpenXmlRelationshipTypes.Image, strUrl);

                        rId = er.Id;

                        this._writer.WriteStartElement("a", "blipFill", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("dpi", "0");
                        this._writer.WriteAttributeString("rotWithShape", "1");

                        this._writer.WriteStartElement("a", "blip", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("r", "link", OpenXmlNamespaces.Relationships, rId);



                        this._writer.WriteEndElement();

                        this._writer.WriteElementString("a", "srcRect", OpenXmlNamespaces.DrawingML, "");

                        if (fillType == 0x3)
                        {
                            this._writer.WriteStartElement("a", "stretch", OpenXmlNamespaces.DrawingML);
                            this._writer.WriteElementString("a", "fillRect", OpenXmlNamespaces.DrawingML, "");
                            this._writer.WriteEndElement();
                        }
                        else
                        {
                            this._writer.WriteStartElement("a", "tile", OpenXmlNamespaces.DrawingML);
                            this._writer.WriteAttributeString("tx", "0");
                            this._writer.WriteAttributeString("ty", "0");
                            this._writer.WriteAttributeString("sx", "100000");
                            this._writer.WriteAttributeString("sy", "100000");
                            this._writer.WriteAttributeString("flip", "none");
                            this._writer.WriteAttributeString("algn", "tl");
                            this._writer.WriteEndElement();
                        }

                        this._writer.WriteEndElement();

                    } else if (blipIndex <= gr.FirstChildWithType<BlipStoreContainer>().Children.Count)
                    {
                        var bse = (BlipStoreEntry)gr.FirstChildWithType<BlipStoreContainer>().Children[(int)blipIndex - 1];

                        if (this._ctx.Ppt.PicturesContainer._pictures.ContainsKey(bse.foDelay))
                        {
                            var rec = this._ctx.Ppt.PicturesContainer._pictures[bse.foDelay];
                           
                            if (rec is BitmapBlip)
                            {
                                var b = (BitmapBlip)this._ctx.Ppt.PicturesContainer._pictures[bse.foDelay];                                
                                imgPart = this._parentSlideMapping.targetPart.AddImagePart(ShapeTreeMapping.getImageType(b.TypeCode));
                                imgPart.TargetDirectory = "..\\media";
                                var outStream = imgPart.GetStream();
                                outStream.Write(b.m_pvBits, 0, b.m_pvBits.Length);
                            }
                            else
                            {
                                var b = (MetafilePictBlip)this._ctx.Ppt.PicturesContainer._pictures[bse.foDelay];
                                imgPart = this._parentSlideMapping.targetPart.AddImagePart(ShapeTreeMapping.getImageType(b.TypeCode));
                                imgPart.TargetDirectory = "..\\media";
                                var outStream = imgPart.GetStream();
                                var decompressed = b.Decrompress();
                                outStream.Write(decompressed, 0, decompressed.Length);
                            }

                            rId = imgPart.RelIdToString;

                            this._writer.WriteStartElement("a", "blipFill", OpenXmlNamespaces.DrawingML);
                            this._writer.WriteAttributeString("dpi", "0");
                            this._writer.WriteAttributeString("rotWithShape", "1");

                            this._writer.WriteStartElement("a", "blip", OpenXmlNamespaces.DrawingML);
                            this._writer.WriteAttributeString("r", "embed", OpenXmlNamespaces.Relationships, rId);



                            this._writer.WriteEndElement();

                            this._writer.WriteElementString("a", "srcRect", OpenXmlNamespaces.DrawingML, "");

                            if (fillType == 0x3)
                            {
                                this._writer.WriteStartElement("a", "stretch", OpenXmlNamespaces.DrawingML);
                                this._writer.WriteElementString("a", "fillRect", OpenXmlNamespaces.DrawingML, "");
                                this._writer.WriteEndElement();
                            }
                            else
                            {
                                this._writer.WriteStartElement("a", "tile", OpenXmlNamespaces.DrawingML);
                                this._writer.WriteAttributeString("tx", "0");
                                this._writer.WriteAttributeString("ty", "0");
                                this._writer.WriteAttributeString("sx", "100000");
                                this._writer.WriteAttributeString("sy", "100000");
                                this._writer.WriteAttributeString("flip", "none");
                                this._writer.WriteAttributeString("algn", "tl");
                                this._writer.WriteEndElement();
                            }

                            this._writer.WriteEndElement();
                        }
                    }
                    break;
                case 0x4: //shade
                case 0x5: //shadecenter
                case 0x6: //shadeshape
                    this._writer.WriteStartElement("a", "gradFill", OpenXmlNamespaces.DrawingML);
                    this._writer.WriteAttributeString("rotWithShape", "1");
                    this._writer.WriteStartElement("a", "gsLst", OpenXmlNamespaces.DrawingML);
                    bool useFillAndBack = true;

                    if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillShadeColors))
                    {

                        var colors = so.OptionsByID[ShapeOptions.PropertyId.fillShadeColors].opComplex;

                        if (colors != null && colors.Length > 0)
                        {

                            useFillAndBack = false;
                            var type = so.OptionsByID[ShapeOptions.PropertyId.fillShadeType];

                            ushort nElems = System.BitConverter.ToUInt16(colors, 0);
                            ushort nElemsAlloc = System.BitConverter.ToUInt16(colors, 2);
                            ushort cbElem = System.BitConverter.ToUInt16(colors, 4);

                            var positions = new List<string>();

                            switch (nElems)
                            {
                                case 1:
                                case 2:
                                case 3:
                                case 4:
                                case 5:
                                    positions.Add("0");
                                    positions.Add("30000");
                                    positions.Add("65000");
                                    positions.Add("90000");
                                    positions.Add("100000");
                                    break;
                                case 6:
                                case 7:
                                case 8:
                                case 9:
                                case 10:
                                default:
                                    positions.Add("0");
                                    positions.Add("8000");
                                    positions.Add("13000");
                                    positions.Add("21000");
                                    positions.Add("52000");
                                    positions.Add("56000");
                                    positions.Add("58000");
                                    positions.Add("71000");
                                    positions.Add("94000");
                                    positions.Add("100000");
                                    break;
                            }


                            var alphas = new string[nElems];
                            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillOpacity))
                            {
                                decimal end = Math.Round(((decimal)so.OptionsByID[ShapeOptions.PropertyId.fillOpacity].op / 65536 * 100000));
                                decimal start = Math.Round(((decimal)so.OptionsByID[ShapeOptions.PropertyId.fillBackOpacity].op / 65536 * 100000));
                                alphas[0] = start.ToString();
                                for (int i = 1; i < nElems - 1; i++)
                                {
                                    alphas[i] = Math.Round(start + (end - start) / 3 * i).ToString();
                                }
                                //alphas[1] = Math.Round(start + (end - start) / 3).ToString();
                                //alphas[2] = Math.Round(start + (end - start) / 3 * 2).ToString();
                                //alphas[3] = Math.Round(start + (end - start) / 3 * 3).ToString();
                                alphas[nElems - 1] = end.ToString();
                            }

                            for (int i = 0; i < nElems * cbElem; i += cbElem)
                            {
                                colorval = Utils.getRGBColorFromOfficeArtCOLORREF(System.BitConverter.ToUInt32(colors, 6 + i), slide, so);
                                this._writer.WriteStartElement("a", "gs", OpenXmlNamespaces.DrawingML);
                                this._writer.WriteAttributeString("pos", positions[i / cbElem]);

                                this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                                this._writer.WriteAttributeString("val", colorval);
                                if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillOpacity) && so.OptionsByID[ShapeOptions.PropertyId.fillOpacity].op != 65536)
                                {
                                    this._writer.WriteStartElement("a", "alpha", OpenXmlNamespaces.DrawingML);
                                    this._writer.WriteAttributeString("val", alphas[i / cbElem]); //we need the percentage of the opacity (65536 means 100%)
                                    this._writer.WriteEndElement();
                                }
                                this._writer.WriteEndElement();

                                this._writer.WriteEndElement();
                            }
                        }
                    }
                    
                    if (useFillAndBack)
                    {
                        
                        colorval = Utils.getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[ShapeOptions.PropertyId.fillColor].op, slide, so);

                        this._writer.WriteStartElement("a", "gs", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("pos", "0");
                        this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("val", colorval);
                        if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillOpacity) && so.OptionsByID[ShapeOptions.PropertyId.fillOpacity].op != 65536)
                        {
                            this._writer.WriteStartElement("a", "alpha", OpenXmlNamespaces.DrawingML);
                            this._writer.WriteAttributeString("val", Math.Round(((decimal)so.OptionsByID[ShapeOptions.PropertyId.fillOpacity].op / 65536 * 100000)).ToString()); //we need the percentage of the opacity (65536 means 100%)
                            this._writer.WriteEndElement();
                        }
                        this._writer.WriteEndElement();
                        this._writer.WriteEndElement();

                        if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillBackColor))
                        {
                            colorval = Utils.getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[ShapeOptions.PropertyId.fillBackColor].op, slide, so);
                        }
                        else
                        {
                            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.shadowColor))
                            {
                                colorval = Utils.getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[ShapeOptions.PropertyId.shadowColor].op, slide, so);
                            }
                            else
                            {
                                //use filColor
                            }
                        }

                        this._writer.WriteStartElement("a", "gs", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("pos", "100000");
                        this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("val", colorval);
                        if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillBackOpacity) && so.OptionsByID[ShapeOptions.PropertyId.fillBackOpacity].op != 65536)
                        {
                            this._writer.WriteStartElement("a", "alpha", OpenXmlNamespaces.DrawingML);
                            this._writer.WriteAttributeString("val", Math.Round(((decimal)so.OptionsByID[ShapeOptions.PropertyId.fillBackOpacity].op / 65536 * 100000)).ToString()); //we need the percentage of the opacity (65536 means 100%)
                            this._writer.WriteEndElement();
                        }
                        this._writer.WriteEndElement();
                        this._writer.WriteEndElement();
                    }

                    this._writer.WriteEndElement(); //gsLst

                    switch (fillType)
                    {
                        case 0x5:
                        case 0x6:
                            this._writer.WriteStartElement("a", "path", OpenXmlNamespaces.DrawingML);
                            this._writer.WriteAttributeString("path", "shape");
                            this._writer.WriteStartElement("a", "fillToRect", OpenXmlNamespaces.DrawingML);
                            this._writer.WriteAttributeString("l", "50000");
                            this._writer.WriteAttributeString("t", "50000");
                            this._writer.WriteAttributeString("r", "50000");
                            this._writer.WriteAttributeString("b", "50000");
                            this._writer.WriteEndElement();
                            this._writer.WriteEndElement(); //path
                            break;
                        default:
                            this._writer.WriteStartElement("a", "path", OpenXmlNamespaces.DrawingML);
                            this._writer.WriteAttributeString("path", "rect");
                            this._writer.WriteStartElement("a", "fillToRect", OpenXmlNamespaces.DrawingML);
                            this._writer.WriteAttributeString("r", "100000");
                            this._writer.WriteAttributeString("b", "100000");
                            this._writer.WriteEndElement();
                            this._writer.WriteEndElement(); //path
                            break;
                    }

                    this._writer.WriteEndElement(); //gradFill

                    break;
                case 0x7: //shadescale
                    this._writer.WriteStartElement("a", "gradFill", OpenXmlNamespaces.DrawingML);
                    this._writer.WriteAttributeString("rotWithShape", "1");
                    this._writer.WriteStartElement("a", "gsLst", OpenXmlNamespaces.DrawingML);

                    decimal angle = 90;
                    bool switchColors = false;
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
                                switchColors = true;
                            }
                        }
                    }

                    var shadeColorsDic = new Dictionary<int, string>();
                    var shadeColors = new List<string>();
                    if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillShadeColors) && so.OptionsByID[ShapeOptions.PropertyId.fillShadeColors].opComplex != null && so.OptionsByID[ShapeOptions.PropertyId.fillShadeColors].opComplex.Length > 0)
                    {
                        uint length = so.OptionsByID[ShapeOptions.PropertyId.fillShadeColors].op;

                        //An IMsoArray record that specifies colors and their relative positions. 
                        //Each element of the array contains an OfficeArtCOLORREF record color and a FixedPoint, as specified in [MS-OSHARED] 
                        //section 2.2.1.6, that specifies its relative position along the gradient vector.
                        var data = so.OptionsByID[ShapeOptions.PropertyId.fillShadeColors].opComplex;

                        int pos = 0;
                        string colval;
                        FixedPointNumber fixedpoint;
                        ushort nElems = BitConverter.ToUInt16(data, pos);
                        pos += 2;
                        ushort nElemsAlloc = BitConverter.ToUInt16(data, pos);
                        pos += 2;
                        ushort cbElem = BitConverter.ToUInt16(data, pos);
                        pos += 2;

                        if (cbElem == 0xFFF0)
                        {
                            //If this value is 0xFFF0 then this record is an array of truncated 8 byte elements. Only the 4 low-order bytes are recorded. Each element's 4 high-order bytes equal 0x00000000 and each element's 4 low-order bytes are contained in data.
                        }
                        else
                        {
                            while (pos < length)
                            {
                                colval = Utils.getRGBColorFromOfficeArtCOLORREF(BitConverter.ToUInt32(data, pos), slide, so);
                                
                                pos += 4;
                                fixedpoint = new FixedPointNumber(BitConverter.ToUInt16(data, pos), BitConverter.ToUInt16(data, pos + 2));
                                shadeColors.Insert(0,colval);
                                pos += 4;
                            }
                        }
                                           
                    }
                    else
                    {
                        bool switchcolors = false; 
                        if (switchColors & so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillBackColor))
                        {
                            colorval = Utils.getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[ShapeOptions.PropertyId.fillBackColor].op, slide, so);
                        }
                        else
                        {
                            if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillColor))
                            {
                                colorval = Utils.getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[ShapeOptions.PropertyId.fillColor].op, slide, so);
                            }
                            else
                            {
                                colorval = "FFFFFF"; //TODO: find out which color to use in this case
                                switchcolors = true;
                            }
                        }

                        if (switchColors | !so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillBackColor))
                        {
                            colorval2 = Utils.getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[ShapeOptions.PropertyId.fillColor].op, slide, so);
                        }
                        else
                        {
                            colorval2 = Utils.getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[ShapeOptions.PropertyId.fillBackColor].op, slide, so);
                        }

                        if (switchcolors)
                        {
                            //this is a workaround for a bug. Further analysis necessarry
                            string dummy = colorval;
                            colorval = colorval2;
                            colorval2 = dummy;
                        }

                        shadeColors.Add(colorval);
                        shadeColors.Add(colorval2);
                    }
                                       

                    int gspos;
                    string col;
                    for (int i = 0; i < shadeColors.Count; i++)
                    {
                        col = shadeColors[i];
                        if (i == 0)
                        {
                            gspos = 0;
                        } else if (i == shadeColors.Count-1)
                        {
                            gspos = 100000;
                        } else {
                            gspos = i * 100000 / shadeColors.Count;
                        }

                        this._writer.WriteStartElement("a", "gs", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("pos", gspos.ToString());
                        this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("val", col);
                        if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillOpacity) && so.OptionsByID[ShapeOptions.PropertyId.fillOpacity].op != 65536)
                        {
                            this._writer.WriteStartElement("a", "alpha", OpenXmlNamespaces.DrawingML);
                            this._writer.WriteAttributeString("val", Math.Round(((decimal)so.OptionsByID[ShapeOptions.PropertyId.fillOpacity].op / 65536 * 100000)).ToString()); //we need the percentage of the opacity (65536 means 100%)
                            this._writer.WriteEndElement();
                        }

                        if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillShadeType))
                        {
                            uint flags = so.OptionsByID[ShapeOptions.PropertyId.fillShadeType].op;
                            bool none = Tools.Utils.BitmaskToBool(flags, 0x1);
                            bool gamma = Tools.Utils.BitmaskToBool(flags, 0x1 << 1);
                            bool sigma = Tools.Utils.BitmaskToBool(flags, 0x1 << 2);
                            bool band = Tools.Utils.BitmaskToBool(flags, 0x1 << 3);
                            bool onecolor = Tools.Utils.BitmaskToBool(flags, 0x1 << 4);

                            if (gamma) this._writer.WriteElementString("a", "gamma", OpenXmlNamespaces.DrawingML, "");
                            if (band)
                            {
                                this._writer.WriteStartElement("a", "shade", OpenXmlNamespaces.DrawingML);
                                this._writer.WriteAttributeString("val", "37255");
                                this._writer.WriteEndElement();
                            }
                            if (gamma) this._writer.WriteElementString("a", "invGamma", OpenXmlNamespaces.DrawingML, "");
                        }
                        this._writer.WriteEndElement();
                        this._writer.WriteEndElement();
                    }




                    ////new colorval
                    //_writer.WriteStartElement("a", "gs", OpenXmlNamespaces.DrawingML);
                    //_writer.WriteAttributeString("pos", "100000");
                    //_writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                    //_writer.WriteAttributeString("val", colorval2);
                    //if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillBackOpacity))
                    //{
                    //    _writer.WriteStartElement("a", "alpha", OpenXmlNamespaces.DrawingML);
                    //    _writer.WriteAttributeString("val", Math.Round(((decimal)so.OptionsByID[ShapeOptions.PropertyId.fillBackOpacity].op / 65536 * 100000)).ToString()); //we need the percentage of the opacity (65536 means 100%)
                    //    _writer.WriteEndElement();
                    //}

                    //_writer.WriteEndElement();
                    //_writer.WriteEndElement();

                    this._writer.WriteEndElement(); //gsLst

                    this._writer.WriteStartElement("a", "lin", OpenXmlNamespaces.DrawingML);

                    angle *= 60000;
                    //if (angle > 5400000) angle = 5400000;

                    this._writer.WriteAttributeString("ang", angle.ToString());
                    this._writer.WriteAttributeString("scaled", "1");
                    this._writer.WriteEndElement();

                    this._writer.WriteEndElement();
                    break;
                case 0x8: //shadetitle
                case 0x9: //background
                    break;

            }
        }
    }
}
