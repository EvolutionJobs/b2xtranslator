using System;
using System.Collections.Generic;
using System.Text;
using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.OfficeDrawing;
using System.Xml;
using b2xtranslator.OpenXmlLib;
using System.IO;
using b2xtranslator.DocFileFormat;
using b2xtranslator.Tools;
using System.Globalization;
using b2xtranslator.OfficeDrawing.Shapetypes;

namespace b2xtranslator.WordprocessingMLMapping
{
    public class VMLShapeMapping: PropertiesMapping,
          IMapping<ShapeContainer>
    {
        private BlipStoreContainer _blipStore = null;
        private ConversionContext _ctx;
        private FileShapeAddress _fspa;
        private PictureDescriptor _pict;
        private ContentPart _targetPart;
        private XmlElement _fill, _stroke, _shadow, _imagedata, _3dstyle, _textpath;
        private static GroupShapeRecord _groupShapeRecord;
        private List<byte> pSegmentInfo = new List<byte>();
        private List<byte> pVertices = new List<byte>();
        private StringBuilder _textPathStyle;

        public VMLShapeMapping(XmlWriter writer, 
            ContentPart targetPart, 
            FileShapeAddress fspa, 
            PictureDescriptor pict,
            ConversionContext ctx)
            : base(writer)
        {
            this._ctx = ctx;
            this._fspa = fspa;
            this._pict = pict;
            this._targetPart = targetPart;
            this._imagedata = this._nodeFactory.CreateElement("v", "imagedata", OpenXmlNamespaces.VectorML);
            this._fill = this._nodeFactory.CreateElement("v", "fill", OpenXmlNamespaces.VectorML);
            this._stroke = this._nodeFactory.CreateElement("v", "stroke", OpenXmlNamespaces.VectorML);
            this._shadow = this._nodeFactory.CreateElement("v", "shadow", OpenXmlNamespaces.VectorML);
            this._3dstyle = this._nodeFactory.CreateElement("o", "extrusion", OpenXmlNamespaces.Office);
            this._textpath = this._nodeFactory.CreateElement("v", "textpath", OpenXmlNamespaces.VectorML);

            this._textPathStyle = new StringBuilder();

            Record recBs = this._ctx.Doc.OfficeArtContent.DrawingGroupData.FirstChildWithType<BlipStoreContainer>();
            if (recBs != null)
            {
                this._blipStore = (BlipStoreContainer)recBs;
            }
        }


        public void Apply(ShapeContainer container)
        {
            var firstRecord = container.Children[0];
            if (firstRecord.GetType() == typeof(Shape))
            {
                //It's a single shape
                convertShape(container);
            }
            else if (firstRecord.GetType() == typeof(GroupShapeRecord))
            { 
                //Its a group of shapes
                convertGroup((GroupContainer)container.ParentRecord);
            }

            this._writer.Flush();
        }


        /// <summary>
        /// Converts a group of shapes
        /// </summary>
        /// <param name="container"></param>
        private void convertGroup(GroupContainer container)
        {
            var groupShape = (ShapeContainer)container.Children[0];
            _groupShapeRecord = (GroupShapeRecord)groupShape.Children[0];
            var shape = (Shape)groupShape.Children[1];
            var options = groupShape.ExtractOptions();
            var anchor = groupShape.FirstChildWithType<ChildAnchor>();

            this._writer.WriteStartElement("v", "group", OpenXmlNamespaces.VectorML);
            this._writer.WriteAttributeString("id", getShapeId(shape));
            this._writer.WriteAttributeString("style", buildStyle(shape, anchor, options, container.Index).ToString());
            this._writer.WriteAttributeString("coordorigin", _groupShapeRecord.rcgBounds.Left + "," + _groupShapeRecord.rcgBounds.Top);
            this._writer.WriteAttributeString("coordsize", _groupShapeRecord.rcgBounds.Width + "," + _groupShapeRecord.rcgBounds.Height);
            
            //write wrap coords
            foreach (var entry in options)
            {
                switch (entry.pid)
                {
                    case ShapeOptions.PropertyId.pWrapPolygonVertices:
                        this._writer.WriteAttributeString("wrapcoords", getWrapCoords(entry));
                        break;
                }
            }

            //convert the shapes/groups in the group
            for (int i = 1; i < container.Children.Count; i++)
            {
                if (container.Children[i].GetType() == typeof(ShapeContainer))
                {
                    var childShape = (ShapeContainer)container.Children[i];
                    childShape.Convert(new VMLShapeMapping(this._writer, this._targetPart, this._fspa, null, this._ctx));
                }
                else if (container.Children[i].GetType() == typeof(GroupContainer))
                {
                    var childGroup = (GroupContainer)container.Children[i];
                    this._fspa = null;
                    convertGroup(childGroup);
                }
            }

            //write wrap
            if (this._fspa != null)
            {
                string wrap = getWrapType(this._fspa);
                if(wrap != "through")
                {
                    this._writer.WriteStartElement("w10", "wrap", OpenXmlNamespaces.OfficeWord);
                    this._writer.WriteAttributeString("type", wrap);
                    this._writer.WriteEndElement();
                }
            }

            this._writer.WriteEndElement();
        }


        /// <summary>
        /// Converts a single shape
        /// </summary>
        /// <param name="container"></param>
        private void convertShape(ShapeContainer container)
        {
            var shape = (Shape)container.Children[0];
            var options = container.ExtractOptions();
            var anchor = container.FirstChildWithType<ChildAnchor>();
            var clientAnchor = container.FirstChildWithType<ClientAnchor>();

            writeStartShapeElement(shape);
            this._writer.WriteAttributeString("id", getShapeId(shape));
            if (shape.ShapeType != null)
            {
                this._writer.WriteAttributeString("type", "#" + VMLShapeTypeMapping.GenerateTypeId(shape.ShapeType));
            }
            this._writer.WriteAttributeString("style", buildStyle(shape, anchor, options, container.Index).ToString());
            if (shape.ShapeType is LineType)
            {
                //append "from" and  "to" attributes
                this._writer.WriteAttributeString("from", getCoordinateFrom(anchor));
                this._writer.WriteAttributeString("to", getCoordinateTo(anchor));
            }

            //temporary variables
            EmuValue shadowOffsetX = null;
            EmuValue shadowOffsetY = null;
            EmuValue secondShadowOffsetX = null;
            EmuValue secondShadowOffsetY = null;
            double shadowOriginX = 0;
            double shadowOriginY = 0;
            EmuValue viewPointX = null;
            EmuValue viewPointY = null;
            EmuValue viewPointZ = null;
            double? viewPointOriginX = null;
            double? viewPointOriginY = null;
            var adjValues = new string[8];
            int numberAdjValues = 0;
            uint xCoord = 0;
            uint yCoord = 0;
            bool stroked = true;
            bool filled = true;
            bool hasTextbox = false;

            foreach (var entry in options)
            {
                switch (entry.pid)
                {
                    //BOOLEANS

                    case ShapeOptions.PropertyId.geometryBooleans:
                        var geometryBooleans = new GeometryBooleans(entry.op);

                        if (geometryBooleans.fUsefLineOK && geometryBooleans.fLineOK==false)
                        {
                            stroked = false;
                        }

                        if (!(geometryBooleans.fUsefFillOK && geometryBooleans.fFillOK))
                        {
                            filled = false;
                        }
                        break;

                    case ShapeOptions.PropertyId.FillStyleBooleanProperties:
                        var fillBooleans = new FillStyleBooleanProperties(entry.op);

                        if (fillBooleans.fUsefFilled && fillBooleans.fFilled == false)
                        {
                            filled = false;
                        }
                        break;

                    case ShapeOptions.PropertyId.lineStyleBooleans:
                        var lineBooleans = new LineStyleBooleans(entry.op);

                        if (lineBooleans.fUsefLine && lineBooleans.fLine == false)
                        {
                            stroked = false;
                        }

                        break;

                    case ShapeOptions.PropertyId.protectionBooleans:
                        var protBools = new ProtectionBooleans(entry.op);

                        break;

                    case ShapeOptions.PropertyId.diagramBooleans:
                        var diaBools = new DiagramBooleans(entry.op);

                        break;

                    // GEOMETRY

                    case ShapeOptions.PropertyId.adjustValue:
                        adjValues[0] = (((int)entry.op).ToString());
                        numberAdjValues++; 
                        break;

                    case ShapeOptions.PropertyId.adjust2Value:
                        adjValues[1] = (((int)entry.op).ToString());
                        numberAdjValues++;
                        break;

                    case ShapeOptions.PropertyId.adjust3Value:
                        adjValues[2] = (((int)entry.op).ToString());
                        numberAdjValues++;
                        break;

                    case ShapeOptions.PropertyId.adjust4Value:
                        adjValues[3] = (((int)entry.op).ToString());
                        numberAdjValues++;
                        break;

                    case ShapeOptions.PropertyId.adjust5Value:
                        adjValues[4] = (((int)entry.op).ToString());
                        numberAdjValues++;
                        break;

                    case ShapeOptions.PropertyId.adjust6Value:
                        adjValues[5] = (((int)entry.op).ToString());
                        numberAdjValues++;
                        break;

                    case ShapeOptions.PropertyId.adjust7Value:
                        adjValues[6] = (((int)entry.op).ToString());
                        numberAdjValues++;
                        break;

                    case ShapeOptions.PropertyId.adjust8Value:
                        adjValues[7] = (((int)entry.op).ToString());
                        numberAdjValues++;
                        break;

                    case ShapeOptions.PropertyId.pWrapPolygonVertices:
                        this._writer.WriteAttributeString("wrapcoords", getWrapCoords(entry));
                        break;

                    case ShapeOptions.PropertyId.geoRight:
                        xCoord = entry.op;
                        break;

                    case ShapeOptions.PropertyId.geoBottom:
                        yCoord = entry.op;
                        break;

                    // OUTLINE

                    case ShapeOptions.PropertyId.lineColor:
                        var lineColor = new RGBColor((int)entry.op, RGBColor.ByteOrder.RedFirst);
                        this._writer.WriteAttributeString("strokecolor", "#" + lineColor.SixDigitHexCode);
                        break;

                    case ShapeOptions.PropertyId.lineWidth:
                        var lineWidth = new EmuValue((int)entry.op);
                        this._writer.WriteAttributeString("strokeweight", lineWidth.ToString());
                        break;

                    case ShapeOptions.PropertyId.lineDashing:
                        var dash = (Global.DashStyle)entry.op;
                        appendValueAttribute(this._stroke, null, "dashstyle", dash.ToString(), null);
                        break;

                    case ShapeOptions.PropertyId.lineStyle:
                        appendValueAttribute(this._stroke, null, "linestyle", getLineStyle(entry.op), null);
                        break;

                    case ShapeOptions.PropertyId.lineEndArrowhead:
                        appendValueAttribute(this._stroke, null, "endarrow", getArrowStyle(entry.op), null);
                        break;

                    case ShapeOptions.PropertyId.lineEndArrowLength:
                        appendValueAttribute(this._stroke, null, "endarrowlength", getArrowLength(entry.op), null);
                        break;

                    case ShapeOptions.PropertyId.lineEndArrowWidth:
                        appendValueAttribute(this._stroke, null, "endarrowwidth", getArrowWidth(entry.op), null);
                        break;

                    case ShapeOptions.PropertyId.lineStartArrowhead:
                        appendValueAttribute(this._stroke, null, "startarrow", getArrowStyle(entry.op), null);
                        break;

                    case ShapeOptions.PropertyId.lineStartArrowLength:
                        appendValueAttribute(this._stroke, null, "startarrowlength", getArrowLength(entry.op), null);
                        break;

                    case ShapeOptions.PropertyId.lineStartArrowWidth:
                        appendValueAttribute(this._stroke, null, "startarrowwidth", getArrowWidth(entry.op), null);
                        break;


                    // FILL

                    case ShapeOptions.PropertyId.fillColor:
                        var fillColor = new RGBColor((int)entry.op, RGBColor.ByteOrder.RedFirst);
                        this._writer.WriteAttributeString("fillcolor", "#" + fillColor.SixDigitHexCode);
                        break;

                    case ShapeOptions.PropertyId.fillBackColor:
                        var fillBackColor = new RGBColor((int)entry.op, RGBColor.ByteOrder.RedFirst);
                        appendValueAttribute(this._fill, null, "color2", "#" + fillBackColor.SixDigitHexCode, null);
                        break;

                    case ShapeOptions.PropertyId.fillAngle:
                        var fllAngl = new FixedPointNumber(entry.op);
                        appendValueAttribute(this._fill, null, "angle", fllAngl.ToAngle().ToString(), null);
                        break;

                    case ShapeOptions.PropertyId.fillShadeType:
                        appendValueAttribute(this._fill, null, "method", getFillMethod(entry.op), null);
                        break;

                    case ShapeOptions.PropertyId.fillShadeColors:
                        appendValueAttribute(this._fill, null, "colors", getFillColorString(entry.opComplex), null);
                        break;

                    case ShapeOptions.PropertyId.fillFocus:
                        appendValueAttribute(this._fill, null, "focus", entry.op + "%", null);
                        break;

                    case ShapeOptions.PropertyId.fillType:
                        appendValueAttribute(this._fill, null, "type", getFillType(entry.op), null);
                        break;

                    case ShapeOptions.PropertyId.fillBlip:
                        ImagePart fillBlipPart = null;
                        if (this._pict != null && this._pict.BlipStoreEntry != null)
                        {
                            // Word Art Texture
                            //fillBlipPart = copyPicture(_pict.BlipStoreEntry);
                        }
                        else
                        {
                            var fillBlip = (BlipStoreEntry)this._blipStore.Children[(int)entry.op - 1];
                            fillBlipPart = copyPicture(fillBlip);
                        }
                        if (fillBlipPart != null)
                        {
                            appendValueAttribute(this._fill, "r", "id", fillBlipPart.RelIdToString, OpenXmlNamespaces.Relationships);
                            appendValueAttribute(this._imagedata, "o", "title", "", OpenXmlNamespaces.Office);
                        }
                        break;

                    case ShapeOptions.PropertyId.fillOpacity:
                        appendValueAttribute(this._fill, null, "opacity", entry.op + "f" , null);
                        break;

                    // SHADOW

                    case ShapeOptions.PropertyId.shadowType:
                        appendValueAttribute(this._shadow, null, "type", getShadowType(entry.op), null);
                        break;

                    case ShapeOptions.PropertyId.shadowColor:
                        var shadowColor = new RGBColor((int)entry.op, RGBColor.ByteOrder.RedFirst);
                        appendValueAttribute(this._shadow, null, "color", "#" + shadowColor.SixDigitHexCode, null);
                        break;

                    case ShapeOptions.PropertyId.shadowOffsetX:
                        shadowOffsetX = new EmuValue((int)entry.op);
                        break;

                    case ShapeOptions.PropertyId.shadowSecondOffsetX:
                        secondShadowOffsetX = new EmuValue((int)entry.op);
                        break;

                    case ShapeOptions.PropertyId.shadowOffsetY:
                        shadowOffsetY = new EmuValue((int)entry.op);
                        break;

                    case ShapeOptions.PropertyId.shadowSecondOffsetY:
                        secondShadowOffsetY = new EmuValue((int)entry.op);
                        break;

                    case ShapeOptions.PropertyId.shadowOriginX:
                        shadowOriginX = entry.op / Math.Pow(2, 16);
                        break;

                    case ShapeOptions.PropertyId.shadowOriginY:
                        shadowOriginY = entry.op / Math.Pow(2, 16);
                        break;

                    case ShapeOptions.PropertyId.shadowOpacity:
                        double shadowOpa = (entry.op / Math.Pow(2, 16));
                        appendValueAttribute(this._shadow, null, "opacity", string.Format(CultureInfo.CreateSpecificCulture("EN"), "{0:0.00}", shadowOpa), null);
                        break;

                    // PICTURE
                    
                    case ShapeOptions.PropertyId.Pib:
                        int index = (int)entry.op - 1;
                        var bse = (BlipStoreEntry)this._blipStore.Children[index];
                        var part = copyPicture(bse);
                        if (part != null)
                        {
                            appendValueAttribute(this._imagedata, "r", "id", part.RelIdToString, OpenXmlNamespaces.Relationships);
                        }
                        break;

                    case ShapeOptions.PropertyId.pibName:
                        string name = Encoding.Unicode.GetString(entry.opComplex);
                        name = Utils.GetWritableString(name.Substring(0, name.Length - 1));
                        appendValueAttribute(this._imagedata, "o", "title", name, OpenXmlNamespaces.Office);
                        break;

                    // 3D STYLE

                    case ShapeOptions.PropertyId.f3D:
                    case ShapeOptions.PropertyId.ThreeDStyleBooleanProperties:
                    case ShapeOptions.PropertyId.ThreeDObjectBooleanProperties:
                        break;
                    case ShapeOptions.PropertyId.c3DExtrudeBackward:
                        var backwardValue = new EmuValue((int)entry.op);
                        appendValueAttribute(this._3dstyle, "backdepth", backwardValue.ToPoints().ToString());
                        break; 
                    case ShapeOptions.PropertyId.c3DSkewAngle:
                        var skewAngle = new FixedPointNumber(entry.op);
                        appendValueAttribute(this._3dstyle, "", "skewangle", skewAngle.ToAngle().ToString(), "");
                        break;
                    case ShapeOptions.PropertyId.c3DXViewpoint:
                        viewPointX = new EmuValue(new FixedPointNumber(entry.op).Integral);
                        break;
                    case ShapeOptions.PropertyId.c3DYViewpoint:
                        viewPointY = new EmuValue(new FixedPointNumber(entry.op).Integral);
                        break;
                    case ShapeOptions.PropertyId.c3DZViewpoint:
                        viewPointZ = new EmuValue(new FixedPointNumber(entry.op).Integral);
                        break;
                    case ShapeOptions.PropertyId.c3DOriginX:
                        var dOriginX = new FixedPointNumber(entry.op);
                        viewPointOriginX = dOriginX.Integral / 65536.0;
                        break;
                    case ShapeOptions.PropertyId.c3DOriginY:
                        var dOriginY = new FixedPointNumber(entry.op);
                        break;

                    // TEXTBOX

                    case ShapeOptions.PropertyId.lTxid:
                        hasTextbox = true;
                        break;

                    // TEXT PATH (Word Art)

                    case ShapeOptions.PropertyId.gtextUNICODE:
                        string text = Encoding.Unicode.GetString(entry.opComplex);
                        text = text.Replace("\n", "");
                        text = text.Replace("\0", "");
                        appendValueAttribute(this._textpath, "", "string", text, "");
                        break;
                    case ShapeOptions.PropertyId.gtextFont:
                        string font = Encoding.Unicode.GetString(entry.opComplex);
                        font = font.Replace("\0", "");
                        appendStyleProperty(this._textPathStyle, "font-family", "\""+font+"\"");
                        break;
                    case ShapeOptions.PropertyId.GeometryTextBooleanProperties:
                        var props = new GeometryTextBooleanProperties(entry.op);
                        if (props.fUsegtextFBestFit && props.gtextFBestFit)
                        {
                            appendValueAttribute(this._textpath, "", "fitshape", "t", "");
                        }
                        if (props.fUsegtextFShrinkFit && props.gtextFShrinkFit)
                        {
                            appendValueAttribute(this._textpath, "", "trim", "t", "");
                        }
                        if (props.fUsegtextFVertical && props.gtextFVertical)
                        {
                            appendStyleProperty(this._textPathStyle, "v-rotate-letters", "t");
                            //_twistDimension = true;
                        }
                        if (props.fUsegtextFKern && props.gtextFKern)
                        {
                            appendStyleProperty(this._textPathStyle, "v-text-kern", "t");
                        }
                        if (props.fUsegtextFItalic && props.gtextFItalic)
                        {
                            appendStyleProperty(this._textPathStyle, "font-style", "italic");
                        }
                        if (props.fUsegtextFBold && props.gtextFBold)
                        {
                            appendStyleProperty(this._textPathStyle, "font-weight", "bold");
                        }
                        break;
                    

                    // PATH
                    case ShapeOptions.PropertyId.shapePath:
                        string path = parsePath(options);
                        if(!string.IsNullOrEmpty(path))
                        {
                            this._writer.WriteAttributeString("path", path);
                        }
                        break;
                }
            }

            if (!filled)
            {
                this._writer.WriteAttributeString("filled", "f");
            }

            if (!stroked)
            {
                this._writer.WriteAttributeString("stroked", "f");
            }

            if (xCoord > 0 && yCoord > 0)
            {
                this._writer.WriteAttributeString("coordsize", xCoord + "," + yCoord);
            }

            //write adj values 
            if (numberAdjValues != 0)
            {
                string adjString = adjValues[0];
                for (int i = 1; i < 8; i++)
                {
                    adjString += "," + adjValues[i];
                }
                this._writer.WriteAttributeString("adj", adjString);
                //string.Format("{0:x4}", adjValues);
            }

            //build shadow offsets
            var offset = new StringBuilder();
            if (shadowOffsetX != null)
            {
                offset.Append(shadowOffsetX.ToPoints());
                offset.Append("pt");
            }
            if (shadowOffsetY != null)
            {
                offset.Append(",");
                offset.Append(shadowOffsetY.ToPoints());
                offset.Append("pt");
            }
            if (offset.Length > 0)
            {
                appendValueAttribute(this._shadow, null, "offset", offset.ToString(), null);
            }
            var offset2 = new StringBuilder();
            if (secondShadowOffsetX != null)
            {
                offset2.Append(secondShadowOffsetX.ToPoints());
                offset2.Append("pt");
            }
            if (secondShadowOffsetY != null)
            {
                offset2.Append(",");
                offset2.Append(secondShadowOffsetY.ToPoints());
                offset2.Append("pt");
            }
            if (offset2.Length > 0)
            {
                appendValueAttribute(this._shadow, null, "offset2", offset2.ToString(), null);
            }

            //build shadow origin
            if (shadowOriginX != 0 && shadowOriginY != 0)
            {
                appendValueAttribute(
                    this._shadow, null, "origin",
                    shadowOriginX + "," + shadowOriginY,
                    null);
            }

            //write shadow
            if (this._shadow.Attributes.Count > 0)
            {
                appendValueAttribute(this._shadow, null, "on", "t", null);
                this._shadow.WriteTo(this._writer);
            }

            //write 3d style 
            if (this._3dstyle.Attributes.Count > 0)
            {
                appendValueAttribute(this._3dstyle, "v", "ext", "view", OpenXmlNamespaces.VectorML);
                appendValueAttribute(this._3dstyle, null, "on", "t", null);

                //write the viewpoint
                if (viewPointX != null || viewPointY != null || viewPointZ != null)
                {
                    var viewPoint = new StringBuilder();
                    if (viewPointX != null)
                    {
                        viewPoint.Append(viewPointX.Value);
                    }
                    if (viewPointY != null)
                    {
                        viewPoint.Append(",");
                        viewPoint.Append(viewPointY.Value);
                    }
                    
                    if (viewPointZ != null)
                    {
                        viewPoint.Append(",");
                        viewPoint.Append(viewPointZ.Value);
                    }
                    appendValueAttribute(this._3dstyle, null, "viewpoint", viewPoint.ToString(), null);
                }

                // write the viewpointorigin
                if (viewPointOriginX != null || viewPointOriginY != null)
                {
                    var viewPointOrigin = new StringBuilder();
                    if (viewPointOriginX != null)
                    {
                        viewPointOrigin.Append(string.Format(CultureInfo.CreateSpecificCulture("EN"), "{0:0.00}", viewPointOriginX));
                    }
                    if (viewPointOriginY != null)
                    {
                        viewPointOrigin.Append(",");
                        viewPointOrigin.Append(string.Format(CultureInfo.CreateSpecificCulture("EN"), "{0:0.00}", viewPointOriginY));
                    }
                    appendValueAttribute(this._3dstyle, null, "viewpointorigin", viewPointOrigin.ToString(), null);
                }

                this._3dstyle.WriteTo(this._writer);
            }

            //write wrap
            if (this._fspa != null)
            {
                string wrap = getWrapType(this._fspa);
                if(wrap != "through")
                {
                    this._writer.WriteStartElement("w10", "wrap", OpenXmlNamespaces.OfficeWord);
                    this._writer.WriteAttributeString("type", wrap);
                    this._writer.WriteEndElement();
                }
            }

            //write stroke
            if (this._stroke.Attributes.Count > 0)
            {
                this._stroke.WriteTo(this._writer);
            }

            //write fill
            if (this._fill.Attributes.Count > 0)
            {
                this._fill.WriteTo(this._writer);
            }

            // text path
            if (this._textpath.Attributes.Count > 0)
            {
                appendValueAttribute(this._textpath, "", "style", this._textPathStyle.ToString(), "");
                this._textpath.WriteTo(this._writer);
            }

            //write imagedata
            if (this._imagedata.Attributes.Count > 0)
            {
                this._imagedata.WriteTo(this._writer);
            }

            //write the textbox
            Record recTextbox = container.FirstChildWithType<ClientTextbox>();
            if (recTextbox != null)
            {
                //Word text box

                //Word appends a ClientTextbox record to the container. 
                //This record stores the index of the textbox.

                var box = (ClientTextbox)recTextbox;
                short textboxIndex = System.BitConverter.ToInt16(box.Bytes, 2);
                textboxIndex--;
                this._ctx.Doc.Convert(new TextboxMapping(this._ctx, textboxIndex, this._targetPart, this._writer));
            }
            else if(hasTextbox)
            {
                //Open Office textbox

                //Open Office doesn't append a ClientTextbox record to the container.
                //We don't know how Word gets the relation to the text, but we assume that the first textbox in the document
                //get the index 0, the second textbox gets the index 1 (and so on).

                this._ctx.Doc.Convert(new TextboxMapping(this._ctx, this._targetPart, this._writer));
            }

            //write the shape
            this._writer.WriteEndElement();
            this._writer.Flush();
        }

        private string getFillColorString(byte[] p)
        {
            var result = new StringBuilder();

            // parse the IMsoArray
            ushort nElems = System.BitConverter.ToUInt16(p, 0);
            ushort nElemsAlloc = System.BitConverter.ToUInt16(p, 2);
            ushort cbElem = System.BitConverter.ToUInt16(p, 4);
            for (int i = 0; i < nElems; i++)
            {
                int pos = 6 + i * cbElem;

                var color = new RGBColor(System.BitConverter.ToInt32(p, pos), RGBColor.ByteOrder.RedFirst);
                int colorPos = System.BitConverter.ToInt32(p, pos + 4);

                result.Append(colorPos);
                result.Append("f #");
                result.Append(color.SixDigitHexCode);
                result.Append(";");
            }

            return result.ToString();
        }

        private string parsePath(List<ShapeOptions.OptionEntry> options)
        {
            string path = "";
            byte[] pVertices = null;
            byte[] pSegmentInfo = null;

            foreach (var e in options)
            {
                if (e.pid == ShapeOptions.PropertyId.pVertices)
                {
                    pVertices = e.opComplex;
                }
                else if (e.pid == ShapeOptions.PropertyId.pSegmentInfo)
                {
                    pSegmentInfo = e.opComplex;
                }
            }

            if (pSegmentInfo != null && pVertices != null)
            {
                var parser = new PathParser(pSegmentInfo, pVertices);
                path = buildVmlPath(parser);
            }
            return path;
        }

        private string buildVmlPath(PathParser parser)
        {
            // build the VML Path
            var VmlPath = new StringBuilder();
            int valuePointer = 0;
            foreach (var seg in parser.Segments)
            {
                try
                {
                    switch (seg.Type)
                    {
                        case PathSegment.SegmentType.msopathLineTo:
                            VmlPath.Append("l");
                            VmlPath.Append(parser.Values[valuePointer].X);
                            VmlPath.Append(",");
                            VmlPath.Append(parser.Values[valuePointer].Y);
                            valuePointer += 1;
                            break;
                        case PathSegment.SegmentType.msopathCurveTo:
                            VmlPath.Append("c");
                            VmlPath.Append(parser.Values[valuePointer].X);
                            VmlPath.Append(",");
                            VmlPath.Append(parser.Values[valuePointer].Y);
                            VmlPath.Append(",");
                            VmlPath.Append(parser.Values[valuePointer + 1].X);
                            VmlPath.Append(",");
                            VmlPath.Append(parser.Values[valuePointer + 1].Y);
                            VmlPath.Append(",");
                            VmlPath.Append(parser.Values[valuePointer + 2].X);
                            VmlPath.Append(",");
                            VmlPath.Append(parser.Values[valuePointer + 2].Y);
                            valuePointer += 3;
                            break;
                        case PathSegment.SegmentType.msopathMoveTo:
                            VmlPath.Append("m");
                            VmlPath.Append(parser.Values[valuePointer].X);
                            VmlPath.Append(",");
                            VmlPath.Append(parser.Values[valuePointer].Y);
                            valuePointer += 1;
                            break;
                        case PathSegment.SegmentType.msopathClose:
                            VmlPath.Append("x");
                            break;
                        case PathSegment.SegmentType.msopathEnd:
                            VmlPath.Append("e");
                            break;
                        case PathSegment.SegmentType.msopathEscape:
                        case PathSegment.SegmentType.msopathClientEscape:
                        case PathSegment.SegmentType.msopathInvalid:
                            //ignore escape segments and invalid segments
                            break;

                    }
                }
                catch (IndexOutOfRangeException)
                {
                    // Sometimes there are more Segments than available Values.
                    // Accordingly to the spec this should never happen :)
                    break;
                }
            }

            // end the path
            if (VmlPath[VmlPath.Length-1] != 'e')
            {
                VmlPath.Append("e");
            }

            return VmlPath.ToString();
        }

        private string getCoordinateFrom(ChildAnchor anchor)
        {
            var from = new StringBuilder();
            if (this._fspa != null)
            {
                var left = new TwipsValue(this._fspa.xaLeft);
                var top = new TwipsValue(this._fspa.yaTop);

                from.Append(left.ToPoints().ToString(CultureInfo.GetCultureInfo("en-US")));
                from.Append("pt,");
                from.Append(top.ToPoints().ToString(CultureInfo.GetCultureInfo("en-US")));
                from.Append("pt");
            }
            else
            {
                from.Append(anchor.rcgBounds.Left);
                from.Append("pt,");
                from.Append(anchor.rcgBounds.Top);
                from.Append("pt");
            }
            return from.ToString();
        }

        private string getCoordinateTo(ChildAnchor anchor)
        {
            var from = new StringBuilder();
            if (this._fspa != null)
            {
                var right = new TwipsValue(this._fspa.xaRight);
                var bottom = new TwipsValue(this._fspa.yaBottom);

                from.Append(right.ToPoints().ToString(CultureInfo.GetCultureInfo("en-US")));
                from.Append("pt,");
                from.Append(bottom.ToPoints().ToString(CultureInfo.GetCultureInfo("en-US")));
                from.Append("pt");
            }
            else
            {
                from.Append(anchor.rcgBounds.Right);
                from.Append("pt,");
                from.Append(anchor.rcgBounds.Bottom);
                from.Append("pt");
            }
            return from.ToString();
        }

        private StringBuilder buildStyle(Shape shape, ChildAnchor anchor, List<ShapeOptions.OptionEntry> options, int zIndex)
        {
            var style = new StringBuilder();

            // Check if some properties are set that cause the dimensions to be twisted
            bool twistDimensions = false;
            foreach (var entry in options)
            {
                if (entry.pid == ShapeOptions.PropertyId.GeometryTextBooleanProperties)
                {
                    var props = new GeometryTextBooleanProperties(entry.op);
                    if (props.fUsegtextFVertical && props.gtextFVertical)
                    {
                        twistDimensions = true;
                    }
                }
            }

            //don't append the dimension info to lines, 
            // because they have "from" and "to" attributes to decline the dimension
            if(!(shape.ShapeType is LineType))
            {
                if (shape.fChild == false && this._fspa != null)
                {
                    //this shape is placed directly in the document, 
                    //so use the FSPA to build the style
                    AppendDimensionToStyle(style, this._fspa, twistDimensions);
                }
                else if(anchor != null)
                {
                    //the style is part of a group, 
                    //so use the anchor
                    AppendDimensionToStyle(style, anchor, twistDimensions);
                }
                else if(this._pict != null)
                {
                    // it is some kind of PICT shape (e.g. WordArt)
                    AppendDimensionToStyle(style, this._pict, twistDimensions);
                }
            }

            if (shape.fFlipH)
            {
                appendStyleProperty(style, "flip", "x");
            }
            if (shape.fFlipV)
            {
                appendStyleProperty(style, "flip", "y");
            }

            AppendOptionsToStyle(style, options);

            appendStyleProperty(style, "z-index", zIndex.ToString());

            return style;
        }

        private void writeStartShapeElement(Shape shape)
        {
            if (shape.ShapeType is OvalType)
            {
                //OVAL
                this._writer.WriteStartElement("v", "oval", OpenXmlNamespaces.VectorML);
            }
            else if (shape.ShapeType is RoundedRectangleType)
            {
                //ROUNDED RECT
                this._writer.WriteStartElement("v", "roundrect", OpenXmlNamespaces.VectorML);
            }
            else if (shape.ShapeType is RectangleType)
            {
                //RECT
                this._writer.WriteStartElement("v", "rect", OpenXmlNamespaces.VectorML);
            }
            else if (shape.ShapeType is LineType)
            {
                //LINE
                this._writer.WriteStartElement("v", "line", OpenXmlNamespaces.VectorML);
            }
            else
            {
                //SHAPE
                if (shape.ShapeType != null)
                {
                    shape.ShapeType.Convert(new VMLShapeTypeMapping(this._writer));
                }
                this._writer.WriteStartElement("v", "shape", OpenXmlNamespaces.VectorML);
            }
        }

        /// <summary>
        /// Returns the OpenXML fill type of a fill effect
        /// </summary>
        private string getFillType(uint p)
        {
            switch (p)
            {
                case 0:
                    return "solid";
                case 1:
                    return "tile";
                case 2:
                    return "pattern";
                case 3:
                    return "frame";
                case 4:
                    return "gradient";
                case 5:
                    return "gradientRadial";
                case 6:
                    return "gradientRadial";
                case 7:
                    return "gradient";
                case 9:
                    return "solid";
                default:
                    return "solid";
            }
        }

        private string getShadowType(uint p)
        {
            switch (p)
            {
                case 0:
                    return "single";
                case 1:
                    return "double";
                case 2:
                    return "perspective";
                case 3:
                    return "shaperelative";
                case 4:
                    return "drawingrelative";
                case 5:
                    return "emboss";
                default:
                    return "single";
            }
        }

        private string getLineStyle(uint p)
        {
            switch (p)
            {
                case 0:
                    return "single";
                case 1:
                    return "thinThin";
                case 2:
                    return "thinThick";
                case 3:
                    return "thickThin";
                case 4:
                    return "thickBetweenThin";
                default:
                    return "single";
            }
        }

        private string getFillMethod(uint p)
        {
            short val = (short)((p & 0xFFFF0000) >> 28);
            switch (val)
            {
                case 0:
                    return "none";
                case 1:
                    return "any";
                case 2:
                    return "linear";
                case 4:
                    return "linear sigma";
                default:
                    return "any";
            }
        }

        /// <summary>
        /// Returns the OpenXML wrap type of the shape
        /// </summary>
        /// <param name="fspa"></param>
        /// <returns></returns>
        private string getWrapType(FileShapeAddress fspa)
        {
            // spec values
            // 0 = like 2 but doesn't equire absolute object
            // 1 = no text next to shape
            // 2 = wrap around absolute object
            // 3 = wrap as if no object present
            // 4 = wrap tightly areound object
            // 5 = wrap tightly but allow holes

            switch (fspa.wr)
            {
                case 0:
                case 2:
                    return "square";
                case 1:
                    return "topAndBottom";
                case 3:
                    return "through";
                case 4:
                case 5:
                    return "tight";
		        default:
                    return "none";
	        }
        }

        private string getArrowWidth(uint op)
        {
            switch (op)
            {
                default:
                    //msolineNarrowArrow
                    return "narrow";
                case 1:
                    //msolineMediumWidthArrow
                    return "medium";
                case 2:
                    //msolineWideArrow
                    return "wide";
            }
        }

        private string getArrowLength(uint op)
        {
            switch (op)
            {
                default:
                    //msolineShortArrow
                    return "short";
                case 1:
                    //msolineMediumLengthArrow
                    return "medium";
                case 2:
                    //msolineLongArrow
                    return "long";
            }
        }

        private string getArrowStyle(uint op)
        {
            switch (op)
            {
                default:
                    //msolineNoEnd
                    return "none";
                case 1:
                    //msolineArrowEnd
                    return "block";
                case 2:
                    //msolineArrowStealthEnd
                    return "classic";
                case 3:
                    //msolineArrowDiamondEnd
                    return "diamond";
                case 4:
                    //msolineArrowOvalEnd
                    return "oval";
                case 5:
                    //msolineArrowOpenEnd
                    return "open";
            }
        }

        /// <summary>
        /// Build the VML wrapcoords string for a given pWrapPolygonVertices
        /// </summary>
        /// <param name="pWrapPolygonVertices"></param>
        /// <returns></returns>
        private string getWrapCoords(ShapeOptions.OptionEntry pWrapPolygonVertices)
        {
            var r = new BinaryReader(new MemoryStream(pWrapPolygonVertices.opComplex));
            var pVertices = new List<int>();

            //skip first 6 bytes (header???)
            r.ReadBytes(6);

            //read the Int32 coordinates
            while (r.BaseStream.Position < r.BaseStream.Length)
            {
                pVertices.Add(r.ReadInt32());
            }

            //build the string
            var coords = new StringBuilder();
            foreach (int coord in pVertices)
            {
                coords.Append(coord);
                coords.Append(" ");
            }

            return coords.ToString().Trim();
        }

        /// <summary>
        /// Copies the picture from the binary stream to the zip archive 
        /// and creates the relationships for the image.
        /// </summary>
        /// <param name="pict">The PictureDescriptor</param>
        /// <returns>The created ImagePart</returns>
        protected ImagePart copyPicture(BlipStoreEntry bse)
        {
            //create the image part
            ImagePart imgPart = null;

            switch (bse.btWin32)
            {
                case BlipStoreEntry.BlipType.msoblipEMF:
                    imgPart = this._targetPart.AddImagePart(ImagePart.ImageType.Emf);
                    break;
                case BlipStoreEntry.BlipType.msoblipWMF:
                    imgPart = this._targetPart.AddImagePart(ImagePart.ImageType.Wmf);
                    break;
                case BlipStoreEntry.BlipType.msoblipJPEG:
                case BlipStoreEntry.BlipType.msoblipCMYKJPEG:
                    imgPart = this._targetPart.AddImagePart(ImagePart.ImageType.Jpeg);
                    break;
                case BlipStoreEntry.BlipType.msoblipPNG:
                    imgPart = this._targetPart.AddImagePart(ImagePart.ImageType.Png);
                    break;
                case BlipStoreEntry.BlipType.msoblipTIFF:
                    imgPart = this._targetPart.AddImagePart(ImagePart.ImageType.Tiff);
                    break;
                case BlipStoreEntry.BlipType.msoblipDIB:
                case BlipStoreEntry.BlipType.msoblipPICT:
                case BlipStoreEntry.BlipType.msoblipERROR:
                case BlipStoreEntry.BlipType.msoblipUNKNOWN:
                case BlipStoreEntry.BlipType.msoblipLastClient:
                case BlipStoreEntry.BlipType.msoblipFirstClient:
                    //throw new MappingException("Cannot convert picture of type " + bse.btWin32);
                    break;
            }

            if (imgPart != null)
            {
                var outStream = imgPart.GetStream();

                this._ctx.Doc.WordDocumentStream.Seek((long)bse.foDelay, SeekOrigin.Begin);
                var reader = new BinaryReader(this._ctx.Doc.WordDocumentStream);

                switch (bse.btWin32)
                {
                    case BlipStoreEntry.BlipType.msoblipEMF:
                    case BlipStoreEntry.BlipType.msoblipWMF:

                        //it's a meta image
                        var metaBlip = (MetafilePictBlip)Record.ReadRecord(reader);

                        //meta images can be compressed
                        var decompressed = metaBlip.Decrompress();
                        outStream.Write(decompressed, 0, decompressed.Length);

                        break;
                    case BlipStoreEntry.BlipType.msoblipJPEG:
                    case BlipStoreEntry.BlipType.msoblipCMYKJPEG:
                    case BlipStoreEntry.BlipType.msoblipPNG:
                    case BlipStoreEntry.BlipType.msoblipTIFF:

                        //it's a bitmap image
                        var bitBlip = (BitmapBlip)Record.ReadRecord(reader);
                        outStream.Write(bitBlip.m_pvBits, 0, bitBlip.m_pvBits.Length);
                        break;
                }
            }

            return imgPart;
        }

        //*******************************************************************
        //                                                     STATIC METHODS
        //*******************************************************************

        private static void AppendDimensionToStyle(StringBuilder style, PictureDescriptor pict, bool twistDimensions)
        {
            double xScaling = pict.mx / 1000.0;
            double yScaling = pict.my / 1000.0;
            var width = new TwipsValue(pict.dxaGoal * xScaling);
            var height = new TwipsValue(pict.dyaGoal * yScaling);

            if (twistDimensions)
            {
                width = new TwipsValue(pict.dyaGoal * yScaling);
                height = new TwipsValue(pict.dxaGoal * xScaling);
            }

            string widthString = Convert.ToString(width.ToPoints(), CultureInfo.GetCultureInfo("en-US"));
            string heightString = Convert.ToString(height.ToPoints(), CultureInfo.GetCultureInfo("en-US"));

            style.Append("width:").Append(widthString).Append("pt;");
            style.Append("height:").Append(heightString).Append("pt;");
        }

        public static void AppendDimensionToStyle(StringBuilder style, FileShapeAddress fspa, bool twistDimensions)
        {
            //append size and position ...
            appendStyleProperty(style, "position", "absolute");
            var left = new TwipsValue(fspa.xaLeft);
            var top = new TwipsValue(fspa.yaTop);
            var width = new TwipsValue(fspa.xaRight - fspa.xaLeft);
            var height = new TwipsValue(fspa.yaBottom - fspa.yaTop);

            if (twistDimensions)
            {
                width = new TwipsValue(fspa.yaBottom - fspa.yaTop);
                height = new TwipsValue(fspa.xaRight - fspa.xaLeft);
            }

            appendStyleProperty(style, "margin-left", Convert.ToString(left.ToPoints(), CultureInfo.GetCultureInfo("en-US")) + "pt");
            appendStyleProperty(style, "margin-top", Convert.ToString(top.ToPoints(), CultureInfo.GetCultureInfo("en-US")) + "pt");
            appendStyleProperty(style, "width", Convert.ToString(width.ToPoints(), CultureInfo.GetCultureInfo("en-US")) + "pt");
            appendStyleProperty(style, "height", Convert.ToString(height.ToPoints(), CultureInfo.GetCultureInfo("en-US")) + "pt");
        }

        public static void AppendDimensionToStyle(StringBuilder style, ChildAnchor anchor, bool twistDimensions)
        {
            //append size and position ...
            appendStyleProperty(style, "position", "absolute");
            appendStyleProperty(style, "left", anchor.rcgBounds.Left.ToString());
            appendStyleProperty(style, "top", anchor.rcgBounds.Top.ToString());
            if (twistDimensions)
            {
                appendStyleProperty(style, "width", anchor.rcgBounds.Height.ToString());
                appendStyleProperty(style, "height", anchor.rcgBounds.Width.ToString());
            }
            else
            {
                appendStyleProperty(style, "width", anchor.rcgBounds.Width.ToString());
                appendStyleProperty(style, "height", anchor.rcgBounds.Height.ToString());
            }
        }

        public static void AppendOptionsToStyle(StringBuilder style, List<ShapeOptions.OptionEntry> options)
        {
            foreach (var entry in options)
            {
                switch (entry.pid)
                {

                    //POSITIONING

                    case ShapeOptions.PropertyId.posh:
                        appendStyleProperty(
                            style,
                            "mso-position-horizontal",
                            mapHorizontalPosition((ShapeOptions.PositionHorizontal)entry.op));
                        break;
                    case ShapeOptions.PropertyId.posrelh:
                        appendStyleProperty(
                            style,
                            "mso-position-horizontal-relative",
                            mapHorizontalPositionRelative((ShapeOptions.PositionHorizontalRelative)entry.op));
                        break;
                    case ShapeOptions.PropertyId.posv:
                        appendStyleProperty(
                            style,
                            "mso-position-vertical",
                            mapVerticalPosition((ShapeOptions.PositionVertical)entry.op));
                        break;
                    case ShapeOptions.PropertyId.posrelv:
                        appendStyleProperty(
                            style,
                            "mso-position-vertical-relative",
                            mapVerticalPositionRelative((ShapeOptions.PositionVerticalRelative)entry.op));
                        break;

                    //BOOLEANS

                    case ShapeOptions.PropertyId.groupShapeBooleans:
                        var groupShapeBoolean = new GroupShapeBooleans(entry.op);

                        if (groupShapeBoolean.fUsefBehindDocument && groupShapeBoolean.fBehindDocument)
                        {
                            //The shape is behind the text, so the z-index must be negative.
                            appendStyleProperty(style, "z-index", "-1");
                        }

                        break;

                    // GEOMETRY

                    case ShapeOptions.PropertyId.rotation:
                        appendStyleProperty(style, "rotation", (entry.op / Math.Pow(2, 16)).ToString());
                        break;

                    //TEXTBOX

                    case ShapeOptions.PropertyId.anchorText:
                        appendStyleProperty(style, "v-text-anchor", getTextboxAnchor(entry.op));
                        break;

                    //WRAP DISTANCE

                    case ShapeOptions.PropertyId.dxWrapDistLeft:
                        appendStyleProperty(style, "mso-wrap-distance-left", new EmuValue((int)entry.op).ToPoints() + "pt");
                        break;

                    case ShapeOptions.PropertyId.dxWrapDistRight:
                        appendStyleProperty(style, "mso-wrap-distance-right", new EmuValue((int)entry.op).ToPoints() + "pt");
                        break;

                    case ShapeOptions.PropertyId.dyWrapDistBottom:
                        appendStyleProperty(style, "mso-wrap-distance-bottom", new EmuValue((int)entry.op).ToPoints() + "pt");
                        break;

                    case ShapeOptions.PropertyId.dyWrapDistTop:
                        appendStyleProperty(style, "mso-wrap-distance-top", new EmuValue((int)entry.op).ToPoints() + "pt");
                        break;

                }
            }
        }

        private static void appendStyleProperty(StringBuilder b, string propName, string propValue)
        {
            b.Append(propName);
            b.Append(":");
            b.Append(propValue);
            b.Append(";");
        }


        private static string getTextboxAnchor(uint anchor)
        {
            switch (anchor)
            {
                case 0:
                    //msoanchorTop
                    return "top";
                case 1:
                    //msoanchorMiddle
                    return "middle";
                case 2:
                    //msoanchorBottom
                    return "bottom";
                case 3:
                    //msoanchorTopCentered
                    return "top-center";
                case 4:
                    //msoanchorMiddleCentered
                    return "middle-center";
                case 5:
                    //msoanchorBottomCentered
                    return "bottom-center";
                case 6:
                    //msoanchorTopBaseline
                    return "top-baseline";
                case 7:
                    //msoanchorBottomBaseline
                    return "bottom-baseline";
                case 8:
                    //msoanchorTopCenteredBaseline
                    return "top-center-baseline";
                case 9:
                    //msoanchorBottomCenteredBaseline
                    return "bottom-center-baseline";
                default:
                    return "top";
            }
        }

        private static string mapVerticalPosition(ShapeOptions.PositionVertical vPos)
        {
            switch (vPos)
            {
                case ShapeOptions.PositionVertical.msopvAbs:
                    return "absolute";
                case ShapeOptions.PositionVertical.msopvTop:
                    return "top";
                case ShapeOptions.PositionVertical.msopvCenter:
                    return "center";
                case ShapeOptions.PositionVertical.msopvBottom:
                    return "bottom";
                case ShapeOptions.PositionVertical.msopvInside:
                    return "inside";
                case ShapeOptions.PositionVertical.msopvOutside:
                    return "outside";
                default:
                    return "absolute";
            }
        }

        private static string mapVerticalPositionRelative(ShapeOptions.PositionVerticalRelative vRel)
        {
            switch (vRel)
            {
                case ShapeOptions.PositionVerticalRelative.msoprvMargin:
                    return "margin";
                case ShapeOptions.PositionVerticalRelative.msoprvPage:
                    return "page";
                case ShapeOptions.PositionVerticalRelative.msoprvText:
                    return "text";
                case ShapeOptions.PositionVerticalRelative.msoprvLine:
                    return "line";
                default:
                    return "margin";
            }
        }

        private static string mapHorizontalPosition(ShapeOptions.PositionHorizontal hPos)
        {
            switch (hPos)
            {
                case ShapeOptions.PositionHorizontal.msophAbs:
                    return "absolute";
                case ShapeOptions.PositionHorizontal.msophLeft:
                    return "left";
                case ShapeOptions.PositionHorizontal.msophCenter:
                    return "center";
                case ShapeOptions.PositionHorizontal.msophRight:
                    return "right";
                case ShapeOptions.PositionHorizontal.msophInside:
                    return "inside";
                case ShapeOptions.PositionHorizontal.msophOutside:
                    return "outside";
                default:
                    return "absolute";
            }
        }

        private static string mapHorizontalPositionRelative(ShapeOptions.PositionHorizontalRelative hRel)
        {
            switch (hRel) 
            {
                case ShapeOptions.PositionHorizontalRelative.msoprhMargin:
                    return "margin";
                case ShapeOptions.PositionHorizontalRelative.msoprhPage:
                    return "page";
                case ShapeOptions.PositionHorizontalRelative.msoprhText:
                    return "text";
                case ShapeOptions.PositionHorizontalRelative.msoprhChar:
                    return "char";
                default:
                    return "margin";
            }
        }

        /// <summary>
        /// Generates a string id for the given shape
        /// </summary>
        /// <param name="shape"></param>
        /// <returns></returns>
        private static string getShapeId(Shape shape)
        {
            var id = new StringBuilder();
            id.Append("_x0000_s");
            id.Append(shape.spid);
            return id.ToString();
        }

    }
}
