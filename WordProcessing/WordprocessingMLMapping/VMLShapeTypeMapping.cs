using System.Text;
using System.Xml;
using b2xtranslator.OpenXmlLib;
using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.OfficeDrawing.Shapetypes;

namespace b2xtranslator.WordprocessingMLMapping
{
    public class VMLShapeTypeMapping : PropertiesMapping,
          IMapping<ShapeType>
    {
        private XmlElement _lock = null;

        public VMLShapeTypeMapping(XmlWriter writer)
            : base(writer)
        {
            this._lock = this._nodeFactory.CreateElement("o", "lock", OpenXmlNamespaces.Office);
            appendValueAttribute(this._lock, "v", "ext", "edit", OpenXmlNamespaces.VectorML);
        }

        public void Apply(ShapeType shapeType)
        {
            if (!(shapeType is b2xtranslator.OfficeDrawing.Shapetypes.OvalType))
            {
                this._writer.WriteStartElement("v", "shapetype", OpenXmlNamespaces.VectorML);

                //id
                this._writer.WriteAttributeString("id", GenerateTypeId(shapeType));

                //coordinate system
                this._writer.WriteAttributeString("coordsize", "21600,21600");

                //shape type code
                this._writer.WriteAttributeString("o", "spt", OpenXmlNamespaces.Office, shapeType.TypeCode.ToString());

                //adj
                if (shapeType.AdjustmentValues != null)
                {
                    this._writer.WriteAttributeString("adj", shapeType.AdjustmentValues);
                }

                //The path
                if (shapeType.Path != null)
                {
                    this._writer.WriteAttributeString("path", shapeType.Path);
                }

                // always write this attribute 
                // if this causes regression bugs, remove it.
                // this was inserted due to a bug in Word 2007 (sf.net item: 2256373)
                if(shapeType.PreferRelative)
                {
                    this._writer.WriteAttributeString("o", "preferrelative",OpenXmlNamespaces.Office, "t");
                }

                //Default fill / stroke
                if (shapeType.Filled == false)
                {
                    this._writer.WriteAttributeString("filled", "f");
                }
                if (shapeType.Stroked == false)
                {
                    this._writer.WriteAttributeString("stroked", "f");
                }

                //stroke
                if (shapeType.Joins != ShapeType.JoinStyle.none)
                {
                    this._writer.WriteStartElement("v", "stroke", OpenXmlNamespaces.VectorML);
                    this._writer.WriteAttributeString("joinstyle", shapeType.Joins.ToString());
                    this._writer.WriteEndElement();
                }

                //Formulas
                if (shapeType.Formulas != null && shapeType.Formulas.Count > 0)
                {
                    this._writer.WriteStartElement("v", "formulas", OpenXmlNamespaces.VectorML);

                    foreach (string formula in shapeType.Formulas)
                    {
                        this._writer.WriteStartElement("v", "f", OpenXmlNamespaces.VectorML);
                        this._writer.WriteAttributeString("eqn", formula);
                        this._writer.WriteEndElement();
                    }

                    this._writer.WriteEndElement();
                }


                //Path
                this._writer.WriteStartElement("v", "path", OpenXmlNamespaces.VectorML);
                if (shapeType.ShapeConcentricFill)
                {
                    this._writer.WriteAttributeString("gradientshapeok", "t");
                }
                if (shapeType.Limo != null)
                {
                    this._writer.WriteAttributeString("limo", shapeType.Limo);
                }
                if (shapeType.TextPath)
                {
                    this._writer.WriteAttributeString("textpathok", "t");
                    
                }
                if (shapeType.ConnectorLocations != null)
                {
                    this._writer.WriteAttributeString("o", "connecttype", OpenXmlNamespaces.Office, "custom");
                    this._writer.WriteAttributeString("o", "connectlocs", OpenXmlNamespaces.Office, shapeType.ConnectorLocations);
                }
                else if(shapeType.ConnectorType != null)
                {
                    this._writer.WriteAttributeString("o", "connecttype", OpenXmlNamespaces.Office, shapeType.ConnectorType);
                }
                if (shapeType.TextboxRectangle != null)
                {
                    this._writer.WriteAttributeString("textboxrect", shapeType.TextboxRectangle);
                }
                if (shapeType.ConnectorAngles != null)
                {
                    this._writer.WriteAttributeString("o", "connectangles", OpenXmlNamespaces.Office, shapeType.ConnectorAngles);
                }

                // always write this attribute 
                // if this causes regression bugs, remove it.
                // this was inserted due to a bug in Word 2007 (sf.net item: 2256373)
                if (shapeType.ExtrusionOk == false)
                {
                    this._writer.WriteAttributeString("o", "extrusionok", OpenXmlNamespaces.Office, "f");
                }

                this._writer.WriteEndElement();

                //Lock
                if (shapeType.Lock.fUsefLockAspectRatio && shapeType.Lock.fLockAspectRatio)
                {
                    appendValueAttribute(this._lock, null, "aspectratio", "t", null);
                }
                if (shapeType.Lock.fUsefLockText && shapeType.Lock.fLockText)
                {
                    appendValueAttribute(this._lock, "v", "ext", "edit", OpenXmlNamespaces.VectorML);
                    appendValueAttribute(this._lock, null, "text", "t", null);
                }
                if (shapeType.LockShapeType)
                {
                    appendValueAttribute(this._lock, null, "shapetype", "t", null);
                }
                if (this._lock.Attributes.Count > 1)
                {
                    this._lock.WriteTo(this._writer);
                }

                // Textpath
                if (shapeType.TextPath)
                {
                    this._writer.WriteStartElement("v", "textpath", OpenXmlNamespaces.VectorML);
                    this._writer.WriteAttributeString("on", "t");
                    var textPathStyle = new StringBuilder();
                    if (shapeType.TextKerning)
                    {
                        textPathStyle.Append("v-text-kern:t;");
                    }
                    if (textPathStyle.Capacity > 0)
                    {
                        this._writer.WriteAttributeString("style", textPathStyle.ToString());
                    }
                    this._writer.WriteAttributeString("fitpath", "t");
                    this._writer.WriteEndElement();
                }

                //Handles
                if (shapeType.Handles != null && shapeType.Handles.Count > 0)
                {
                    this._writer.WriteStartElement("v", "handles", OpenXmlNamespaces.VectorML);

                    foreach (var handle in shapeType.Handles)
                    {
                        this._writer.WriteStartElement("v", "h", OpenXmlNamespaces.VectorML);

                        if (handle.position != null)
                            this._writer.WriteAttributeString("position", handle.position);

                        if (handle.switchHandle != null)
                            this._writer.WriteAttributeString("switch", handle.switchHandle);

                        if (handle.xrange != null)
                            this._writer.WriteAttributeString("xrange", handle.xrange);

                        if (handle.yrange != null)
                            this._writer.WriteAttributeString("yrange", handle.yrange);

                        if (handle.polar != null)
                            this._writer.WriteAttributeString("polar", handle.polar);

                        if (handle.radiusrange != null)
                            this._writer.WriteAttributeString("radiusrange", handle.radiusrange);

                        this._writer.WriteEndElement();
                    }

                    this._writer.WriteEndElement();
                }



                this._writer.WriteEndElement();

            }
            this._writer.Flush();
        }


        /// <summary>
        /// Returns the id of the referenced type
        /// </summary>
        public static string GenerateTypeId(ShapeType shapeType)
        {
            var type = new StringBuilder();
            type.Append("_x0000_t");
            type.Append(shapeType.TypeCode);
            return type.ToString();
        }

        
    }
}
