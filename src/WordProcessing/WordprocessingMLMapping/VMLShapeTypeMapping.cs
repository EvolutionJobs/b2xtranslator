using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.OfficeDrawing.Shapetypes;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class VMLShapeTypeMapping : PropertiesMapping,
          IMapping<ShapeType>
    {
        private XmlElement _lock = null;

        public VMLShapeTypeMapping(XmlWriter writer)
            : base(writer)
        {
            _lock = _nodeFactory.CreateElement("o", "lock", OpenXmlNamespaces.Office);
            appendValueAttribute(_lock, "v", "ext", "edit", OpenXmlNamespaces.VectorML);
        }

        public void Apply(ShapeType shapeType)
        {
            if (!(shapeType is DIaLOGIKa.b2xtranslator.OfficeDrawing.Shapetypes.OvalType))
            {
                _writer.WriteStartElement("v", "shapetype", OpenXmlNamespaces.VectorML);

                //id
                _writer.WriteAttributeString("id", GenerateTypeId(shapeType));

                //coordinate system
                _writer.WriteAttributeString("coordsize", "21600,21600");

                //shape type code
                _writer.WriteAttributeString("o", "spt", OpenXmlNamespaces.Office, shapeType.TypeCode.ToString());

                //adj
                if (shapeType.AdjustmentValues != null)
                {
                    _writer.WriteAttributeString("adj", shapeType.AdjustmentValues);
                }

                //The path
                if (shapeType.Path != null)
                {
                    _writer.WriteAttributeString("path", shapeType.Path);
                }

                // always write this attribute 
                // if this causes regression bugs, remove it.
                // this was inserted due to a bug in Word 2007 (sf.net item: 2256373)
                if(shapeType.PreferRelative)
                {
                    _writer.WriteAttributeString("o", "preferrelative",OpenXmlNamespaces.Office, "t");
                }

                //Default fill / stroke
                if (shapeType.Filled == false)
                {
                    _writer.WriteAttributeString("filled", "f");
                }
                if (shapeType.Stroked == false)
                {
                    _writer.WriteAttributeString("stroked", "f");
                }

                //stroke
                if (shapeType.Joins != ShapeType.JoinStyle.none)
                {
                    _writer.WriteStartElement("v", "stroke", OpenXmlNamespaces.VectorML);
                    _writer.WriteAttributeString("joinstyle", shapeType.Joins.ToString());
                    _writer.WriteEndElement();
                }

                //Formulas
                if (shapeType.Formulas != null && shapeType.Formulas.Count > 0)
                {
                    _writer.WriteStartElement("v", "formulas", OpenXmlNamespaces.VectorML);

                    foreach (string formula in shapeType.Formulas)
                    {
                        _writer.WriteStartElement("v", "f", OpenXmlNamespaces.VectorML);
                        _writer.WriteAttributeString("eqn", formula);
                        _writer.WriteEndElement();
                    }

                    _writer.WriteEndElement();
                }


                //Path
                _writer.WriteStartElement("v", "path", OpenXmlNamespaces.VectorML);
                if (shapeType.ShapeConcentricFill)
                {
                    _writer.WriteAttributeString("gradientshapeok", "t");
                }
                if (shapeType.Limo != null)
                {
                    _writer.WriteAttributeString("limo", shapeType.Limo);
                }
                if (shapeType.TextPath)
                {
                    _writer.WriteAttributeString("textpathok", "t");
                    
                }
                if (shapeType.ConnectorLocations != null)
                {
                    _writer.WriteAttributeString("o", "connecttype", OpenXmlNamespaces.Office, "custom");
                    _writer.WriteAttributeString("o", "connectlocs", OpenXmlNamespaces.Office, shapeType.ConnectorLocations);
                }
                else if(shapeType.ConnectorType != null)
                {
                    _writer.WriteAttributeString("o", "connecttype", OpenXmlNamespaces.Office, shapeType.ConnectorType);
                }
                if (shapeType.TextboxRectangle != null)
                {
                    _writer.WriteAttributeString("textboxrect", shapeType.TextboxRectangle);
                }
                if (shapeType.ConnectorAngles != null)
                {
                    _writer.WriteAttributeString("o", "connectangles", OpenXmlNamespaces.Office, shapeType.ConnectorAngles);
                }

                // always write this attribute 
                // if this causes regression bugs, remove it.
                // this was inserted due to a bug in Word 2007 (sf.net item: 2256373)
                if (shapeType.ExtrusionOk == false)
                {
                    _writer.WriteAttributeString("o", "extrusionok", OpenXmlNamespaces.Office, "f");
                }

                _writer.WriteEndElement();

                //Lock
                if (shapeType.Lock.fUsefLockAspectRatio && shapeType.Lock.fLockAspectRatio)
                {
                    appendValueAttribute(_lock, null, "aspectratio", "t", null);
                }
                if (shapeType.Lock.fUsefLockText && shapeType.Lock.fLockText)
                {
                    appendValueAttribute(_lock, "v", "ext", "edit", OpenXmlNamespaces.VectorML);
                    appendValueAttribute(_lock, null, "text", "t", null);
                }
                if (shapeType.LockShapeType)
                {
                    appendValueAttribute(_lock, null, "shapetype", "t", null);
                }
                if (_lock.Attributes.Count > 1)
                {
                    _lock.WriteTo(_writer);
                }

                // Textpath
                if (shapeType.TextPath)
                {
                    _writer.WriteStartElement("v", "textpath", OpenXmlNamespaces.VectorML);
                    _writer.WriteAttributeString("on", "t");
                    StringBuilder textPathStyle = new StringBuilder();
                    if (shapeType.TextKerning)
                    {
                        textPathStyle.Append("v-text-kern:t;");
                    }
                    if (textPathStyle.Capacity > 0)
                    {
                        _writer.WriteAttributeString("style", textPathStyle.ToString());
                    }
                    _writer.WriteAttributeString("fitpath", "t");
                    _writer.WriteEndElement();
                }

                //Handles
                if (shapeType.Handles != null && shapeType.Handles.Count > 0)
                {
                    _writer.WriteStartElement("v", "handles", OpenXmlNamespaces.VectorML);

                    foreach (ShapeType.Handle handle in shapeType.Handles)
                    {
                        _writer.WriteStartElement("v", "h", OpenXmlNamespaces.VectorML);

                        if (handle.position != null)
                            _writer.WriteAttributeString("position", handle.position);

                        if (handle.switchHandle != null)
                            _writer.WriteAttributeString("switch", handle.switchHandle);

                        if (handle.xrange != null)
                            _writer.WriteAttributeString("xrange", handle.xrange);

                        if (handle.yrange != null)
                            _writer.WriteAttributeString("yrange", handle.yrange);

                        if (handle.polar != null)
                            _writer.WriteAttributeString("polar", handle.polar);

                        if (handle.radiusrange != null)
                            _writer.WriteAttributeString("radiusrange", handle.radiusrange);

                        _writer.WriteEndElement();
                    }

                    _writer.WriteEndElement();
                }



                _writer.WriteEndElement();

            }
            _writer.Flush();
        }


        /// <summary>
        /// Returns the id of the referenced type
        /// </summary>
        public static string GenerateTypeId(ShapeType shapeType)
        {
            StringBuilder type = new StringBuilder();
            type.Append("_x0000_t");
            type.Append(shapeType.TypeCode);
            return type.ToString();
        }

        
    }
}
