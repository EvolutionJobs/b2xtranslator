using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.DocFileFormat;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;

namespace DIaLOGIKa.b2xtranslator.WordprocessingMLMapping
{
    public class DateMapping : AbstractOpenXmlMapping,
          IMapping<DateAndTime>
    {
        XmlElement _parent;

        /// <summary>
        /// Writes a date attribute to the given writer
        /// </summary>
        /// <param name="writer"></param>
        public DateMapping(XmlWriter writer)
            : base(writer)
        {
        }

        /// <summary>
        /// Appends a date attribute to the given Element
        /// </summary>
        /// <param name="parent"></param>
        public DateMapping(XmlElement parent)
            : base(null)
        {
            _parent = parent;
            _nodeFactory = parent.OwnerDocument;
        }

        public void Apply(DateAndTime dttm)
        {
            StringBuilder date = new StringBuilder();
            date.Append(String.Format("{0:0000}", dttm.yr));
            date.Append("-");
            date.Append(String.Format("{0:00}", dttm.mon));
            date.Append("-");
            date.Append(String.Format("{0:00}", dttm.dom));
            date.Append("T");
            date.Append(String.Format("{0:00}", dttm.hr));
            date.Append(":");
            date.Append(String.Format("{0:00}", dttm.mint));
            date.Append(":00Z");

            XmlAttribute xml = _nodeFactory.CreateAttribute("w", "date", OpenXmlNamespaces.WordprocessingML);
            xml.Value = date.ToString() ;

            //append or write
            if (_writer != null)
            {
                xml.WriteTo(_writer);
            }
            else if (_parent != null)
            {
                _parent.Attributes.Append(xml);
            }
        }
    }
}
