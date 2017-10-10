using System;
using System.Text;
using b2xtranslator.DocFileFormat;
using b2xtranslator.CommonTranslatorLib;
using System.Xml;
using b2xtranslator.OpenXmlLib;

namespace b2xtranslator.WordprocessingMLMapping
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
            this._parent = parent;
            this._nodeFactory = parent.OwnerDocument;
        }

        public void Apply(DateAndTime dttm)
        {
            var date = new StringBuilder();
            date.Append(string.Format("{0:0000}", dttm.yr));
            date.Append("-");
            date.Append(string.Format("{0:00}", dttm.mon));
            date.Append("-");
            date.Append(string.Format("{0:00}", dttm.dom));
            date.Append("T");
            date.Append(string.Format("{0:00}", dttm.hr));
            date.Append(":");
            date.Append(string.Format("{0:00}", dttm.mint));
            date.Append(":00Z");

            var xml = this._nodeFactory.CreateAttribute("w", "date", OpenXmlNamespaces.WordprocessingML);
            xml.Value = date.ToString() ;

            //append or write
            if (this._writer != null)
            {
                xml.WriteTo(this._writer);
            }
            else if (this._parent != null)
            {
                this._parent.Attributes.Append(xml);
            }
        }
    }
}
