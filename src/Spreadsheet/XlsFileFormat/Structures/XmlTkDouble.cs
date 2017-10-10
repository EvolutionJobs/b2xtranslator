

using System;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkDouble
    {
        public XmlTkHeader xtHeader;

        public Double dValue;

        public XmlTkDouble(IStreamReader reader)
        {
            this.xtHeader = new XmlTkHeader(reader);

            //unused
            reader.ReadBytes(4);

            this.dValue = reader.ReadDouble();       
        }
    }
}
