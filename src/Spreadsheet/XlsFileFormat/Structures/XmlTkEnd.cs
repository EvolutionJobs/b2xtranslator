

using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkEnd
    {
        public XmlTkHeader xtHeader;

        public XmlTkEnd(IStreamReader reader)
        {
            this.xtHeader = new XmlTkHeader(reader);   
        }
    }
}
