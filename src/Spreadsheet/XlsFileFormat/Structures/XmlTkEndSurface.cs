

using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkEndSurface
    {
        public XmlTkEnd endSurface;

        public XmlTkEndSurface(IStreamReader reader)
        {
            this.endSurface = new XmlTkEnd(reader);   
        }
    }
}
