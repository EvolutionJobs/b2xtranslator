

using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkOverlay
    {
        public XmlTkBool fOverlay;

        public XmlTkOverlay(IStreamReader reader)
        {
            this.fOverlay = new XmlTkBool(reader);    
        }
    }
}
