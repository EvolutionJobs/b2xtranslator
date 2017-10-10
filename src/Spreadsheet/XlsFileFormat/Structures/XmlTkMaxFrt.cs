

using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkMaxFrt
    {
        public XmlTkDouble maxScale;

        public XmlTkMaxFrt(IStreamReader reader)
        {
            this.maxScale = new XmlTkDouble(reader);   
        }
    }
}
