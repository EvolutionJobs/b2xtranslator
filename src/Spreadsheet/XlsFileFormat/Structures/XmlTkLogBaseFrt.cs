

using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkLogBaseFrt
    {
        public XmlTkDouble logScale;

        public XmlTkLogBaseFrt(IStreamReader reader)
        {
            this.logScale = new XmlTkDouble(reader);   
        }
    }
}
