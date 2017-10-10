

using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkTickLabelPositionFrt
    {
        public XmlTkToken xmltkHigh;

        public XmlTkTickLabelPositionFrt(IStreamReader reader)
        {
            this.xmltkHigh = new XmlTkToken(reader);   
        }
    }
}
