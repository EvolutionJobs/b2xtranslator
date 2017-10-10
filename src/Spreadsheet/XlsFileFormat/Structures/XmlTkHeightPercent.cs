

using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkHeightPercent
    {
        public XmlTkDouble heightPercent;

        public XmlTkHeightPercent(IStreamReader reader)
        {
            this.heightPercent = new XmlTkDouble(reader);   
        }
    }
}
