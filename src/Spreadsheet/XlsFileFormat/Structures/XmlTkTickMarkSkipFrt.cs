

using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkTickMarkSkipFrt
    {
        public XmlTkDWord nInternal;

        public XmlTkTickMarkSkipFrt(IStreamReader reader)
        {
            this.nInternal = new XmlTkDWord(reader);   
        }
    }
}
