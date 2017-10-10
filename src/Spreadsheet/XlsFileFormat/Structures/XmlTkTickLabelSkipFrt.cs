

using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkTickLabelSkipFrt
    {
        public XmlTkDWord nInternal;

        public XmlTkTickLabelSkipFrt(IStreamReader reader)
        {
            this.nInternal = new XmlTkDWord(reader);   
        }
    }
}
