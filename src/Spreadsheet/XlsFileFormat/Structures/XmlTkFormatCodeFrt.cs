

using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkFormatCodeFrt
    {
        public XmlTkString stFormat;

        public XmlTkFormatCodeFrt(IStreamReader reader)
        {
            this.stFormat = new XmlTkString(reader);   
        }
    }
}
