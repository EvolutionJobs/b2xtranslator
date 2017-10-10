

using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkSpb
    {
        public XmlTkBlob shapePropsStream;

        public XmlTkSpb(IStreamReader reader)
        {
            this.shapePropsStream = new XmlTkBlob(reader);   
        }
    }
}
