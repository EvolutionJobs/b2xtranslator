

using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkRAngAxOffFrt
    {
        public XmlTkBool fRightAngAxOff;

        public XmlTkRAngAxOffFrt(IStreamReader reader)
        {
            this.fRightAngAxOff = new XmlTkBool(reader);   
        }
    }
}
