

using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkThemeOverride
    {
        public XmlTkBlob rgThemeOverride;

        public XmlTkThemeOverride(IStreamReader reader)
        {
            this.rgThemeOverride = new XmlTkBlob(reader);   
        }
    }
}
