

using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    public class XmlTkColorMappingOverride
    {
        public XmlTkBlob rgThemeOverride;

        public XmlTkColorMappingOverride(IStreamReader reader)
        {
            this.rgThemeOverride = new XmlTkBlob(reader);   
        }
    }
}
