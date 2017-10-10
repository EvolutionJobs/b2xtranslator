using System.Collections.Generic;
using b2xtranslator.OfficeDrawing;
using System.IO;

namespace b2xtranslator.PptFileFormat
{
    [OfficeRecord(1016)]
    public class MainMaster : Slide
    {
        public Dictionary<string, string> Layouts = new Dictionary<string, string>();
        public MainMaster(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) {
                foreach (var rec in this.Children)
                {
                    if (rec is RoundTripContentMasterInfo12)
                    {
                        var info = (RoundTripContentMasterInfo12)rec;
                        string xml = info.XmlDocumentElement.OuterXml;
                        xml = xml.Replace("http://schemas.openxmlformats.org/drawingml/2006/3/main", "http://schemas.openxmlformats.org/drawingml/2006/main");
                        if (info.XmlDocumentElement.Attributes["type"] != null)
                        {
                            string title = info.XmlDocumentElement.Attributes["type"].InnerText;
                        this.Layouts.Add(title, xml);
                        }
                    }           
                }
        }
    }
}
