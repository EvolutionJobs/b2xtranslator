

using b2xtranslator.OpenXmlLib;
using System.Xml;
using b2xtranslator.CommonTranslatorLib;

namespace b2xtranslator.PresentationMLMapping
{
    public class CoreMapping : AbstractOpenXmlMapping,
          IMapping<IVisitable>
    {

        public CoreMapping(CorePropertiesPart corePart, XmlWriterSettings xws)
            : base(XmlWriter.Create(corePart.GetStream(), xws))
        {
        }

        public void Apply(IVisitable x)
        {
            // Start the document
            this._writer.WriteStartDocument();
            this._writer.WriteStartElement("cp", "coreProperties", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties");

            // Force declaration of these namespaces at document start
            this._writer.WriteAttributeString("xmlns", "cp", null, "http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
            this._writer.WriteAttributeString("xmlns", "dc", null, "http://purl.org/dc/elements/1.1/");
            this._writer.WriteAttributeString("xmlns", "dcterms", null, "http://purl.org/dc/terms/");
            this._writer.WriteAttributeString("xmlns", "dcmitype", null, "http://purl.org/dc/dcmitype/");
            this._writer.WriteAttributeString("xmlns", "xsi", null, "http://www.w3.org/2001/XMLSchema-instance");

            this._writer.WriteElementString("dc", "title", "http://purl.org/dc/elements/1.1/", "Title");
            this._writer.WriteElementString("dc", "creator", "http://purl.org/dc/elements/1.1/", "Creator");
            this._writer.WriteElementString("dc", "description", "http://purl.org/dc/elements/1.1/", "Description");
            this._writer.WriteElementString("cp", "lastModifiedBy", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties", "Modifier");
            this._writer.WriteElementString("cp", "revision", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties", "10");
            this._writer.WriteElementString("cp", "lastPrinted", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties", "1601-01-01T00:00:00Z");

            this._writer.WriteStartElement("dcterms", "created", "http://purl.org/dc/terms/");
            this._writer.WriteAttributeString("xsi", "type", "http://www.w3.org/2001/XMLSchema-instance", "dcterms:W3CDTF");
            this._writer.WriteString("2008-05-13T17:08:28Z");
            this._writer.WriteEndElement();

            this._writer.WriteStartElement("dcterms", "modified", "http://purl.org/dc/terms/");
            this._writer.WriteAttributeString("xsi", "type", "http://www.w3.org/2001/XMLSchema-instance", "dcterms:W3CDTF");
            this._writer.WriteString("2009-03-30T11:17:58Z");
            this._writer.WriteEndElement();


            // End the document
            this._writer.WriteEndElement(); //coreProperties
            this._writer.WriteEndDocument();

            this._writer.Flush();
        }        
    }
}
