

using b2xtranslator.OpenXmlLib;
using System.Xml;
using b2xtranslator.CommonTranslatorLib;

namespace b2xtranslator.PresentationMLMapping
{
    public class AppMapping : AbstractOpenXmlMapping,
          IMapping<IVisitable>
    {

        public AppMapping(AppPropertiesPart appPart, XmlWriterSettings xws)
            : base(XmlWriter.Create(appPart.GetStream(), xws))
        {
        }

        public void Apply(IVisitable x)
        {
            // Start the document
            this._writer.WriteStartDocument();
            this._writer.WriteStartElement("Properties", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties");

            // Force declaration of these namespaces at document start
            //_writer.WriteAttributeString("xmlns", null, null, OpenXmlRelationshipTypes.ExtendedProperties);
            this._writer.WriteAttributeString("xmlns", "vt", null, OpenXmlNamespaces.docPropsVTypes);

            //TotalTime
            this._writer.WriteElementString("TotalTime", "0");
            //Words
            this._writer.WriteElementString("Words", "0");
            //PresentationFormat
            this._writer.WriteElementString("PresentationFormat", "Custom");
            //Paragraphs
            this._writer.WriteElementString("Paragraphs", "0");
            //Slides
            this._writer.WriteElementString("Slides", "0");
            //Notes
            this._writer.WriteElementString("Notes", "0");
            //HiddenSlides
            this._writer.WriteElementString("HiddenSlides", "0");
            //MMClips
            this._writer.WriteElementString("MMClips", "0");
            //ScaleCrop
            this._writer.WriteElementString("ScaleCrop", "false");
            //HeadingPairs
            //TitlesOfParts           

            //LinksUpToDate
            this._writer.WriteElementString("LinksUpToDate", "false");
            //SharedDoc
            this._writer.WriteElementString("SharedDoc", "false");
            //HyperlinksChanged
            this._writer.WriteElementString("HyperlinksChanged", "false");
            //AppVersion
            this._writer.WriteElementString("AppVersion", "12.0000");

            // End the document
            this._writer.WriteEndElement(); //Properties
            this._writer.WriteEndDocument();

            this._writer.Flush();
        }        
    }
}
