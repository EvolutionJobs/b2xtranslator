using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.DocFileFormat;
using System.Xml;
using b2xtranslator.OpenXmlLib;

namespace b2xtranslator.WordprocessingMLMapping
{
    public class ApplicationPropertiesMapping : AbstractOpenXmlMapping,
          IMapping<DocumentProperties>
    {

        public ApplicationPropertiesMapping(AppPropertiesPart appPart, XmlWriterSettings xws)
            : base(XmlWriter.Create(appPart.GetStream(), xws))
        {
        }

        public void Apply(DocumentProperties dop)
        {
            //start Properties
            this._writer.WriteStartElement("w", "Properties", OpenXmlNamespaces.WordprocessingML);

            //Application
            //AppVersion
            //Company
            //DigSig
            //DocSecurity
            //HeadingPairs
            //HiddenSlides
            //HLinks
            //HyperlinkBase
            //HyperlinksChanged
            //LinksUpToDate
            //Manager
            //MMClips
            //Notes
            //PresentationFormat
            //ScaleCrop
            //SharedDoc
            //Slides
            //Template
            //TitlesOfParts
            //TotalTime

            //WordCount statistics

            //CharactersWithSpaces
            this._writer.WriteStartElement("CharactersWithSpaces");
            this._writer.WriteString(dop.cChWS.ToString());
            this._writer.WriteEndElement();

            //Characters
            this._writer.WriteStartElement("Characters");
            this._writer.WriteString(dop.cCh.ToString());
            this._writer.WriteEndElement();

            //Lines
            this._writer.WriteStartElement("Lines");
            this._writer.WriteString(dop.cLines.ToString());
            this._writer.WriteEndElement();

            //Pages
            this._writer.WriteStartElement("Pages");
            this._writer.WriteString(dop.cPg.ToString());
            this._writer.WriteEndElement();

            //Paragraphs
            this._writer.WriteStartElement("Paragraphs");
            this._writer.WriteString(dop.cParas.ToString());
            this._writer.WriteEndElement();

            //Words
            this._writer.WriteStartElement("Words");
            this._writer.WriteString(dop.cWords.ToString());
            this._writer.WriteEndElement();

            //end Properties
            this._writer.WriteEndElement();
        }
    }
}
