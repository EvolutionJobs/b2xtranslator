

using b2xtranslator.PptFileFormat;
using b2xtranslator.OpenXmlLib;
using b2xtranslator.OpenXmlLib.PresentationML;
using b2xtranslator.Tools;
using b2xtranslator.OfficeDrawing;

namespace b2xtranslator.PresentationMLMapping
{
    public class TitleMasterMapping : PresentationMapping<RegularContainer>
    {
        public TitleMasterMapping(ConversionContext ctx, SlideLayoutPart part)
            : base(ctx, part)
        {

        }

        override public void Apply(RegularContainer slide)
        {
            TraceLogger.DebugInternal("TitleMasterMapping.Apply");

            // Start the document
            this._writer.WriteStartDocument();
            this._writer.WriteStartElement("p", "sldLayout", OpenXmlNamespaces.PresentationML);

            this._writer.WriteAttributeString("showMasterSp", "0");
            this._writer.WriteAttributeString("type", "title");
            this._writer.WriteAttributeString("preserve", "1");

            // Force declaration of these namespaces at document start
            this._writer.WriteAttributeString("xmlns", "a", null, OpenXmlNamespaces.DrawingML);
            // Force declaration of these namespaces at document start
            this._writer.WriteAttributeString("xmlns", "r", null, OpenXmlNamespaces.Relationships);

            this._writer.WriteStartElement("p", "cSld", OpenXmlNamespaces.PresentationML);



            this._writer.WriteStartElement("p", "spTree", OpenXmlNamespaces.PresentationML);
            var stm = new ShapeTreeMapping(this._ctx, this._writer);
            stm.parentSlideMapping = this;
            stm.Apply(slide.FirstChildWithType<PPDrawing>());
            this._writer.WriteEndElement();
            this._writer.WriteEndElement();

            // End the document
            this._writer.WriteEndElement();
            this._writer.WriteEndDocument();

            this._writer.Flush();
        }
    }
}
