

using b2xtranslator.PptFileFormat;
using b2xtranslator.OpenXmlLib;
using System.Xml;
using b2xtranslator.CommonTranslatorLib;

namespace b2xtranslator.PresentationMLMapping
{
    public class viewPropsMapping : AbstractOpenXmlMapping,
          IMapping<IVisitable>
    {
        private ConversionContext _ctx;

        public viewPropsMapping(ViewPropertiesPart viewPart, XmlWriterSettings xws, ConversionContext ctx)
            : base(XmlWriter.Create(viewPart.GetStream(), xws))
        {
            this._ctx = ctx;
        }

        public void Apply(IVisitable x)
        {
            // Start the document
            this._writer.WriteStartDocument();
            this._writer.WriteStartElement("p", "viewPr", OpenXmlNamespaces.PresentationML);

            // Force declaration of these namespaces at document start
            this._writer.WriteAttributeString("xmlns", "a", null, OpenXmlNamespaces.DrawingML);
            this._writer.WriteAttributeString("xmlns", "r", null, OpenXmlNamespaces.Relationships);


            //_writer.WriteStartElement("p", "normalViewPr", OpenXmlNamespaces.PresentationML);
            //_writer.WriteStartElement("p", "restoredLeft", OpenXmlNamespaces.PresentationML);
            //_writer.WriteAttributeString("sz", "10031");
            //_writer.WriteAttributeString("autoAdjust", "0");
            //_writer.WriteEndElement(); //restoredLeft
            //_writer.WriteStartElement("p", "restoredTop", OpenXmlNamespaces.PresentationML);
            //_writer.WriteAttributeString("sz", "94660");
            //_writer.WriteEndElement(); //restoredLeft
            //_writer.WriteEndElement(); //normalViewPr


            //NormalViewSetInfoAtom a = _ctx.Ppt.DocumentRecord.DocInfoListContainer.FirstChildWithType<NormalViewSetInfoContainer>().FirstChildWithType<NormalViewSetInfoAtom>();
            var z = this._ctx.Ppt.DocumentRecord.DocInfoListContainer.FirstChildWithType<SlideViewInfoContainer>().FirstChildWithType<ZoomViewInfoAtom>();

            this._writer.WriteStartElement("p", "slideViewPr", OpenXmlNamespaces.PresentationML);
            this._writer.WriteStartElement("p", "cSldViewPr", OpenXmlNamespaces.PresentationML);
            this._writer.WriteStartElement("p", "cViewPr", OpenXmlNamespaces.PresentationML);
            this._writer.WriteStartElement("p", "scale", OpenXmlNamespaces.PresentationML);
            this._writer.WriteStartElement("a", "sx", OpenXmlNamespaces.DrawingML);
            this._writer.WriteAttributeString("n", z.curScale.x.numer.ToString());
            this._writer.WriteAttributeString("d", z.curScale.x.denom.ToString());
            this._writer.WriteEndElement(); //sx
            this._writer.WriteStartElement("a", "sy", OpenXmlNamespaces.DrawingML);
            this._writer.WriteAttributeString("n", z.curScale.y.numer.ToString());
            this._writer.WriteAttributeString("d", z.curScale.y.denom.ToString());
            this._writer.WriteEndElement(); //sy
            this._writer.WriteEndElement(); //scale
            this._writer.WriteStartElement("p", "origin", OpenXmlNamespaces.PresentationML);
            this._writer.WriteAttributeString("x", z.origin.x.ToString());
            this._writer.WriteAttributeString("y", z.origin.y.ToString());
            this._writer.WriteEndElement(); //origin
            this._writer.WriteEndElement(); //cViewPr
            this._writer.WriteStartElement("p", "guideLst", OpenXmlNamespaces.PresentationML);
            this._writer.WriteStartElement("p", "guide", OpenXmlNamespaces.PresentationML);
            this._writer.WriteAttributeString("orient", "horz");
            this._writer.WriteAttributeString("pos", "2160");
            this._writer.WriteEndElement(); //guide
            this._writer.WriteStartElement("p", "guide", OpenXmlNamespaces.PresentationML);
            this._writer.WriteAttributeString("pos", "2880");
            this._writer.WriteEndElement(); //guide
            this._writer.WriteEndElement(); //guideLst
            this._writer.WriteEndElement(); //cSldViewPr
            this._writer.WriteEndElement(); //slideViewPr

            //_writer.WriteStartElement("p", "notesTextViewPr", OpenXmlNamespaces.PresentationML);
            //_writer.WriteStartElement("p", "cViewPr", OpenXmlNamespaces.PresentationML);
            //_writer.WriteStartElement("p", "scale", OpenXmlNamespaces.PresentationML);
            //_writer.WriteStartElement("a", "sx", OpenXmlNamespaces.DrawingML);
            //_writer.WriteAttributeString("n", "100");
            //_writer.WriteAttributeString("d", "100");
            //_writer.WriteEndElement(); //sx
            //_writer.WriteStartElement("a", "sy", OpenXmlNamespaces.DrawingML);
            //_writer.WriteAttributeString("n", "100");
            //_writer.WriteAttributeString("d", "100");
            //_writer.WriteEndElement(); //sy
            //_writer.WriteEndElement(); //scale
            //_writer.WriteStartElement("p", "origin", OpenXmlNamespaces.PresentationML);
            //_writer.WriteAttributeString("x", "0");
            //_writer.WriteAttributeString("y", "0");
            //_writer.WriteEndElement(); //origin
            //_writer.WriteEndElement(); //cViewPr
            //_writer.WriteEndElement(); //notesTextViewPr

            //_writer.WriteStartElement("p", "gridSpacing", OpenXmlNamespaces.PresentationML);
            //_writer.WriteAttributeString("cx", "73736200");
            //_writer.WriteAttributeString("cy", "73736200");
            //_writer.WriteEndElement(); //gridSpacing


            this._writer.WriteEndElement(); //viewPr
            this._writer.WriteEndDocument();

            this._writer.Flush();
        }        
    }
}
