

using System;
using System.Collections.Generic;
using b2xtranslator.PptFileFormat;
using b2xtranslator.OfficeDrawing;
using b2xtranslator.OpenXmlLib;
using System.Xml;
using b2xtranslator.OpenXmlLib.PresentationML;
using b2xtranslator.Tools;

namespace b2xtranslator.PresentationMLMapping
{
    public class HandoutMasterMapping : PresentationMapping<RegularContainer>
    {
        public HandoutMasterPart MasterPart;
        protected Handout Master;
        protected uint MasterId;

        public HandoutMasterMapping(ConversionContext ctx)
            : base(ctx, ctx.Pptx.PresentationPart.AddHandoutMasterPart())
        {
            this.MasterPart = (HandoutMasterPart)this.targetPart;
        }

        override public void Apply(RegularContainer pmaster)
        {
            var master = (Handout)pmaster;

            TraceLogger.DebugInternal("HandoutMasterMapping.Apply");
            uint masterId = 0; // master.PersistAtom.SlideId;
            this._ctx.RegisterHandoutMasterMapping(masterId, this);

            this.Master = master;
            this.MasterId = 0; // master.PersistAtom.SlideId;
        }

        public void Write()
        {
            // Start the document
            this._writer.WriteStartDocument();
            this._writer.WriteStartElement("p", "handoutMaster", OpenXmlNamespaces.PresentationML);
            // Force declaration of these namespaces at document start
            this._writer.WriteAttributeString("xmlns", "a", null, OpenXmlNamespaces.DrawingML);
            this._writer.WriteAttributeString("xmlns", "r", null, OpenXmlNamespaces.Relationships);

            this._writer.WriteStartElement("p", "cSld", OpenXmlNamespaces.PresentationML);

            var sc = this.Master.FirstChildWithType<PPDrawing>().FirstChildWithType<DrawingContainer>().FirstChildWithType<ShapeContainer>();
            if (sc != null)
            {
                var sh = sc.FirstChildWithType<Shape>();
                var so = sc.FirstChildWithType<ShapeOptions>();
               
                if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillType))
                {
                    this._writer.WriteStartElement("p", "bg", OpenXmlNamespaces.PresentationML);
                    this._writer.WriteStartElement("p", "bgPr", OpenXmlNamespaces.PresentationML);
                    new FillMapping(this._ctx, this._writer, this).Apply(so);
                    this._writer.WriteElementString("a", "effectLst", OpenXmlNamespaces.DrawingML, "");
                    this._writer.WriteEndElement(); //p:bgPr
                    this._writer.WriteEndElement(); //p:bg
                }
                else if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillColor))
                {
                    string colorval;
                    if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillColor))
                    {
                        colorval = Utils.getRGBColorFromOfficeArtCOLORREF(so.OptionsByID[ShapeOptions.PropertyId.fillColor].op, this.Master, so);
                    }
                    else
                    {
                        colorval = "000000"; //TODO: find out which color to use in this case
                    }
                    this._writer.WriteStartElement("p", "bg", OpenXmlNamespaces.PresentationML);
                    this._writer.WriteStartElement("p", "bgPr", OpenXmlNamespaces.PresentationML);
                    this._writer.WriteStartElement("a", "solidFill", OpenXmlNamespaces.DrawingML);
                    this._writer.WriteStartElement("a", "srgbClr", OpenXmlNamespaces.DrawingML);
                    this._writer.WriteAttributeString("val", colorval);
                    if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillOpacity) && so.OptionsByID[ShapeOptions.PropertyId.fillOpacity].op != 65536)
                    {
                        this._writer.WriteStartElement("a", "alpha", OpenXmlNamespaces.DrawingML);
                        this._writer.WriteAttributeString("val", Math.Round(((decimal)so.OptionsByID[ShapeOptions.PropertyId.fillOpacity].op / 65536 * 100000)).ToString()); //we need the percentage of the opacity (65536 means 100%)
                        this._writer.WriteEndElement();
                    }
                    this._writer.WriteEndElement();
                    this._writer.WriteEndElement();
                    this._writer.WriteElementString("a", "effectLst", OpenXmlNamespaces.DrawingML, "");
                    this._writer.WriteEndElement(); //p:bgPr
                    this._writer.WriteEndElement(); //p:bg
                }
            }

            this._writer.WriteStartElement("p", "spTree", OpenXmlNamespaces.PresentationML);
            var stm = new ShapeTreeMapping(this._ctx, this._writer);
            stm.parentSlideMapping = this;
            stm.Apply(this.Master.FirstChildWithType<PPDrawing>());

            this._writer.WriteEndElement();
            this._writer.WriteEndElement();

            // Write clrMap
            var clrMap = this.Master.FirstChildWithType<ColorMappingAtom>();
            if (clrMap != null)
            {
                // clrMap from ColorMappingAtom wrongly uses namespace DrawingML
                this._writer.WriteStartElement("p", "clrMap", OpenXmlNamespaces.PresentationML);

                foreach (XmlAttribute attr in clrMap.XmlDocumentElement.Attributes)
                    if (attr.Prefix != "xmlns")
                        this._writer.WriteAttributeString(attr.LocalName, attr.Value);

                this._writer.WriteEndElement();
            }
            else
            {
                // In absence of ColorMappingAtom write default clrMap
                Utils.GetDefaultDocument("clrMap").WriteTo(this._writer);
            }


            // Write txStyles
            var roundTripTxStyles = this.Master.FirstChildWithType<RoundTripOArtTextStyles12>();
            if (false & roundTripTxStyles != null)
            {
                roundTripTxStyles.XmlDocumentElement.WriteTo(this._writer);
            }
            else
            {
                //throw new NotImplementedException("Write txStyles in case of PPT without roundTripTxStyles"); // TODO (pre PP2007)
                
                //XmlDocument slideLayoutDoc = Utils.GetDefaultDocument("txStyles");
                //slideLayoutDoc.WriteTo(_writer);

                //new TextMasterStyleMapping(_ctx, _writer, this).ApplyNotesMaster(this.Master);
            }

            // Write theme
            //
            // Note: We need to create a new theme part for each master,
            // even if it they have the same content.
            //
            // Otherwise PPT will complain about the structure of the file.
            var themePart = this._ctx.Pptx.PresentationPart.AddThemePart();

            XmlNode xmlDoc;
            var theme = this.Master.FirstChildWithType<Theme>();

            if (theme != null)
            {
                xmlDoc = theme.XmlDocumentElement;
                xmlDoc.WriteTo(themePart.XmlWriter);
            }
            else
            {
                var schemes = this.Master.AllChildrenWithType<ColorSchemeAtom>();
                if (schemes.Count > 0)
                {
                    new ColorSchemeMapping(this._ctx, themePart.XmlWriter).Apply(schemes);                    
                }
                else
                {
                    // In absence of Theme record use default theme
                    xmlDoc = Utils.GetDefaultDocument("theme");
                    xmlDoc.WriteTo(themePart.XmlWriter);
                }
            }

            themePart.XmlWriter.Flush();
           

            this.MasterPart.ReferencePart(themePart);

            // End the document
            this._writer.WriteEndElement();
            this._writer.WriteEndDocument();

            this._writer.Flush();
        }
    }
}
