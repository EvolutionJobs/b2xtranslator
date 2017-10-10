

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
    public class MasterMapping : PresentationMapping<RegularContainer>
    {
        public SlideMasterPart MasterPart;
        protected Slide Master;
        protected uint MasterId;
        protected MasterLayoutManager LayoutManager;

        public MasterMapping(ConversionContext ctx)
            : base(ctx, ctx.Pptx.PresentationPart.AddSlideMasterPart())
        {
            this.MasterPart = (SlideMasterPart)this.targetPart;
        }

        override public void Apply(RegularContainer pmaster)
        {
            var master = (Slide)pmaster;

            TraceLogger.DebugInternal("MasterMapping.Apply");
            uint masterId = master.PersistAtom.SlideId;
            this._ctx.RegisterMasterMapping(masterId, this);

            this.Master = master;
            this.MasterId = master.PersistAtom.SlideId;
            this.LayoutManager = this._ctx.GetOrCreateLayoutManagerByMasterId(this.MasterId);

            // Add PPT2007 roundtrip slide layouts
            var rtSlideLayouts = this.Master.AllChildrenWithType<RoundTripContentMasterInfo12>();

            foreach (var slideLayout in rtSlideLayouts)
            {
                var layoutPart = this.LayoutManager.AddLayoutPartWithInstanceId(slideLayout.Instance);
                XmlNode e = slideLayout.XmlDocumentElement;

                var nsm = new XmlNamespaceManager(new NameTable());
                nsm.AddNamespace("a", OpenXmlNamespaces.DrawingML);

                //for the moment remove blips that reference pictures
                foreach (XmlNode bublip in  e.SelectNodes("//a:buBlip",nsm))
                {
                    bublip.ParentNode.RemoveChild(bublip);
                }

                Tools.Utils.replaceOutdatedNamespaces(ref e);
                e.WriteTo(layoutPart.XmlWriter);
                layoutPart.XmlWriter.Flush();
            }
        }

        public void Write()
        {
            // Start the document
            this._writer.WriteStartDocument();
            this._writer.WriteStartElement("p", "sldMaster", OpenXmlNamespaces.PresentationML);
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

            // Write slide layout part id list
            this._writer.WriteStartElement("p", "sldLayoutIdLst", OpenXmlNamespaces.PresentationML);

            var layoutParts = this.LayoutManager.GetAllLayoutParts();

            // Master must have at least one SlideLayout or RepairDialog will appear
            if (layoutParts.Count == 0)
            {
                var layoutPart = this.LayoutManager.GetOrCreateLayoutPartByLayoutType(0, null);
                layoutParts.Add(layoutPart);
            }

            foreach (var slideLayoutPart in layoutParts)
            {
                this._writer.WriteStartElement("p", "sldLayoutId", OpenXmlNamespaces.PresentationML);
                this._writer.WriteAttributeString("r", "id", OpenXmlNamespaces.Relationships, slideLayoutPart.RelIdToString);
                this._writer.WriteEndElement();
            }

            this._writer.WriteEndElement();

            if (this.Master.FirstChildWithType<SlideShowSlideInfoAtom>() != null)
            {
                new SlideTransitionMapping(this._ctx, this._writer).Apply(this.Master.FirstChildWithType<SlideShowSlideInfoAtom>());
            }

            if (this.Master.FirstChildWithType<ProgTags>() != null)
                if (this.Master.FirstChildWithType<ProgTags>().FirstChildWithType<ProgBinaryTag>() != null)
                    if (this.Master.FirstChildWithType<ProgTags>().FirstChildWithType<ProgBinaryTag>().FirstChildWithType<ProgBinaryTagDataBlob>() != null)
                    {
                        new AnimationMapping(this._ctx, this._writer).Apply(this.Master.FirstChildWithType<ProgTags>().FirstChildWithType<ProgBinaryTag>().FirstChildWithType<ProgBinaryTagDataBlob>(), this, stm.animinfos, stm);
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

                new TextMasterStyleMapping(this._ctx, this._writer, this).Apply(this.Master);
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
                                
                //Tools.Utils.recursiveReplaceOutdatedNamespaces(ref xmlDoc);
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
