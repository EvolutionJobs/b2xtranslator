using System.Collections.Generic;
using b2xtranslator.PptFileFormat;
using b2xtranslator.OpenXmlLib;
using b2xtranslator.OpenXmlLib.PresentationML;
using b2xtranslator.Tools;
using b2xtranslator.OfficeDrawing;

namespace b2xtranslator.PresentationMLMapping
{
    public class SlideMapping : PresentationMapping<RegularContainer>
    {
        public Slide Slide;
        public ShapeTreeMapping shapeTreeMapping;

        public SlideMapping(ConversionContext ctx)
            : base(ctx, ctx.Pptx.PresentationPart.AddSlidePart())
        {
        }

        /// <summary>
        /// Get the id of our real main master.
        /// 
        /// This need not be the id of our immediate master as it can be a title master.
        /// </summary>
        /// <param name="slideAtom">SlideAtom of slide to find main master id for</param>
        /// <returns>Id of main master</returns>
        private uint GetMainMasterId(SlideAtom slideAtom)
        {
            var masterSlide = this._ctx.Ppt.FindMasterRecordById(slideAtom.MasterId);
            
            // Is our immediate master a title master?
            if (!(masterSlide is MainMaster))
            {
                // Then our main master is the title master's master
                var titleSlideAtom = masterSlide.FirstChildWithType<SlideAtom>();
                return titleSlideAtom.MasterId;
            }

            return slideAtom.MasterId;
        }

        override public void Apply(RegularContainer slide)
        {            
            this.Slide = (Slide)slide;
            TraceLogger.DebugInternal("SlideMapping.Apply");

            // Associate slide with slide layout
            var slideAtom = slide.FirstChildWithType<SlideAtom>();
            uint mainMasterId = GetMainMasterId(slideAtom);
            var layoutManager = this._ctx.GetOrCreateLayoutManagerByMasterId(mainMasterId);

            SlideLayoutPart layoutPart = null;
            var masterInfo = slide.FirstChildWithType<RoundTripContentMasterId12>();

            // PPT2007 OOXML-Layout
            if (masterInfo != null)
            {
                layoutPart = layoutManager.GetLayoutPartByInstanceId(masterInfo.ContentMasterInstanceId);
            }
            // Pre-PPT2007 Title master layout
            else if (mainMasterId != slideAtom.MasterId)
            {
                layoutPart = layoutManager.GetOrCreateLayoutPartForTitleMasterId(slideAtom.MasterId);
            }
            // Pre-PPT2007 SSlideLayoutAtom primitive SlideLayoutType layout
            else
            {
                var m = (MainMaster)this._ctx.Ppt.FindMasterRecordById(slideAtom.MasterId);
                if (m.Layouts.Count == 1 && slideAtom.Layout.Geom == SlideLayoutType.Blank)
                {
                    foreach (string layout in m.Layouts.Values)
                    {
                        string output = Tools.Utils.replaceOutdatedNamespaces(layout);
                        layoutPart = layoutManager.GetOrCreateLayoutPartByCode(output);
                    }
                }
                else
                {
                    layoutPart = layoutManager.GetOrCreateLayoutPartByLayoutType(slideAtom.Layout.Geom, slideAtom.Layout.PlaceholderTypes);
                }
            }

            this.targetPart.ReferencePart(layoutPart);

            // Start the document
            this._writer.WriteStartDocument();
            this._writer.WriteStartElement("p", "sld", OpenXmlNamespaces.PresentationML);

            // Force declaration of these namespaces at document start
            this._writer.WriteAttributeString("xmlns", "a", null, OpenXmlNamespaces.DrawingML);
            // Force declaration of these namespaces at document start
            this._writer.WriteAttributeString("xmlns", "r", null, OpenXmlNamespaces.Relationships);


            if (Tools.Utils.BitmaskToBool(slideAtom.Flags, 0x1 << 0) == false)
            {
                this._writer.WriteAttributeString("showMasterSp", "0");
            }

            // TODO: Write slide data of master slide
            this._writer.WriteStartElement("p", "cSld", OpenXmlNamespaces.PresentationML);

            var sc = slide.FirstChildWithType<PPDrawing>().FirstChildWithType<DrawingContainer>().FirstChildWithType<ShapeContainer>();
            if (sc != null)
            {
                var sh = sc.FirstChildWithType<Shape>();
                var so = sc.FirstChildWithType<ShapeOptions>();                

                if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.FillStyleBooleanProperties))
                {
                    bool ignore = false;
                    if (sc.AllChildrenWithType<ShapeOptions>().Count > 1)
                    {
                        var so2 = sc.AllChildrenWithType<ShapeOptions>()[1];
                        if (so2.OptionsByID.ContainsKey(ShapeOptions.PropertyId.FillStyleBooleanProperties))
                        {
                            var p2 = new FillStyleBooleanProperties(so2.OptionsByID[ShapeOptions.PropertyId.FillStyleBooleanProperties].op);
                            if (!p2.fUsefFilled || !p2.fFilled) ignore = true;
                        }
                    }

                    var sa = slide.FirstChildWithType<SlideAtom>();
                    if (Tools.Utils.BitmaskToBool(sa.Flags, 0x1 << 2)) ignore = true; //this means the slide gets its background from the master

                    if (!ignore)
                    {
                        this._writer.WriteStartElement("p", "bg", OpenXmlNamespaces.PresentationML);
                        this._writer.WriteStartElement("p", "bgPr", OpenXmlNamespaces.PresentationML);
                        var p = new FillStyleBooleanProperties(so.OptionsByID[ShapeOptions.PropertyId.FillStyleBooleanProperties].op);
                        if (p.fUsefFilled & p.fFilled) //  so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.fillType))
                        {
                            new FillMapping(this._ctx, this._writer, this).Apply(so);
                        }
                        this._writer.WriteElementString("a", "effectLst", OpenXmlNamespaces.DrawingML, "");
                        this._writer.WriteEndElement(); //p:bgPr
                        this._writer.WriteEndElement(); //p:bg
                    }
                }                      
            }


            this._writer.WriteStartElement("p", "spTree", OpenXmlNamespaces.PresentationML);

            this.shapeTreeMapping = new ShapeTreeMapping(this._ctx, this._writer);
            this.shapeTreeMapping.parentSlideMapping = this;
            this.shapeTreeMapping.Apply(slide.FirstChildWithType<PPDrawing>());

            checkHeaderFooter(this.shapeTreeMapping);

            this._writer.WriteEndElement(); //spTree
            this._writer.WriteEndElement(); //cSld

            // TODO: Write clrMapOvr

            if (slide.FirstChildWithType<SlideShowSlideInfoAtom>() != null)
            {
                new SlideTransitionMapping(this._ctx, this._writer).Apply(slide.FirstChildWithType<SlideShowSlideInfoAtom>());
            }            

            if (slide.FirstChildWithType<ProgTags>() != null)
            if (slide.FirstChildWithType<ProgTags>().FirstChildWithType<ProgBinaryTag>() != null)
            if (slide.FirstChildWithType<ProgTags>().FirstChildWithType<ProgBinaryTag>().FirstChildWithType<ProgBinaryTagDataBlob>() != null)
            {
                new AnimationMapping(this._ctx, this._writer).Apply(slide.FirstChildWithType<ProgTags>().FirstChildWithType<ProgBinaryTag>().FirstChildWithType<ProgBinaryTagDataBlob>(), this, this.shapeTreeMapping.animinfos, this.shapeTreeMapping);
            }


            // End the document
            this._writer.WriteEndElement(); //sld
            this._writer.WriteEndDocument();

            this._writer.Flush();
        }

        private void InsertMasterStylePlaceholders(ShapeTreeMapping stm)
        {
            var slideAtom = this.Slide.FirstChildWithType<SlideAtom>();
            foreach (var master in this._ctx.Ppt.MainMasterRecords)
            {
                if (master.PersistAtom.SlideId == slideAtom.MasterId)
                {
                    var shapes = master.AllChildrenWithType<PPDrawing>()[0].AllChildrenWithType<OfficeDrawing.DrawingContainer>()[0].AllChildrenWithType<OfficeDrawing.GroupContainer>()[0].AllChildrenWithType<OfficeDrawing.ShapeContainer>();
                    foreach (var shapecontainer in shapes)
                    {
                        foreach (var data in shapecontainer.AllChildrenWithType<OfficeDrawing.ClientData>())
                        {
                            var ms = new System.IO.MemoryStream(data.bytes);
                            var rec = OfficeDrawing.Record.ReadRecord(ms);

                            if (rec.TypeCode == 3011)
                            {
                                var placeholder = (OEPlaceHolderAtom)rec;

                                if (placeholder != null)
                                {
                                    if (placeholder.PlacementId != PlaceholderEnum.MasterFooter)
                                    {
                                        stm.Apply(shapecontainer, "","","");
                                    }
                                }
                            }
                        }
                    }

                }
            }
        }

        public string readFooterFromClientTextBox(ClientTextbox textbox)
        {
            var ms = new System.IO.MemoryStream(textbox.Bytes);
            TextHeaderAtom thAtom = null;
            TextStyleAtom style = null;
            var lst = new List<int>();
            while (ms.Position < ms.Length)
            {
                var rec = Record.ReadRecord(ms);

                switch (rec.TypeCode)
                {
                    case 0xf9e: //OutlineTextRefAtom
                        var otrAtom = (OutlineTextRefAtom)rec;
                        var slideListWithText = this._ctx.Ppt.DocumentRecord.RegularSlideListWithText;

                        var thAtoms = slideListWithText.SlideToPlaceholderTextHeaders[textbox.FirstAncestorWithType<Slide>().PersistAtom];
                        thAtom = thAtoms[otrAtom.Index];

                        //if (thAtom.TextAtom != null) text = thAtom.TextAtom.Text;
                        if (thAtom.TextStyleAtom != null) style = thAtom.TextStyleAtom;
                        break;
                    case 0xf9f: //TextHeaderAtom
                        thAtom = (TextHeaderAtom)rec;
                        break;
                    case 0xfa0: //TextCharsAtom
                        thAtom.TextAtom = (TextAtom)rec;
                        break;
                    case 0xfa1: //StyleTextPropAtom
                        style = (TextStyleAtom)rec;
                        style.TextHeaderAtom = thAtom;
                        break;
                    case 0xfa2: //MasterTextPropAtom
                        break;
                    case 0xfa8: //TextBytesAtom
                        //text = ((TextBytesAtom)rec).Text;
                        thAtom.TextAtom = (TextAtom)rec;
                        return thAtom.TextAtom.Text;
                    case 0xfaa: //TextSpecialInfoAtom
                    case 0xfd8: //SlideNumberMCAtom
                    case 0xff9: //HeaderMCAtom
                        break;
                    case 0xffa: //FooterMCAtom
                        break;
                    case 0xff8: //GenericDateMCAtom                        
                        break;
                    default:
                        break;
                }
            }

            return "";
        }

        private void checkHeaderFooter(ShapeTreeMapping stm)
        {
            var slideAtom = this.Slide.FirstChildWithType<SlideAtom>();

            string footertext = "";
            string headertext = "";
            string userdatetext = "";
            var headersfooters = this.Slide.FirstChildWithType<SlideHeadersFootersContainer>();
            if (headersfooters != null)
            {
                foreach (var text in headersfooters.AllChildrenWithType<CStringAtom>())
                {
                    switch (text.Instance)
                    {
                        case 0:
                            userdatetext = text.Text;
                            break;
                        case 1:
                            headertext = text.Text;
                            break;
                        case 2:
                            footertext = text.Text;
                            break;
                    }
                }
                //CStringAtom text = headersfooters.FirstChildWithType<CStringAtom>();
                //if (text != null)
                //{
                //    footertext = text.Text;
                //}
            }

            bool footer = false;
            bool slideNumber = false;
            bool date = false;
            //bool userDate = false;
            if (!(this._ctx.Ppt.DocumentRecord.FirstChildWithType<DocumentAtom>().OmitTitlePlace && this.Slide.FirstChildWithType<SlideAtom>().Layout.Geom == SlideLayoutType.TitleSlide))
            foreach (var c in this._ctx.Ppt.DocumentRecord.AllChildrenWithType<SlideHeadersFootersContainer>())
            {
                switch (c.Instance)
                {
                    case 0: //PerSlideHeadersFootersContainer
                        break;
                    case 3: //SlideHeadersFootersContainer
                        foreach (var a in c.AllChildrenWithType<HeadersFootersAtom>())
                        {
                            if (a.fHasFooter) footer = true;
                            if (a.fHasSlideNumber) slideNumber = true;
                            if (a.fHasDate) date = true;
                            //if (a.fHasUserDate) userDate = true;

                            //if (a.fHasHeader) header = true;
                        }

                        if (footer && footertext.Length == 0 && c.FirstChildWithType<CStringAtom>() != null)
                        {
                            footertext = c.FirstChildWithType<CStringAtom>().Text;
                        }
                        break;
                    case 4: //NotesHeadersFootersContainer
                        break;
                }

            }

            //if (footertext.Length == 0) footer = false;

            if (slideNumber)
            {
                foreach (var master in this._ctx.Ppt.MainMasterRecords)
                {
                    if (master.PersistAtom.SlideId == slideAtom.MasterId)
                    {
                        var shapes = master.AllChildrenWithType<PPDrawing>()[0].AllChildrenWithType<OfficeDrawing.DrawingContainer>()[0].AllChildrenWithType<OfficeDrawing.GroupContainer>()[0].AllChildrenWithType<OfficeDrawing.ShapeContainer>();
                        foreach (var shapecontainer in shapes)
                        {
                            foreach (var data in shapecontainer.AllChildrenWithType<OfficeDrawing.ClientData>())
                            {
                                var ms = new System.IO.MemoryStream(data.bytes);
                                var rec = OfficeDrawing.Record.ReadRecord(ms);

                                if (rec.TypeCode == 3011)
                                {
                                    var placeholder = (OEPlaceHolderAtom)rec;

                                    if (placeholder != null)
                                    {
                                        if (placeholder.PlacementId == PlaceholderEnum.MasterSlideNumber)
                                        {
                                            stm.Apply(shapecontainer, "","","");
                                        }
                                    }
                                }
                            }
                        }

                    }
                }
            }

            if (date)
            {
                //if (!(userDate & userdatetext.Length == 0))
                //{
                    foreach (var master in this._ctx.Ppt.MainMasterRecords)
                    {
                        if (master.PersistAtom.SlideId == slideAtom.MasterId)
                        {
                            var shapes = master.AllChildrenWithType<PPDrawing>()[0].AllChildrenWithType<OfficeDrawing.DrawingContainer>()[0].AllChildrenWithType<OfficeDrawing.GroupContainer>()[0].AllChildrenWithType<OfficeDrawing.ShapeContainer>();
                            foreach (var shapecontainer in shapes)
                            {
                                foreach (var data in shapecontainer.AllChildrenWithType<OfficeDrawing.ClientData>())
                                {
                                    var ms = new System.IO.MemoryStream(data.bytes);
                                    var rec = OfficeDrawing.Record.ReadRecord(ms);

                                    if (rec.TypeCode == 3011)
                                    {
                                        var placeholder = (OEPlaceHolderAtom)rec;

                                        if (placeholder != null)
                                        {
                                            if (placeholder.PlacementId == PlaceholderEnum.MasterDate)
                                            {
                                                stm.Apply(shapecontainer, "", "","");
                                            }
                                        }
                                    }
                                }
                            }

                        }
                    }
               // }
            }

            if (footer)
            foreach (var master in this._ctx.Ppt.TitleMasterRecords)
            {
                if (master.PersistAtom.SlideId == slideAtom.MasterId)
                {
                    var shapes = master.AllChildrenWithType<PPDrawing>()[0].AllChildrenWithType<OfficeDrawing.DrawingContainer>()[0].AllChildrenWithType<OfficeDrawing.GroupContainer>()[0].AllChildrenWithType<OfficeDrawing.ShapeContainer>();
                    foreach (var shapecontainer in shapes)
                    {
                        foreach (var data in shapecontainer.AllChildrenWithType<OfficeDrawing.ClientData>())
                        {
                            var ms = new System.IO.MemoryStream(data.bytes);
                            var rec = OfficeDrawing.Record.ReadRecord(ms);

                            if (rec.TypeCode == 3011)
                            {
                                var placeholder = (OEPlaceHolderAtom)rec;

                                if (placeholder != null)
                                {
                                    if (placeholder.PlacementId == PlaceholderEnum.MasterFooter)
                                    {
                                        bool doit = footertext.Length > 0;
                                        if (!doit)
                                        {
                                            foreach (var so in shapecontainer.AllChildrenWithType<ShapeOptions>())
                                                if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.FillStyleBooleanProperties))
                                                {
                                                    var props = new FillStyleBooleanProperties(so.OptionsByID[ShapeOptions.PropertyId.FillStyleBooleanProperties].op);
                                                    if (props.fFilled && props.fUsefFilled) doit = true;
                                                }
                                        }
                                        if (doit) stm.Apply(shapecontainer, footertext, "", "");
                                        footer = false;
                                    }
                                }
                            }
                        }
                    }

                }
            }

            if (footer)
            foreach (var master in this._ctx.Ppt.MainMasterRecords)
                {
                    if (master.PersistAtom.SlideId == slideAtom.MasterId)
                    {
                        var shapes = master.AllChildrenWithType<PPDrawing>()[0].AllChildrenWithType<OfficeDrawing.DrawingContainer>()[0].AllChildrenWithType<OfficeDrawing.GroupContainer>()[0].AllChildrenWithType<OfficeDrawing.ShapeContainer>();
                        foreach (var shapecontainer in shapes)
                        {
                            foreach (var data in shapecontainer.AllChildrenWithType<OfficeDrawing.ClientData>())
                            {
                                var ms = new System.IO.MemoryStream(data.bytes);
                                var rec = OfficeDrawing.Record.ReadRecord(ms);

                                if (rec.TypeCode == 3011)
                                {
                                    var placeholder = (OEPlaceHolderAtom)rec;

                                    if (placeholder != null)
                                    {
                                        if (placeholder.PlacementId == PlaceholderEnum.MasterFooter)
                                        {
                                            if (footertext.Length == 0 & shapecontainer.AllChildrenWithType<ClientTextbox>().Count > 0) footertext = readFooterFromClientTextBox(shapecontainer.FirstChildWithType<ClientTextbox>());

                                            bool doit = footertext.Length > 0;
                                            if (!doit)
                                            {
                                                foreach(var so in shapecontainer.AllChildrenWithType<ShapeOptions>())
                                                    if (so.OptionsByID.ContainsKey(ShapeOptions.PropertyId.FillStyleBooleanProperties))
                                                    {
                                                        var props = new FillStyleBooleanProperties(so.OptionsByID[ShapeOptions.PropertyId.FillStyleBooleanProperties].op);
                                                        if (props.fFilled && props.fUsefFilled) doit = true;
                                                    }
                                            }
                                            if (doit) stm.Apply(shapecontainer, footertext, "","");
                                            footer = false;
                                        }
                                    }
                                }
                            }
                        }

                    }
                }
        }
    }
}
