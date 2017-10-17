

using System.Collections.Generic;
using b2xtranslator.PptFileFormat;
using b2xtranslator.OpenXmlLib;
using b2xtranslator.Tools;
using b2xtranslator.OfficeDrawing;

namespace b2xtranslator.PresentationMLMapping
{
    public class NoteMapping : PresentationMapping<RegularContainer>
    {
        public Note Note;
        private SlideMapping SlideMapping;

        public NoteMapping(ConversionContext ctx, SlideMapping slideMapping)
            : base(ctx, ctx.Pptx.PresentationPart.AddNotePart())
        {
            this.SlideMapping = slideMapping;
        }


        override public void Apply(RegularContainer note)
        {
            this.Note = (Note)note;
            TraceLogger.DebugInternal("NoteMapping.Apply");

            // Associate slide with slide layout
            var notesAtom = note.FirstChildWithType<NotesAtom>();

            //Add relationship to slide
            this.targetPart.ReferencePart(this.SlideMapping.targetPart);
            this.SlideMapping.targetPart.ReferencePart(this.targetPart);

            //Add relationship to notes master


            // Start the document
            this._writer.WriteStartDocument();
            this._writer.WriteStartElement("p", "notes", OpenXmlNamespaces.PresentationML);

            // Force declaration of these namespaces at document start
            this._writer.WriteAttributeString("xmlns", "a", null, OpenXmlNamespaces.DrawingML);
            // Force declaration of these namespaces at document start
            this._writer.WriteAttributeString("xmlns", "r", null, OpenXmlNamespaces.Relationships);

            // TODO: Write slide data of master slide
            this._writer.WriteStartElement("p", "cSld", OpenXmlNamespaces.PresentationML);

            //TODO: write background properties (p:bg)

            this._writer.WriteStartElement("p", "spTree", OpenXmlNamespaces.PresentationML);

            var stm = new ShapeTreeMapping(this._ctx, this._writer);
            stm.parentSlideMapping = this;
            stm.Apply(note.FirstChildWithType<PPDrawing>());

            checkHeaderFooter(stm);

            this._writer.WriteEndElement(); //spTree
            this._writer.WriteEndElement(); //cSld

            // TODO: Write clrMapOvr

            // End the document
            this._writer.WriteEndElement(); //sld
            this._writer.WriteEndDocument();

            this._writer.Flush();

        }

        private void checkHeaderFooter(ShapeTreeMapping stm)
        {
            var slideAtom = this.Note.FirstChildWithType<NotesAtom>();

            string footertext = "";
            string headertext = "";
            string userdatetext = "";
            var headersfooters = this.Note.FirstChildWithType<SlideHeadersFootersContainer>();
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
            //bool header = false;
            bool slideNumber = false;
            //bool date = false;
            //bool userDate = false;

            foreach (var c in this._ctx.Ppt.DocumentRecord.AllChildrenWithType<SlideHeadersFootersContainer>())
            {
                switch (c.Instance)
                {
                    case 0: //PerSlideHeadersFootersContainer
                        break;
                    case 3: //SlideHeadersFootersContainer
                        break;
                    case 4: //NotesHeadersFootersContainer
                        foreach (var a in c.AllChildrenWithType<HeadersFootersAtom>())
                        {
                            if (a.fHasFooter) footer = true;
                            //if (a.fHasHeader) header = true;
                            if (a.fHasSlideNumber) slideNumber = true;
                            //if (a.fHasDate) date = true;
                            //if (a.fHasUserDate) userDate = true;
                        }
                        break;
                }

            }

            var master = this._ctx.Ppt.NotesMasterRecords[0];

            if (slideNumber)
            {
                //foreach (Slide master in this._ctx.Ppt.MainMasterRecords)
                //{
                //if (master.PersistAtom.SlideId == slideAtom.MasterId)
                //{
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
                                    stm.Apply(shapecontainer, "", "", "");
                                }
                            }
                        }
                    }
                }

                //}
                //}
            }

            //if (date)
            //{
            //if (!(userDate & userdatetext.Length == 0))
            //    {
            //    foreach (Slide master in this._ctx.Ppt.MainMasterRecords)
            //    {
            //        if (master.PersistAtom.SlideId == slideAtom.MasterId)
            //        {
            //            List<OfficeDrawing.ShapeContainer> shapes = master.AllChildrenWithType<PPDrawing>()[0].AllChildrenWithType<OfficeDrawing.DrawingContainer>()[0].AllChildrenWithType<OfficeDrawing.GroupContainer>()[0].AllChildrenWithType<OfficeDrawing.ShapeContainer>();
            //            foreach (OfficeDrawing.ShapeContainer shapecontainer in shapes)
            //            {
            //                foreach (OfficeDrawing.ClientData data in shapecontainer.AllChildrenWithType<OfficeDrawing.ClientData>())
            //                {
            //                    System.IO.MemoryStream ms = new System.IO.MemoryStream(data.bytes);
            //                    OfficeDrawing.Record rec = OfficeDrawing.Record.ReadRecord(ms, 0);

            //                    if (rec.TypeCode == 3011)
            //                    {
            //                        OEPlaceHolderAtom placeholder = (OEPlaceHolderAtom)rec;

            //                        if (placeholder != null)
            //                        {
            //                            if (placeholder.PlacementId == PlaceholderEnum.MasterDate)
            //                            {
            //                                stm.Apply(shapecontainer, "");
            //                            }
            //                        }
            //                    }
            //                }
            //            }

            //        }
            //    }
            //  }
            //}

            //if (header)
            //{
            //    string s = "";
            //}

            // if (footer)
            // //foreach (Note master in this._ctx.Ppt.NotesMasterRecords)
            // {                    
            //     List<OfficeDrawing.ShapeContainer> shapes = master.AllChildrenWithType<PPDrawing>()[0].AllChildrenWithType<OfficeDrawing.DrawingContainer>()[0].AllChildrenWithType<OfficeDrawing.GroupContainer>()[0].AllChildrenWithType<OfficeDrawing.ShapeContainer>();
            //     foreach (OfficeDrawing.ShapeContainer shapecontainer in shapes)
            //     {
            //         foreach (OfficeDrawing.ClientData data in shapecontainer.AllChildrenWithType<OfficeDrawing.ClientData>())
            //         {
            //             System.IO.MemoryStream ms = new System.IO.MemoryStream(data.bytes);
            //             OfficeDrawing.Record rec = OfficeDrawing.Record.ReadRecord(ms);

            //             if (rec.TypeCode == 3011)
            //             {
            //                 OEPlaceHolderAtom placeholder = (OEPlaceHolderAtom)rec;

            //                 if (placeholder != null)
            //                 {
            //                     if (placeholder.PlacementId == PlaceholderEnum.MasterFooter)
            //                     {
            //                         stm.Apply(shapecontainer, footertext, "", "");
            //                         footer = false;
            //                     }
            //                 }
            //             }
            //         }
            //     }                    
            //}

            if (footer)
            //foreach (Slide master in this._ctx.Ppt.TitleMasterRecords)
            {
                //if (master.PersistAtom.SlideId == slideAtom.MasterId)
                //{
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

                //}
            }

            if (footer)
            //foreach (Slide master in this._ctx.Ppt.MainMasterRecords)
            {
                //if (master.PersistAtom.SlideId == slideAtom.MasterId)
                //{
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
                                    if (footertext.Length == 0 & shapecontainer.AllChildrenWithType<ClientTextbox>().Count > 0) footertext = new SlideMapping(this._ctx).readFooterFromClientTextBox(shapecontainer.FirstChildWithType<ClientTextbox>());

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

        //}
    }
}
