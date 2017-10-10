/*
 * Copyright (c) 2008, DIaLOGIKa
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *     * Redistributions of source code must retain the above copyright
 *        notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright
 *       notice, this list of conditions and the following disclaimer in the
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the name of DIaLOGIKa nor the
 *       names of its contributors may be used to endorse or promote products
 *       derived from this software without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY DIaLOGIKa ''AS IS'' AND ANY
 * EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 * DISCLAIMED. IN NO EVENT SHALL DIaLOGIKa BE LIABLE FOR ANY
 * DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 * (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 * ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 * (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */

using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.PptFileFormat;
using DIaLOGIKa.b2xtranslator.OpenXmlLib;
using System.Xml;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.PresentationML;
using DIaLOGIKa.b2xtranslator.Tools;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;

namespace DIaLOGIKa.b2xtranslator.PresentationMLMapping
{
    public class NoteMapping : PresentationMapping<RegularContainer>
    {
        public Note Note;
        private SlideMapping SlideMapping;

        public NoteMapping(ConversionContext ctx, SlideMapping slideMapping)
            : base(ctx, ctx.Pptx.PresentationPart.AddNotePart())
        {
            SlideMapping = slideMapping;
        }

       
        override public void Apply(RegularContainer note)
        {            
            this.Note = (Note)note;
            TraceLogger.DebugInternal("NoteMapping.Apply");

            // Associate slide with slide layout
            NotesAtom notesAtom = note.FirstChildWithType<NotesAtom>();
            
            //Add relationship to slide
            this.targetPart.ReferencePart(SlideMapping.targetPart);
            SlideMapping.targetPart.ReferencePart(this.targetPart);

            //Add relationship to notes master

            
            // Start the document
            _writer.WriteStartDocument();
            _writer.WriteStartElement("p", "notes", OpenXmlNamespaces.PresentationML);
            
            // Force declaration of these namespaces at document start
            _writer.WriteAttributeString("xmlns", "a", null, OpenXmlNamespaces.DrawingML);
            // Force declaration of these namespaces at document start
            _writer.WriteAttributeString("xmlns", "r", null, OpenXmlNamespaces.Relationships);

            // TODO: Write slide data of master slide
            _writer.WriteStartElement("p", "cSld", OpenXmlNamespaces.PresentationML);

            //TODO: write background properties (p:bg)

            _writer.WriteStartElement("p", "spTree", OpenXmlNamespaces.PresentationML);

            var stm = new ShapeTreeMapping(_ctx, _writer);
            stm.parentSlideMapping = this;
            stm.Apply(note.FirstChildWithType<PPDrawing>());

            checkHeaderFooter(stm);
         
            _writer.WriteEndElement(); //spTree
            _writer.WriteEndElement(); //cSld

            // TODO: Write clrMapOvr

            // End the document
            _writer.WriteEndElement(); //sld
            _writer.WriteEndDocument();

            _writer.Flush();
                
        }

        private void checkHeaderFooter(ShapeTreeMapping stm)
        {
            NotesAtom slideAtom = this.Note.FirstChildWithType<NotesAtom>();

            string footertext = "";
            string headertext = "";
            string userdatetext = "";
            SlideHeadersFootersContainer headersfooters = this.Note.FirstChildWithType<SlideHeadersFootersContainer>();
            if (headersfooters != null)
            {
                foreach (CStringAtom text in headersfooters.AllChildrenWithType<CStringAtom>())
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
            bool header = false;
            bool slideNumber = false;
            bool date = false;
            bool userDate = false;

            foreach (SlideHeadersFootersContainer c in this._ctx.Ppt.DocumentRecord.AllChildrenWithType<SlideHeadersFootersContainer>())
            {
                switch (c.Instance)
                {
                    case 0: //PerSlideHeadersFootersContainer
                        break;
                    case 3: //SlideHeadersFootersContainer
                         break;
                    case 4: //NotesHeadersFootersContainer
                        foreach (HeadersFootersAtom a in c.AllChildrenWithType<HeadersFootersAtom>())
                        {
                            if (a.fHasFooter) footer = true;
                            if (a.fHasHeader) header = true;
                            if (a.fHasSlideNumber) slideNumber = true;
                            if (a.fHasDate) date = true;
                            if (a.fHasUserDate) userDate = true;
                        }
                        break;
                }

            }

            Note master = _ctx.Ppt.NotesMasterRecords[0];

            if (slideNumber)
            {
                //foreach (Slide master in this._ctx.Ppt.MainMasterRecords)
                //{
                    //if (master.PersistAtom.SlideId == slideAtom.MasterId)
                    //{
                        List<OfficeDrawing.ShapeContainer> shapes = master.AllChildrenWithType<PPDrawing>()[0].AllChildrenWithType<OfficeDrawing.DrawingContainer>()[0].AllChildrenWithType<OfficeDrawing.GroupContainer>()[0].AllChildrenWithType<OfficeDrawing.ShapeContainer>();
                        foreach (OfficeDrawing.ShapeContainer shapecontainer in shapes)
                        {
                            foreach (OfficeDrawing.ClientData data in shapecontainer.AllChildrenWithType<OfficeDrawing.ClientData>())
                            {
                                var ms = new System.IO.MemoryStream(data.bytes);
                                OfficeDrawing.Record rec = OfficeDrawing.Record.ReadRecord(ms);

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

            if (header)
            {
                string s = "";
            }

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
                    List<OfficeDrawing.ShapeContainer> shapes = master.AllChildrenWithType<PPDrawing>()[0].AllChildrenWithType<OfficeDrawing.DrawingContainer>()[0].AllChildrenWithType<OfficeDrawing.GroupContainer>()[0].AllChildrenWithType<OfficeDrawing.ShapeContainer>();
                    foreach (OfficeDrawing.ShapeContainer shapecontainer in shapes)
                    {
                        foreach (OfficeDrawing.ClientData data in shapecontainer.AllChildrenWithType<OfficeDrawing.ClientData>())
                        {
                            var ms = new System.IO.MemoryStream(data.bytes);
                            OfficeDrawing.Record rec = OfficeDrawing.Record.ReadRecord(ms);

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
                                            foreach (ShapeOptions so in shapecontainer.AllChildrenWithType<ShapeOptions>())
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
                        List<OfficeDrawing.ShapeContainer> shapes = master.AllChildrenWithType<PPDrawing>()[0].AllChildrenWithType<OfficeDrawing.DrawingContainer>()[0].AllChildrenWithType<OfficeDrawing.GroupContainer>()[0].AllChildrenWithType<OfficeDrawing.ShapeContainer>();
                        foreach (OfficeDrawing.ShapeContainer shapecontainer in shapes)
                        {
                            foreach (OfficeDrawing.ClientData data in shapecontainer.AllChildrenWithType<OfficeDrawing.ClientData>())
                            {
                                var ms = new System.IO.MemoryStream(data.bytes);
                                OfficeDrawing.Record rec = OfficeDrawing.Record.ReadRecord(ms);

                                if (rec.TypeCode == 3011)
                                {
                                    var placeholder = (OEPlaceHolderAtom)rec;

                                    if (placeholder != null)
                                    {
                                        if (placeholder.PlacementId == PlaceholderEnum.MasterFooter)
                                        {
                                            if (footertext.Length == 0 & shapecontainer.AllChildrenWithType<ClientTextbox>().Count > 0) footertext = new SlideMapping(_ctx).readFooterFromClientTextBox(shapecontainer.FirstChildWithType<ClientTextbox>());

                                            bool doit = footertext.Length > 0;
                                            if (!doit)
                                            {
                                                foreach(ShapeOptions so in shapecontainer.AllChildrenWithType<ShapeOptions>())
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
                        
        //}
    }
}
