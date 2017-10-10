

using System.Collections.Generic;
using b2xtranslator.PptFileFormat;
using b2xtranslator.OpenXmlLib;

namespace b2xtranslator.PresentationMLMapping
{
    public class PresentationPartMapping : PresentationMapping<PowerpointDocument>
    {
        public List<SlideMapping> SlideMappings = new List<SlideMapping>();
        private List<NoteMapping> NoteMappings = new List<NoteMapping>();

        public PresentationPartMapping(ConversionContext ctx)
            : base(ctx, ctx.Pptx.PresentationPart)
        {
        }

        public override void Apply(PowerpointDocument ppt)
        {
            var documentRecord = ppt.DocumentRecord;

            // Start the document
            this._writer.WriteStartDocument();
            this._writer.WriteStartElement("p", "presentation", OpenXmlNamespaces.PresentationML);

            // Force declaration of these namespaces at document start
            this._writer.WriteAttributeString("xmlns", "r", null, OpenXmlNamespaces.Relationships);

            CreateMainMasters(ppt);
            CreateNotesMasters(ppt);
            CreateHandoutMasters(ppt);
            CreateVbaProject(ppt);
            CreateSlides(ppt, documentRecord);
                        
            WriteMainMasters(ppt);
            WriteSlides(ppt, documentRecord);

            var viewProps = new viewPropsMapping(this._ctx.Pptx.PresentationPart.AddViewPropertiesPart(), this._ctx.WriterSettings, this._ctx);
            viewProps.Apply(null);

            // sldSz and notesSz
            WriteSizeInfo(ppt, documentRecord);

            WriteDefaultTextStyle(ppt, documentRecord);

            // End the document
            this._writer.WriteEndElement();
            this._writer.WriteEndDocument();

            this._writer.Flush();
        }

        private void WriteDefaultTextStyle(PowerpointDocument ppt, DocumentContainer documentRecord)
        {
            this._writer.WriteStartElement("p", "defaultTextStyle", OpenXmlNamespaces.PresentationML);

            this._writer.WriteStartElement("a", "defPPr", OpenXmlNamespaces.DrawingML);
            this._writer.WriteStartElement("a", "defRPr", OpenXmlNamespaces.DrawingML);
            this._writer.WriteAttributeString("lang", "en-US");
            this._writer.WriteEndElement(); //defRPr
            this._writer.WriteEndElement(); //defPPr


            var defaultStyle = this._ctx.Ppt.DocumentRecord.FirstChildWithType<b2xtranslator.PptFileFormat.Environment>().FirstChildWithType<TextMasterStyleAtom>();

            var map = new TextMasterStyleMapping(this._ctx, this._writer, null);
            
            for (int i = 0; i < defaultStyle.IndentLevelCount; i++)
            {
                map.writepPr(defaultStyle.CRuns[i], defaultStyle.PRuns[i], null, i, false, true);
            }
            for (int i = defaultStyle.IndentLevelCount; i < 9; i++)
            {
                map.writepPr(defaultStyle.CRuns[0], defaultStyle.PRuns[0], null, i, false, true);
            }


            this._writer.WriteEndElement(); //defaultTextStyle
        }

        private void WriteSizeInfo(PowerpointDocument ppt, DocumentContainer documentRecord)
        {
            var doc = documentRecord.FirstChildWithType<DocumentAtom>();

            // Write slide size and type
            WriteSlideSizeInfo(doc);

            // Write notes size
            WriteNotesSizeInfo(doc);
        }

        private void WriteNotesSizeInfo(DocumentAtom doc)
        {
            int notesWidth = Utils.MasterCoordToEMU(doc.NotesSize.X);
            int notesHeight = Utils.MasterCoordToEMU(doc.NotesSize.Y);

            this._writer.WriteStartElement("p", "notesSz", OpenXmlNamespaces.PresentationML);

            this._writer.WriteAttributeString("cx", notesWidth.ToString());
            this._writer.WriteAttributeString("cy", notesHeight.ToString());

            this._writer.WriteEndElement();
        }

        private void WriteSlideSizeInfo(DocumentAtom doc)
        {
            int slideWidth = Utils.MasterCoordToEMU(doc.SlideSize.X);
            int slideHeight = Utils.MasterCoordToEMU(doc.SlideSize.Y);
            string slideType = Utils.SlideSizeTypeToXMLValue(doc.SlideSizeType);

            this._writer.WriteStartElement("p", "sldSz", OpenXmlNamespaces.PresentationML);

            this._writer.WriteAttributeString("cx", slideWidth.ToString());
            this._writer.WriteAttributeString("cy", slideHeight.ToString());
            this._writer.WriteAttributeString("type", slideType);

            this._writer.WriteEndElement();

        }

       
       private void CreateSlides(PowerpointDocument ppt, DocumentContainer documentRecord)
        {
            foreach (var lst in ppt.DocumentRecord.AllChildrenWithType<SlideListWithText>())
            {
                if (lst.Instance == 0)
                {
                    foreach (var at in lst.SlidePersistAtoms)
                    {
                        foreach (var slide in ppt.SlideRecords)
                        {
                            if (slide.PersistAtom == at)
                            {
                                var sMapping = new SlideMapping(this._ctx);
                                sMapping.Apply(slide);
                                this.SlideMappings.Add(sMapping);
                            }
                        }


                    }
                }
                //bool found = false;
                if (lst.Instance == 2) //notes
                {
                    foreach (var at in lst.SlidePersistAtoms)
                    {
                        //found = false;
                        foreach (var note in ppt.NoteRecords)
                        {
                            if (note.PersistAtom.SlideId == at.SlideId)
                            {
                                var a = note.FirstChildWithType<NotesAtom>();
                                foreach (var slideMapping in this.SlideMappings)
                                {
                                    if (slideMapping.Slide.PersistAtom.SlideId == a.SlideIdRef)
                                    {
                                        var nMapping = new NoteMapping(this._ctx, slideMapping);
                                        nMapping.Apply(note);
                                        this.NoteMappings.Add(nMapping);
                                        //found = true;
                                    }
                                }
                                
                            }
                        }
                        //if (!found)
                        //{
                        //    string s = "";
                        //}

                    }
                }
            }

       }

        private void WriteSlides(PowerpointDocument ppt, DocumentContainer documentRecord)
        {
            this._writer.WriteStartElement("p", "sldIdLst", OpenXmlNamespaces.PresentationML);

            foreach (var sMapping in this.SlideMappings)
            {
                WriteSlide(sMapping);
            }

            this._writer.WriteEndElement();
        }

        private void WriteSlide(SlideMapping sMapping)
        {
            var slide = sMapping.Slide;

            this._writer.WriteStartElement("p", "sldId", OpenXmlNamespaces.PresentationML);

            var slideAtom = slide.FirstChildWithType<SlideAtom>();

            this._writer.WriteAttributeString("id", slide.PersistAtom.SlideId.ToString());
            this._writer.WriteAttributeString("r", "id", OpenXmlNamespaces.Relationships, sMapping.targetPart.RelIdToString);

            this._writer.WriteEndElement();
        }

        private void CreateVbaProject(PowerpointDocument ppt)
        {
            if (ppt.VbaProject != null)
            {
                ppt.VbaProject.Convert(new VbaProjectMapping(this._ctx.Pptx.PresentationPart.VbaProjectPart));
            }
        }

        private void CreateMainMasters(PowerpointDocument ppt)
        {
            foreach (var m in ppt.MainMasterRecords)
            {
                this._ctx.GetOrCreateMasterMappingByMasterId(m.PersistAtom.SlideId).Apply(m);
            }
        }

        private void CreateNotesMasters(PowerpointDocument ppt)
        {
            foreach (var m in ppt.NotesMasterRecords)
            {
                this._ctx.GetOrCreateNotesMasterMappingByMasterId(0).Apply(m);
            }
        }

        private void CreateHandoutMasters(PowerpointDocument ppt)
        {
            foreach (var m in ppt.HandoutMasterRecords)
            {
                this._ctx.GetOrCreateHandoutMasterMappingByMasterId(0).Apply(m);
            }
        }

        private void WriteMainMasters(PowerpointDocument ppt)
        {
            this._writer.WriteStartElement("p", "sldMasterIdLst", OpenXmlNamespaces.PresentationML);

            foreach (var m in ppt.MainMasterRecords)
            {
                this.WriteMainMaster(ppt, m);
            }

            this._writer.WriteEndElement();

            WriteNoteMaster(ppt);
            WriteHandoutMaster(ppt);
        }

        /// <summary>
        /// Writes a slide master.
        /// 
        /// A slide master can either be a main master (type MainMaster) or title master (type Slide).
        /// <param name="ppt">PowerpointDocument record</param>
        /// <param name="m">Main master record</param>
        private void WriteMainMaster(PowerpointDocument ppt, MainMaster m)
        {
            this._writer.WriteStartElement("p", "sldMasterId", OpenXmlNamespaces.PresentationML);

            var mapping = this._ctx.GetOrCreateMasterMappingByMasterId(m.PersistAtom.SlideId);
            mapping.Write();

            string relString = mapping.targetPart.RelIdToString;

            this._writer.WriteAttributeString("r", "id", OpenXmlNamespaces.Relationships, relString);

            this._writer.WriteEndElement();
        }

        private void WriteNoteMaster(PowerpointDocument ppt)
        {
            if (ppt.NotesMasterRecords.Count > 0)
            {
                this._writer.WriteStartElement("p", "notesMasterIdLst", OpenXmlNamespaces.PresentationML);

                foreach (var m in ppt.NotesMasterRecords)
                {
                    this.WriteNoteMaster2(ppt, m);
                }

                this._writer.WriteEndElement();
            }
        }

        /// <summary>
        /// Writes a notes master.
        /// 
        /// <param name="ppt">PowerpointDocument record</param>
        /// <param name="m">Notes master record</param>
        private void WriteNoteMaster2(PowerpointDocument ppt, Note m)
        {
            this._writer.WriteStartElement("p", "notesMasterId", OpenXmlNamespaces.PresentationML);

            var mapping = this._ctx.GetOrCreateNotesMasterMappingByMasterId(0);
            mapping.Write();

            string relString = mapping.targetPart.RelIdToString;

            this._writer.WriteAttributeString("r", "id", OpenXmlNamespaces.Relationships, relString);

            this._writer.WriteEndElement();
        }

        private void WriteHandoutMaster(PowerpointDocument ppt)
        {
            if (ppt.HandoutMasterRecords.Count > 0)
            {

                this._writer.WriteStartElement("p", "handoutMasterIdLst", OpenXmlNamespaces.PresentationML);

                foreach (var m in ppt.HandoutMasterRecords)
                {
                    this.WriteHandoutMaster2(ppt, m);
                }

                this._writer.WriteEndElement();
            }
        }

        /// <summary>
        /// Writes a handout master.
        /// 
        /// <param name="ppt">PowerpointDocument record</param>
        /// <param name="m">Handout master record</param>
        private void WriteHandoutMaster2(PowerpointDocument ppt, Handout m)
        {
            this._writer.WriteStartElement("p", "handoutMasterId", OpenXmlNamespaces.PresentationML);

            var mapping = this._ctx.GetOrCreateHandoutMasterMappingByMasterId(0);
            mapping.Write();

            string relString = mapping.targetPart.RelIdToString;

            this._writer.WriteAttributeString("r", "id", OpenXmlNamespaces.Relationships, relString);

            this._writer.WriteEndElement();
        }
    }
}
