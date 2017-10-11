using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.OfficeDrawing;
using b2xtranslator.StructuredStorage.Common;
using b2xtranslator.StructuredStorage.Reader;
using System;
using System.Collections.Generic;
using System.Reflection;

namespace b2xtranslator.DocFileFormat
{
    public class WordDocument : IVisitable
    {
        static WordDocument()
        {
            Record.UpdateTypeToRecordClassMapping(Assembly.GetExecutingAssembly(), typeof(WordDocument).Namespace);
        }

        /// <summary>
        /// A dictionary that contains all SEPX of the document.<br/>
        /// The key is the CP at which sections ends.<br/>
        /// The value is the SEPX that formats the section.
        /// </summary>
        public Dictionary<int, SectionPropertyExceptions> AllSepx;

        /// <summary>
        /// A dictionary that contains all PAPX of the document.<br/>
        /// The key is the FC at which the paragraph starts.<br/>
        /// The value is the PAPX that formats the paragraph.
        /// </summary>
        public Dictionary<int, ParagraphPropertyExceptions> AllPapx;

        /// <summary>
        /// 
        /// </summary>
        public PieceTable PieceTable;

        public CommandTable CommandTable;

        /// <summary>
        /// A Plex containing all section descriptors
        /// </summary>
        public Plex<SectionDescriptor> SectionPlex;

        /// <summary>
        /// Contains the names of all author who revised something in the document
        /// </summary>
        public StringTable RevisionAuthorTable;

        /// <summary>
        /// The stream "WordDocument"
        /// </summary>
        public VirtualStream WordDocumentStream;

        /// <summary>
        /// The stream "0Table" or "1Table"
        /// </summary>
        public VirtualStream TableStream;

        /// <summary>
        /// The stream called "Data"
        /// </summary>
        public VirtualStream DataStream;

        /// <summary>
        /// The StructuredStorageFile itself
        /// </summary>
        public StructuredStorageReader Storage;

        /// <summary>
        /// The file information block of the word document
        /// </summary>
        public FileInformationBlock FIB;

        /// <summary>
        /// All text of the Word document
        /// </summary>
        public List<char> Text;

        /// <summary>
        /// The style sheet of the document
        /// </summary>
        public StyleSheet Styles;

        /// <summary>
        /// A list of all font names, used in the doucument
        /// </summary>
        public StringTable FontTable;

        public StringTable BookmarkNames;

        public StringTable AutoTextNames;

        //public StringTable ProtectionUsers;

        /// <summary>
        /// A plex with all ATRDPre10 structs
        /// </summary>
        public Plex<AnnotationReferenceDescriptor> AnnotationsReferencePlex;

        public AnnotationOwnerList AnnotationOwners;

        /// <summary>
        /// An array with all ATRDPost10 structs
        /// </summary>
        public AnnotationReferenceExtraTable AnnotationReferenceExtraTable;

        /// <summary>
        /// A list that contains all formatting information of 
        /// the lists and numberings in the document.
        /// </summary>
        public ListTable ListTable;

        /// <summary>
        /// The drawing object table ....
        /// </summary>
        public OfficeArtContent OfficeArtContent;

        /// <summary>
        /// 
        /// </summary>
        public Plex<FileShapeAddress> OfficeDrawingPlex;

        /// <summary>
        /// 
        /// </summary>
        public Plex<FileShapeAddress> OfficeDrawingPlexHeader;

        /// <summary>
        /// Each character position specifies the beginning of a range of text 
        /// that constitutes the contents of an AutoText item.
        /// </summary>
        public Plex<Exception> AutoTextPlex;

        public Plex<short> EndnoteReferencePlex;
        public Plex<short> FootnoteReferencePlex;

        /// <summary>
        /// Describes the breaks inside the textbox subdocument
        /// </summary>
        public Plex<BreakDescriptor> TextboxBreakPlex;

        /// <summary>
        /// Describes the breaks inside the header textbox subdocument
        /// </summary>
        public Plex<BreakDescriptor> TextboxBreakPlexHeader;

        public Plex<BookmarkFirst> BookmarkStartPlex;

        public Plex<Exception> BookmarkEndPlex;

        /// <summary>
        /// The DocumentProperties of the word document
        /// </summary>
        public DocumentProperties DocumentProperties;

        /// <summary>
        /// A list that contains all overriding formatting information
        /// of the lists and numberings in the document.
        /// </summary>
        public ListFormatOverrideTable ListFormatOverrideTable;

        /// <summary>
        /// A list of all FKPs that contain PAPX
        /// </summary>
        public List<FormattedDiskPagePAPX> AllPapxFkps;

        /// <summary>
        /// A list of all FKPs that contain CHPX
        /// </summary>
        public List<FormattedDiskPageCHPX> AllChpxFkps;

        /// <summary>
        /// A table that contains the positions of the headers and footer in the text.
        /// </summary>
        public HeaderAndFooterTable HeaderAndFooterTable;

        public StwStructure UserVariables;

        public WordDocument Glossary;

        public WordDocument(StructuredStorageReader reader, int fibFC = 0)
        {
            Parse(reader, fibFC);
        }

        void Parse(StructuredStorageReader reader, int fibFC)
        {
            this.Storage = reader;
            this.WordDocumentStream = reader.GetStream("WordDocument");

            //parse FIB
            this.WordDocumentStream.Seek(fibFC, System.IO.SeekOrigin.Begin);
            this.FIB = new FileInformationBlock(new VirtualStreamReader(this.WordDocumentStream));

            //check the file version
            if ((int)this.FIB.nFib != 0)
            {
                if (this.FIB.nFib < FileInformationBlock.FibVersion.Fib1997Beta)
                    throw new ByteParseException("Could not parse the file because it was created by an unsupported application (Word 95 or older).");
            }
            else
            {
                if (this.FIB.nFibNew < FileInformationBlock.FibVersion.Fib1997Beta)
                    throw new ByteParseException("Could not parse the file because it was created by an unsupported application (Word 95 or older).");
            }

            //get the streams
            this.TableStream = reader.GetStream(this.FIB.fWhichTblStm ? "1Table" : "0Table");

            try
            {
                this.DataStream = reader.GetStream("Data");
            }
            catch (StreamNotFoundException)
            {
                this.DataStream = null;
            }

            //Read all needed STTBs
            this.RevisionAuthorTable = new StringTable(typeof(string), this.TableStream, this.FIB.fcSttbfRMark, this.FIB.lcbSttbfRMark);
            this.FontTable = new StringTable(typeof(FontFamilyName), this.TableStream, this.FIB.fcSttbfFfn, this.FIB.lcbSttbfFfn);
            this.BookmarkNames = new StringTable(typeof(string), this.TableStream, this.FIB.fcSttbfBkmk, this.FIB.lcbSttbfBkmk);
            this.AutoTextNames = new StringTable(typeof(string), this.TableStream, this.FIB.fcSttbfGlsy, this.FIB.lcbSttbfGlsy);
            //this.ProtectionUsers = new StringTable(typeof(String), this.TableStream, this.FIB.fcSttbProtUser, this.FIB.lcbSttbProtUser);
            //
            this.UserVariables = new StwStructure(this.TableStream, this.FIB.fcStwUser, this.FIB.lcbStwUser);

            //Read all needed PLCFs
            this.AnnotationsReferencePlex = new Plex<AnnotationReferenceDescriptor>(30, this.TableStream, this.FIB.fcPlcfandRef, this.FIB.lcbPlcfandRef);
            this.TextboxBreakPlex = new Plex<BreakDescriptor>(6, this.TableStream, this.FIB.fcPlcfTxbxBkd, this.FIB.lcbPlcfTxbxBkd);
            this.TextboxBreakPlexHeader = new Plex<BreakDescriptor>(6, this.TableStream, this.FIB.fcPlcfTxbxHdrBkd, this.FIB.lcbPlcfTxbxHdrBkd);
            this.OfficeDrawingPlex = new Plex<FileShapeAddress>(26, this.TableStream, this.FIB.fcPlcSpaMom, this.FIB.lcbPlcSpaMom);
            this.OfficeDrawingPlexHeader = new Plex<FileShapeAddress>(26, this.TableStream, this.FIB.fcPlcSpaHdr, this.FIB.lcbPlcSpaHdr);
            this.SectionPlex = new Plex<SectionDescriptor>(12, this.TableStream, this.FIB.fcPlcfSed, this.FIB.lcbPlcfSed);
            this.BookmarkStartPlex = new Plex<BookmarkFirst>(4, this.TableStream, this.FIB.fcPlcfBkf, this.FIB.lcbPlcfBkf);
            this.EndnoteReferencePlex = new Plex<short>(2, this.TableStream, this.FIB.fcPlcfendRef, this.FIB.lcbPlcfendRef);
            this.FootnoteReferencePlex = new Plex<short>(2, this.TableStream, this.FIB.fcPlcffndRef, this.FIB.lcbPlcffndRef);
            // PLCFs without types
            this.BookmarkEndPlex = new Plex<Exception>(0, this.TableStream, this.FIB.fcPlcfBkl, this.FIB.lcbPlcfBkl);
            this.AutoTextPlex = new Plex<Exception>(0, this.TableStream, this.FIB.fcPlcfGlsy, this.FIB.lcbPlcfGlsy);

            //read the FKPs
            this.AllPapxFkps = FormattedDiskPagePAPX.GetAllPAPXFKPs(this.FIB, this.WordDocumentStream, this.TableStream, this.DataStream);
            this.AllChpxFkps = FormattedDiskPageCHPX.GetAllCHPXFKPs(this.FIB, this.WordDocumentStream, this.TableStream);

            //read custom tables
            this.DocumentProperties = new DocumentProperties(this.FIB, this.TableStream);
            this.Styles = new StyleSheet(this.FIB, this.TableStream, this.DataStream);
            this.ListTable = new ListTable(this.FIB, this.TableStream);
            this.ListFormatOverrideTable = new ListFormatOverrideTable(this.FIB, this.TableStream);
            this.OfficeArtContent = new OfficeArtContent(this.FIB, this.TableStream);
            this.HeaderAndFooterTable = new HeaderAndFooterTable(this);
            this.AnnotationReferenceExtraTable = new AnnotationReferenceExtraTable(this.FIB, this.TableStream);
            this.CommandTable = new CommandTable(this.FIB, this.TableStream);
            this.AnnotationOwners = new AnnotationOwnerList(this.FIB, this.TableStream);

            //parse the piece table and construct a list that contains all chars
            this.PieceTable = new PieceTable(this.FIB, this.TableStream);
            this.Text = this.PieceTable.GetAllChars(this.WordDocumentStream);

            //build a dictionaries of all PAPX
            this.AllPapx = new Dictionary<int, ParagraphPropertyExceptions>();
            for (int i = 0; i < this.AllPapxFkps.Count; i++)
            {
                for (int j = 0; j < this.AllPapxFkps[i].grppapx.Length; j++)
                {
                    this.AllPapx.Add(this.AllPapxFkps[i].rgfc[j], this.AllPapxFkps[i].grppapx[j]);
                }
            }

            //build a dictionary of all SEPX
            this.AllSepx = new Dictionary<int, SectionPropertyExceptions>();
            for (int i = 0; i < this.SectionPlex.Elements.Count; i++)
            {
                //Read the SED
                var sed = (SectionDescriptor)this.SectionPlex.Elements[i];
                int cp = this.SectionPlex.CharacterPositions[i + 1];

                //Get the SEPX
                var wordReader = new VirtualStreamReader(this.WordDocumentStream);
                this.WordDocumentStream.Seek(sed.fcSepx, System.IO.SeekOrigin.Begin);
                short cbSepx = wordReader.ReadInt16();
                var sepx = new SectionPropertyExceptions(wordReader.ReadBytes(cbSepx - 2));

                this.AllSepx.Add(cp, sepx);
            }

            //read the Glossary
            if (this.FIB.pnNext > 0)
            {
                this.Glossary = new WordDocument(this.Storage, (int)(this.FIB.pnNext * 512));
            }
        }

        /// <summary>
        /// Returns a list of all CHPX which are valid for the given FCs.
        /// </summary>
        /// <param name="fcMin">The lower boundary</param>
        /// <param name="fcMax">The upper boundary</param>
        /// <returns>The FCs</returns>
        public List<int> GetFileCharacterPositions(int fcMin, int fcMax)
        {
            var list = new List<int>();

            for (int i = 0; i < this.AllChpxFkps.Count; i++)
            {
                var fkp = this.AllChpxFkps[i];

                //if the last fc of this fkp is smaller the fcMin
                //this fkp is before the requested range
                if (fkp.rgfc[fkp.rgfc.Length - 1] < fcMin)
                {
                    continue;
                }

                //if the first fc of this fkp is larger the Max
                //this fkp is beyond the requested range
                if (fkp.rgfc[0] > fcMax)
                {
                    break;
                }

                //don't add the duplicated values of the FKP boundaries (Length-1)
                int max = fkp.rgfc.Length - 1;

                //last fkp? 
                //use full table
                if (i == (this.AllChpxFkps.Count - 1))
                {
                    max = fkp.rgfc.Length;
                }

                for (int j = 0; j < max; j++)
                {
                    if (fkp.rgfc[j] < fcMin && fkp.rgfc[j + 1] > fcMin)
                    {
                        //this chpx starts before fcMin
                        list.Add(fkp.rgfc[j]);
                    }
                    else if (fkp.rgfc[j] >= fcMin && fkp.rgfc[j] < fcMax)
                    {
                        //this chpx is in the range
                        list.Add(fkp.rgfc[j]);
                    }
                }
            }

            return list;
        }


        /// <summary>
        /// Returnes a list of all CharacterPropertyExceptions which correspond to text 
        /// between the given boundaries.
        /// </summary>
        /// <param name="fcMin">The lower boundary</param>
        /// <param name="fcMax">The upper boundary</param>
        /// <returns>The FCs</returns>
        public List<CharacterPropertyExceptions> GetCharacterPropertyExceptions(int fcMin, int fcMax)
        {
            var list = new List<CharacterPropertyExceptions>();

            foreach (var fkp in this.AllChpxFkps)
            {
                //get the CHPX
                for (int j = 0; j < fkp.grpchpx.Length; j++)
                {
                    if (fkp.rgfc[j] < fcMin && fkp.rgfc[j + 1] > fcMin)
                    {
                        //this chpx starts before fcMin
                        list.Add(fkp.grpchpx[j]);
                    }
                    else if (fkp.rgfc[j] >= fcMin && fkp.rgfc[j] < fcMax)
                    {
                        //this chpx is in the range
                        list.Add(fkp.grpchpx[j]);
                    }
                }
            }

            return list;
        }

        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<WordDocument>)mapping).Apply(this);
        }

        #endregion
    }
}
