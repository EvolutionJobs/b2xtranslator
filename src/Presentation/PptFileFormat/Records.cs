using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorageReader;
using System.Reflection;
using System.Collections;
using System.IO;
using DIaLOGIKa.b2xtranslator.Tools;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat.Records
{
    /// <summary>
    /// Used for mapping Office record TypeCodes to the classes implementing them.
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    public class OfficeRecordAttribute : Attribute
    {
        public UInt16 TypeCode;
    }

    public class Record : IEnumerable<Record>
    {
        public const uint HEADER_SIZE_IN_BYTES = (16 + 16 + 32) / 8;

        public uint TotalSize
        {
            get { return HeaderSize + BodySize; }
        }

        private Record _ParentRecord = null;

        public Record ParentRecord
        {
            get { return _ParentRecord; }
            set {
                if (_ParentRecord != null)
                    throw new Exception("Can only set ParentRecord once");

                _ParentRecord = value;
                this.AfterParentSet();
            }
        }

        public uint HeaderSize = HEADER_SIZE_IN_BYTES;
        public uint BodySize;

        public byte[] RawData;

        protected BinaryReader Reader;

        public uint TypeCode;
        public uint Version;
        public uint Instance;

        public Record(BinaryReader _reader, uint bodySize, uint typeCode, uint version, uint instance)
        {
            this.BodySize = bodySize;
            this.TypeCode = typeCode;
            this.Version = version;
            this.Instance = instance;

            this.RawData = _reader.ReadBytes((int)this.BodySize);

            this.Reader = new BinaryReader(new MemoryStream(this.RawData));
        }

        public virtual void AfterParentSet() { }

        public void DumpToStream(Stream output)
        {
            using (BinaryWriter writer = new BinaryWriter(output))
            {
                writer.Write(this.RawData, 0, this.RawData.Length);
            }
        }

        public string GetIdentifier()
        {
            StringBuilder result = new StringBuilder();

            Record r = this;
            bool isFirst = true;

            while (r != null)
            {
                if (!isFirst)
                    result.Insert(0, "-");

                result.Insert(0, String.Format("{0}i{1}", r.FormatType(), r.Instance));

                r = r.ParentRecord;
                isFirst = false;
            }

            return result.ToString();
        }

        public string FormatType()
        {
            bool isEscherRecord = (this.TypeCode >= 0xF000 && this.TypeCode <= 0xFFFF);
            return String.Format(isEscherRecord ? "0x{0:X}" : "{0}", this.TypeCode);
        }

        public virtual string ToString(uint depth)
        {
            return String.Format("{0}{2}:\n{1}Type = {3}, Version = {4}, Instance = {5}, BodySize = {6}",
                IndentationForDepth(depth), IndentationForDepth(depth + 1),
                this.GetType(), this.FormatType(), this.Version, this.Instance, this.BodySize
            );
        }

        public override string ToString()
        {
            return this.ToString(0);
        }

        public void VerifyReadToEnd()
        {
            long streamPos = this.Reader.BaseStream.Position;
            long streamLen = this.Reader.BaseStream.Length;

            if (streamPos != streamLen)
            {
                throw new Exception(String.Format(
                    "Record {3} didn't read to end: (stream position: {1} of {2})\n{0}",
                    this, streamPos, streamLen, this.GetIdentifier()));
            }
        }

        #region IEnumerable<Record> Members

        public virtual IEnumerator<Record> GetEnumerator()
        {
            yield return this;
        }

        #endregion

        #region IEnumerable Members

        IEnumerator IEnumerable.GetEnumerator()
        {
            foreach (Record record in this)
                yield return record;
        }

        #endregion

        #region Static attributes and methods

        internal static string IndentationForDepth(uint depth)
        {
            StringBuilder result = new StringBuilder();

            for (uint i = 0; i < depth; i++)
                result.Append("  ");

            return result.ToString();
        }

        public static Dictionary<UInt16, Type> TypeToRecordClassMapping = GetTypeToRecordClassMapping();

        private static Dictionary<UInt16, Type> GetTypeToRecordClassMapping()
        {
            Dictionary<UInt16, Type> result = new Dictionary<UInt16, Type>();

            // Note: We return a Dictionary that maps Office record TypeCodes to Office record classes.
            // We do this by querying all classes in the current assembly, filtering by namespace
            // PptFileFormat.Records and looking for attributes of type OfficeRecord.
            //
            // If in doubt see usage below.
            foreach (Type t in Assembly.GetExecutingAssembly().GetTypes())
            {
                if (t.Namespace == typeof(Record).Namespace)
                {
                    object[] attrs = t.GetCustomAttributes(typeof(OfficeRecordAttribute), false);

                    OfficeRecordAttribute attr = null;
                    
                    if (attrs.Length > 0)
                        attr = attrs[0] as OfficeRecordAttribute;

                    if (attr != null)
                    {
                        UInt16 typeCode = attr.TypeCode;

                        if (result.ContainsKey(typeCode))
                        {
                            throw new Exception(String.Format(
                                "Tried to register TypeCode {0} to {1}, but it is already registered to {2}",
                                typeCode, t, result[typeCode]));
                        }

                        result.Add(attr.TypeCode, t);
                    }
                }
            }

            return result;
        }

        public static Record readRecord(Stream stream)
        {
            return readRecord(new BinaryReader(stream));
        }

        public static Record readRecord(BinaryReader reader)
        {
            UInt16 verAndInstance = reader.ReadUInt16();
            uint version = verAndInstance & 0x000FU;         // first 4 bit of field verAndInstance
            uint instance = (verAndInstance & 0xFFF0U) >> 4; // last 12 bit of field verAndInstance

            UInt16 typeCode = reader.ReadUInt16();
            UInt32 size = reader.ReadUInt32();

            bool isContainer = (version == 0xF);

            Record result;
            Type cls;

            if (TypeToRecordClassMapping.TryGetValue(typeCode, out cls))
            {
                ConstructorInfo constructor = cls.GetConstructor(new Type[] {
                    typeof(BinaryReader), typeof(uint), typeof(uint), typeof(uint), typeof(uint) 
                });

                if (constructor == null)
                {
                    throw new Exception(String.Format(
                        "Internal error: Could not find a matching constructor for class {0}",
                        cls));
                }

                try
                {
                    result = (Record)constructor.Invoke(new object[] {
                        reader, size, typeCode, version, instance
                    });
                }
                catch (TargetInvocationException e)
                {
                    Console.WriteLine(e.InnerException);
                    throw e.InnerException;
                }
            }
            else
            {
                result = new UnknownRecord(reader, size, typeCode, version, instance);
            }

            return result;
        }

        #endregion
    }

    public class UnknownRecord : Record
    {
        public UnknownRecord(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.Reader.ReadBytes((int)size);
        }
    }

    /// <summary>
    /// Regular containers are containers with Record children.
    /// 
    /// (There also is containers that only have a zipped XML payload.
    /// </summary>
    public class RegularContainer : Record
    {
        public List<Record> Children = new List<Record>();

        public RegularContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            uint readSize = 0;

            while (readSize < this.BodySize)
            {
                Record child = Record.readRecord(this.Reader);

                this.Children.Add(child);
                child.ParentRecord = this;

                child.VerifyReadToEnd();

                readSize += child.TotalSize;
            }
        }

        override public string ToString(uint depth)
        {
            StringBuilder result = new StringBuilder(base.ToString(depth));

            depth++;

            if (this.Children.Count > 0)
            {
                result.AppendLine();
                result.Append(IndentationForDepth(depth));
                result.Append("Children:");
            }

            foreach (Record record in this.Children)
            {
                result.AppendLine();
                result.Append(record.ToString(depth + 1));
            }

            return result.ToString();
        }

        #region IEnumerable<Record> Members

        public override IEnumerator<Record> GetEnumerator()
        {
            yield return this;

            foreach (Record recordChild in this.Children)
                foreach (Record record in recordChild)
                    yield return record;
        }

        #endregion
    }

    #region PowerPoint records
    [OfficeRecordAttribute(TypeCode = 1000)]
    public class PptDocumentRecord : RegularContainer
    {
        public PptDocumentRecord(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

    [OfficeRecordAttribute(TypeCode = 1001)]
    public class DocumentAtom : Record
    {
        public GPointAtom SlideSize;
        public GPointAtom NotesSize;
        public GRatioAtom ServerZoom;

        public UInt32 NotesMasterPersist;
        public UInt32 HandoutMasterPersist;
        public UInt16 FirstSlideNum;
        public Int16 SlideSizeType;

        public bool SaveWithFonts;
        public bool OmitTitlePlace;
        public bool RightToLeft;
        public bool ShowComments;

        public DocumentAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.SlideSize = new GPointAtom(this.Reader);
            this.NotesSize = new GPointAtom(this.Reader);
            this.ServerZoom = new GRatioAtom(this.Reader);

            this.NotesMasterPersist = this.Reader.ReadUInt32();
            this.HandoutMasterPersist = this.Reader.ReadUInt32();
            this.FirstSlideNum = this.Reader.ReadUInt16();
            this.SlideSizeType = this.Reader.ReadInt16();

            this.SaveWithFonts = this.Reader.ReadByte() != 0;
            this.OmitTitlePlace = this.Reader.ReadByte() != 0;
            this.RightToLeft = this.Reader.ReadByte() != 0;
            this.ShowComments = this.Reader.ReadByte() != 0;
        }

        override public string ToString(uint depth)
        {
            return String.Format("{0}\n{1}SlideSize = {2}, NotesSize = {3}, ServerZoom = {4}\n{1}" +
                "NotesMasterPersist = {5}, HandoutMasterPersist = {6}, FirstSlideNum = {7}, SlideSizeType = {8}\n{1}" +
                "SaveWithFonts = {9}, OmitTitlePlace = {10}, RightToLeft = {11}, ShowComments = {12}",

                base.ToString(depth), IndentationForDepth(depth + 1),

                this.SlideSize, this.NotesSize, this.ServerZoom,

                this.NotesMasterPersist, this.HandoutMasterPersist, this.FirstSlideNum, this.SlideSizeType,

                this.SaveWithFonts, this.OmitTitlePlace, this.RightToLeft, this.ShowComments);
        }
    }

    public class GPointAtom
    {
        public Int32 X;
        public Int32 Y;

        public GPointAtom(BinaryReader reader)
        {
            this.X = reader.ReadInt32();
            this.Y = reader.ReadInt32();
        }

        override public string ToString()
        {
            return String.Format("PointAtom({0}, {1})", this.X, this.Y);
        }
    }

    public class GRatioAtom
    {
        public Int32 Numer;
        public Int32 Denom;

        public GRatioAtom(BinaryReader reader)
        {
            this.Numer = reader.ReadInt32();
            this.Denom = reader.ReadInt32();
        }

        override public string ToString()
        {
            return String.Format("RatioAtom({0}, {1})", this.Numer, this.Denom);
        }
    }

    [OfficeRecordAttribute(TypeCode = 1006)]
    public class Slide : RegularContainer
    {
        public Slide(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

    [OfficeRecordAttribute(TypeCode = 1007)]
    public class SlideAtom : Record
    {
        public SSlideLayoutAtom Layout;
        public Int32 MasterId;
        public Int32 NotesId;
        public UInt16 Flags;

        public SlideAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.Layout = new SSlideLayoutAtom(this.Reader);
            this.MasterId = this.Reader.ReadInt32();
            this.NotesId = this.Reader.ReadInt32();
            this.Flags = this.Reader.ReadUInt16();
            this.Reader.ReadUInt16(); // Throw away undocumented data
        }

        override public string ToString(uint depth)
        {
            return String.Format("{0}\n{1}Layout = {2}\n{1}MasterId = {3}, NotesId = {4}, Flags = {5})",
                base.ToString(depth), IndentationForDepth(depth + 1),
                this.Layout, this.MasterId, this.NotesId, this.Flags);
        }
    }

    public class SSlideLayoutAtom
    {
        public Int32 Geom;
        public byte[] PlaceholderIds = new byte[8];

        public SSlideLayoutAtom(BinaryReader reader)
        {
            this.Geom = reader.ReadInt32();
 
            for (int i = 0; i < 8; i++)
                this.PlaceholderIds[i] = reader.ReadByte();
        }

        public override string ToString()
        {
            string s = String.Join(", ",
                Array.ConvertAll<byte, string>(this.PlaceholderIds,
                delegate(byte b) { return b.ToString(); }));

            return String.Format("SSlideLayoutAtom(Geom = {0}, PlaceholderIds = [{1}])",
                this.Geom, s);
        }
    }

    [OfficeRecordAttribute(TypeCode = 1016)]
    public class List : RegularContainer
    {
        public List(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

    [OfficeRecordAttribute(TypeCode = 1035)]
    public class PPDrawingGroup : RegularContainer
    {
        public PPDrawingGroup(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

    [OfficeRecordAttribute(TypeCode = 1036)]
    public class PPDrawing : RegularContainer
    {
        public PPDrawing(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

    [OfficeRecordAttribute(TypeCode = 3999)]
    public class TextHeaderAtom : Record
    {
        public UInt32 TextType;

        public TextHeaderAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.TextType = this.Reader.ReadUInt32();
        }
    }

    [OfficeRecordAttribute(TypeCode = 4001)]
    public class StyleTextPropAtom : Record
    {
        [FlagsAttribute]
        public enum ParagraphMask : uint
        {
            None                        = 0,
            HasCustomBullet             = 1 << 0,
            HasCustomBulletTypeface     = 1 << 1,
            HasCustomBulletColor        = 1 << 2,
            HasCustomBulletSize         = 1 << 3,

            BulletFlagsFieldPresent     = HasCustomBullet | HasCustomBulletTypeface |
                                          HasCustomBulletColor | HasCustomBulletSize,

            BulletTypefacePresent       = 1 << 4,
            BulletSizePresent           = 1 << 5,
            BulletColorPresent          = 1 << 6,
            BulletCharPresent           = 1 << 7,
            LeftMarginPresent           = 1 << 8,

            // Bit 9 is unused

            IndentPresent               = 1 << 10,
            AlignmentPresent            = 1 << 11,
            LineSpacingPresent          = 1 << 12,
            SpaceBeforePresent          = 1 << 13,
            SpaceAfterPresent           = 1 << 14,
            DefaultTabSizePresent       = 1 << 15,
            BaseLinePresent             = 1 << 16,

            HasCustomCharWrap           = 1 << 17,
            HasCustomWordWrap           = 1 << 18,
            HasCustomOverflow           = 1 << 19,

            LineBreakFlagsFieldPresent  = HasCustomCharWrap | HasCustomWordWrap | HasCustomOverflow,

            TabStopsPresent             = 1 << 20,
            TextDirectionPresent        = 1 << 21
        }

        public class ParagraphRun
        {
            public UInt32 Length;
            public UInt16 IndentLevel;
            public ParagraphMask Mask;

            #region Presence flag getters
            public bool BulletFlagsFieldPresent
            {
                get { return (this.Mask & ParagraphMask.BulletFlagsFieldPresent) != 0; }
            }

            public bool BulletCharPresent
            {
                get { return (this.Mask & ParagraphMask.BulletCharPresent) != 0; }
            }

            public bool BulletTypefacePresent
            {
                get { return (this.Mask & ParagraphMask.BulletTypefacePresent) != 0; }
            }

            public bool BulletSizePresent
            {
                get { return (this.Mask & ParagraphMask.BulletSizePresent) != 0; }
            }

            public bool BulletColorPresent
            {
                get { return (this.Mask & ParagraphMask.BulletColorPresent) != 0; }
            }

            public bool AlignmentPresent
            {
                get { return (this.Mask & ParagraphMask.AlignmentPresent) != 0; }
            }

            public bool LineSpacingPresent
            {
                get { return (this.Mask & ParagraphMask.LineSpacingPresent) != 0; }
            }

            public bool SpaceBeforePresent
            {
                get { return (this.Mask & ParagraphMask.SpaceBeforePresent) != 0; }
            }

            public bool SpaceAfterPresent
            {
                get { return (this.Mask & ParagraphMask.SpaceAfterPresent) != 0; }
            }

            public bool LeftMarginPresent
            {
                get { return (this.Mask & ParagraphMask.LeftMarginPresent) != 0; }
            }

            public bool IndentPresent
            {
                get { return (this.Mask & ParagraphMask.IndentPresent) != 0; }
            }

            public bool DefaultTabSizePresent
            {
                get { return (this.Mask & ParagraphMask.DefaultTabSizePresent) != 0; }
            }

            public bool TabStopsPresent
            {
                get { return (this.Mask & ParagraphMask.TabStopsPresent) != 0; }
            }

            public bool BaseLinePresent
            {
                get { return (this.Mask & ParagraphMask.BaseLinePresent) != 0; }
            }

            public bool LineBreakFlagsFieldPresent
            {
                get { return (this.Mask & ParagraphMask.LineBreakFlagsFieldPresent) != 0; }
            }

            public bool TextDirectionPresent
            {
                get { return (this.Mask & ParagraphMask.TextDirectionPresent) != 0; }
            }
            #endregion

            public UInt16? BulletFlags;
            public char? BulletChar;
            public UInt16? BulletTypefaceIdx;
            public Int16? BulletSize;
            public GrColorAtom BulletColor;
            public Int16? Alignment;
            public Int16? LineSpacing;
            public Int16? SpaceBefore;
            public Int16? SpaceAfter;
            public Int16? LeftMargin;
            public Int16? Indent;
            public Int16? DefaultTabSize;
            public UInt16? BaseLine;
            public UInt16? LineBreakFlags;
            public UInt16? TextDirection;

            public ParagraphRun(BinaryReader reader)
            {
                this.Length = reader.ReadUInt32();
                this.IndentLevel = reader.ReadUInt16();
                this.Mask = (ParagraphMask)reader.ReadUInt32();

                // Note: These appear in Mask as well -- there they are true
                // when the flag differs from the Master style.
                // The actual value for the differing flags is stored here.
                // (TODO: This is still a guess. Verify.)
                if (this.BulletFlagsFieldPresent)
                    this.BulletFlags = reader.ReadUInt16();

                if (this.BulletCharPresent)
                    this.BulletChar = (char)reader.ReadUInt16();

                if (this.BulletTypefacePresent)
                    this.BulletTypefaceIdx = reader.ReadUInt16();

                if (this.BulletSizePresent)
                    this.BulletSize = reader.ReadInt16();

                if (this.BulletColorPresent)
                    this.BulletColor = new GrColorAtom(reader);

                if (this.AlignmentPresent)
                    this.Alignment = reader.ReadInt16();

                if (this.LineSpacingPresent)
                    this.LineSpacing = reader.ReadInt16();

                if (this.SpaceBeforePresent)
                    this.SpaceBefore = reader.ReadInt16();

                if (this.SpaceAfterPresent)
                    this.SpaceAfter = reader.ReadInt16();

                if (this.LeftMarginPresent)
                    this.LeftMargin = reader.ReadInt16();

                if (this.IndentPresent)
                    this.Indent = reader.ReadInt16();

                if (this.DefaultTabSizePresent)
                    this.DefaultTabSize = reader.ReadInt16();

                if (this.TabStopsPresent)
                {
                    UInt16 tabStopsCount = reader.ReadUInt16();
                    if (tabStopsCount != 0)
                        throw new NotImplementedException("Tab stop reading not yet implemented"); // TODO
                }

                if (this.BaseLinePresent)
                    this.BaseLine = reader.ReadUInt16();

                if (this.LineBreakFlagsFieldPresent)
                    this.LineBreakFlags = reader.ReadUInt16();

                if (this.TextDirectionPresent)
                    this.TextDirection = reader.ReadUInt16();
            }

            public string ToString(uint depth)
            {
                StringBuilder result = new StringBuilder();

                result.Append(IndentationForDepth(depth));
                result.Append(base.ToString());

                depth++;
                string indent = IndentationForDepth(depth);

                result.AppendFormat("\n{0}Length = {1}", indent, this.Length);
                result.AppendFormat("\n{0}IndentLevel = {1}", indent, this.IndentLevel);
                result.AppendFormat("\n{0}Mask = {1}", indent, this.Mask);

                if (this.BulletFlagsFieldPresent)
                    result.AppendFormat("\n{0}BulletFlags = {1}", indent, this.BulletFlags);

                if (this.BulletCharPresent)
                    result.AppendFormat("\n{0}BulletChar = {1}", indent, this.BulletChar);

                if (this.BulletTypefacePresent)
                    result.AppendFormat("\n{0}BulletTypefaceIdx = {1}", indent, this.BulletTypefaceIdx);

                if (this.BulletSizePresent)
                    result.AppendFormat("\n{0}BulletSize = {1}", indent, this.BulletSize);

                if (this.BulletColorPresent)
                    result.AppendFormat("\n{0}BulletColor = {1}", indent, this.BulletColor);

                if (this.AlignmentPresent)
                    result.AppendFormat("\n{0}Alignment = {1}", indent, this.Alignment);

                if (this.LineSpacingPresent)
                    result.AppendFormat("\n{0}LineSpacing = {1}", indent, this.LineSpacing);

                if (this.SpaceBeforePresent)
                    result.AppendFormat("\n{0}SpaceBefore = {1}", indent, this.SpaceBefore);

                if (this.SpaceAfterPresent)
                    result.AppendFormat("\n{0}SpaceAfter = {1}", indent, this.SpaceAfter);

                if (this.LeftMarginPresent)
                    result.AppendFormat("\n{0}LeftMargin = {1}", indent, this.LeftMargin);

                if (this.IndentPresent)
                    result.AppendFormat("\n{0}Indent = {1}", indent, this.Indent);

                if (this.DefaultTabSizePresent)
                    result.AppendFormat("\n{0}DefaultTabSize = {1}", indent, this.DefaultTabSize);

                if (this.BaseLinePresent)
                    result.AppendFormat("\n{0}BaseLine = {1}", indent, this.BaseLine);

                if (this.LineBreakFlagsFieldPresent)
                    result.AppendFormat("\n{0}LineBreakFlags = {1}", indent, this.LineBreakFlags);

                if (this.TextDirectionPresent)
                    result.AppendFormat("\n{0}TextDirection = {1}", indent, this.TextDirection);

                return result.ToString();
            }

            public override string ToString()
            {
                return this.ToString(0);
            }
        }

        [FlagsAttribute]
        public enum CharacterMask : uint
        {
            None = 0,

            // Bit 0 - 15 are used for marking style flag presence
            StyleFlagsFieldPresent = 0xFFFF,

            TypefacePresent = 1 << 16,
            SizePresent = 1 << 17,
            ColorPresent = 1 << 18,
            PositionPresent = 1 << 19,

            // Bit 20 is unused

            FEOldTypefacePresent = 1 << 21,
            ANSITypefacePresent = 1 << 22,
            SymbolTypefacePresent = 1 << 23
        }

        [FlagsAttribute]
        public enum StyleMask : uint
        {
            None = 0,

            IsBold = 1 << 0,
            IsItalic = 1 << 1,
            IsUnderlined = 1 << 2,

            // Bit 3 is unused

            HasShadow = 1 << 4,
            HasAsianSmartQuotes = 1 << 5,

            // Bit 6 is unused

            HasHorizonNumRendering = 1 << 7,

            // Bit 8 is unused

            IsEmbossed = 1 << 9,

            ExtensionNibble = 0xF << 10 // Bit 10 - 13

            // Bit 14 - 15 are unused
        }

        public class CharacterRun
        {
            public UInt32 Length;
            public CharacterMask Mask;

            #region Presence flag getters
            public bool StyleFlagsFieldPresent
            {
                get { return (this.Mask & CharacterMask.StyleFlagsFieldPresent) != 0; }
            }
            
            public bool TypefacePresent
            {
                get { return (this.Mask & CharacterMask.TypefacePresent) != 0; }
            }

            public bool FEOldTypefacePresent
            {
                get { return (this.Mask & CharacterMask.FEOldTypefacePresent) != 0; }
            }

            public bool ANSITypefacePresent
            {
                get { return (this.Mask & CharacterMask.ANSITypefacePresent) != 0; }
            }

            public bool SymbolTypefacePresent
            {
                get { return (this.Mask & CharacterMask.SymbolTypefacePresent) != 0; }
            }

            public bool SizePresent
            {
                get { return (this.Mask & CharacterMask.SizePresent) != 0; }
            }

            public bool PositionPresent
            {
                get { return (this.Mask & CharacterMask.PositionPresent) != 0; }
            }

            public bool ColorPresent
            {
                get { return (this.Mask & CharacterMask.ColorPresent) != 0; }
            }
            #endregion

            public StyleMask? Style;
            public UInt16? TypefaceIdx;
            public UInt16? FEOldTypefaceIdx;
            public UInt16? ANSITypefaceIdx;
            public UInt16? SymbolTypefaceIdx;
            public UInt16? Size;
            public UInt16? Position;
            public GrColorAtom Color;

            public CharacterRun(BinaryReader reader)
            {
                this.Length = reader.ReadUInt32();
                this.Mask = (CharacterMask)reader.ReadUInt32();

                if (this.StyleFlagsFieldPresent)
                    this.Style = (StyleMask)reader.ReadUInt16();

                if (this.TypefacePresent)
                    this.TypefaceIdx = reader.ReadUInt16();

                if (this.FEOldTypefacePresent)
                    this.FEOldTypefaceIdx = reader.ReadUInt16();

                if (this.ANSITypefacePresent)
                    this.ANSITypefaceIdx = reader.ReadUInt16();

                if (this.SymbolTypefacePresent)
                    this.SymbolTypefaceIdx = reader.ReadUInt16();

                if (this.SizePresent)
                    this.Size = reader.ReadUInt16();

                if (this.PositionPresent)
                    this.Position = reader.ReadUInt16();

                if (this.ColorPresent)
                    this.Color = new GrColorAtom(reader);
            }

            public string ToString(uint depth)
            {
                StringBuilder result = new StringBuilder();

                result.Append(IndentationForDepth(depth));
                result.Append(base.ToString());

                depth++;
                string indent = IndentationForDepth(depth);

                result.AppendFormat("\n{0}Length = {1}", indent, this.Length);
                result.AppendFormat("\n{0}Mask = {1}", indent, this.Mask);

                if (this.StyleFlagsFieldPresent)
                    result.AppendFormat("\n{0}Style = {1}", indent, this.Style);

                if (this.TypefacePresent)
                    result.AppendFormat("\n{0}TypefaceIdx = {1}", indent, this.TypefaceIdx);

                if (this.FEOldTypefacePresent)
                    result.AppendFormat("\n{0}FEOldTypefaceIdx = {1}", indent, this.FEOldTypefaceIdx);

                if (this.ANSITypefacePresent)
                    result.AppendFormat("\n{0}ANSITypefaceIdx = {1}", indent, this.ANSITypefaceIdx);

                if (this.SymbolTypefacePresent)
                    result.AppendFormat("\n{0}SymbolTypefaceIdx = {1}", indent, this.SymbolTypefaceIdx);

                if (this.SizePresent)
                    result.AppendFormat("\n{0}Size = {1}", indent, this.Size);

                if (this.PositionPresent)
                    result.AppendFormat("\n{0}Position = {1}", indent, this.Position);

                if (this.ColorPresent)
                    result.AppendFormat("\n{0}Color = {1}", indent, this.Color);

                return result.ToString();
            }

            public override string ToString()
            {
                return this.ToString(0);
            }
        }

        public StyleTextPropAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
        }

        override public void AfterParentSet()
        {
            // Anmerkung: In OOXML kann ein Character-Properties-Element sich
            // nicht über mehrere Paragraphen hinweg erstrecken.
            // TODO: War das im Binärformat der Fall?

            // TODO: FindParentByType? FindChildByType? FindSiblingByType?

            ClientTextbox textbox = this.ParentRecord as ClientTextbox;
            if (textbox == null)
                throw new Exception("Record of type StyleTextPropAtom doesn't have parent of type ClientTextbox");

            TextAtom textAtom = textbox.Children.Find(
                delegate(Record sibling) { return sibling is TextAtom; }
            ) as TextAtom;

            /* This can legitimately happen... */
            if (textAtom == null)
            {
                this.Reader.Read(new char[this.BodySize], 0, (int)this.BodySize);
                return;
            }

            // TODO: Length in bytes? UTF-16 characters? Full width unicode characters?

            uint seenLength = 0;
            while (seenLength < textAtom.Text.Length)
            {
                ParagraphRun run = new ParagraphRun(this.Reader);
                Console.WriteLine(run.ToString());
                Console.WriteLine("  Text = {0}", Utils.StringInspect(
                    textAtom.Text.Substring((int)seenLength, (int)run.Length)));
                Console.WriteLine();

                seenLength += run.Length;
            }

            seenLength = 0;
            while (seenLength < textAtom.Text.Length)
            {
                CharacterRun run = new CharacterRun(this.Reader);
                Console.WriteLine(run.ToString());
                Console.WriteLine("  Text = {0}", Utils.StringInspect(
                    textAtom.Text.Substring((int)seenLength, (int)run.Length)));
                Console.WriteLine();

                seenLength += run.Length;
            }
        }

/*        public override string ToString(uint depth)
        {
            return String.Format("{0}\n{1}RunLength = {2}",
                base.ToString(depth), IndentationForDepth(depth + 1), this.RunLength);
        }*/
    }

    public class GrColorAtom
    {
        public byte Red;
        public byte Green;
        public byte Blue;
        public byte Index;

        public bool IsSchemeColor
        {
            get { return this.Index != 0xFE; }
        }

        public GrColorAtom(BinaryReader reader)
        {
            this.Red = reader.ReadByte();
            this.Green = reader.ReadByte();
            this.Blue = reader.ReadByte();
            this.Index = reader.ReadByte();
        }

        public override string ToString()
        {
            return String.Format("GrColorAtom({0}, {1}, {2}): Index = {3}",
                this.Red, this.Green, this.Blue, this.Index);
        }
    }

    public class TextAtom : Record
    {
        public string Text;

        public TextAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance, Encoding encoding)
            : base(_reader, size, typeCode, version, instance)
        {
            byte[] bytes = new byte[size];
            this.Reader.Read(bytes, 0, (int)size);

            this.Text = new String(encoding.GetChars(bytes)) + "\n";
        }

        public override string ToString(uint depth)
        {
            return String.Format("{0}\n{1}Text = {2}",
                base.ToString(depth), IndentationForDepth(depth + 1), this.Text);
        }
    }

    [OfficeRecordAttribute(TypeCode = 4000)]
    public class TextCharsAtom : TextAtom
    {
        public static Encoding ENCODING = Encoding.Unicode;

        public TextCharsAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance, ENCODING) { }
    }

    [OfficeRecordAttribute(TypeCode = 4008)]
    public class TextBytesAtom : TextAtom
    {
        public static Encoding ENCODING = Encoding.GetEncoding("iso-8859-1");

        public TextBytesAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance, ENCODING) { }
    }

    [OfficeRecordAttribute(TypeCode = 4080)]
    public class SlideListWithText : RegularContainer
    {
        public SlideListWithText(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

    #endregion

    #region Drawing records

    [OfficeRecordAttribute(TypeCode = 0xF000)]
    public class DrawingGroup : RegularContainer
    {
        public DrawingGroup(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

    [OfficeRecordAttribute(TypeCode = 0xF002)]
    public class DrawingContainer : RegularContainer
    {
        public DrawingContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

    [OfficeRecordAttribute(TypeCode = 0xF003)]
    public class GroupContainer : RegularContainer
    {
        public GroupContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

    [OfficeRecordAttribute(TypeCode = 0xF004)]
    public class ShapeContainer : RegularContainer
    {
        public ShapeContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

    [OfficeRecordAttribute(TypeCode = 0xF006)]
    public class DrawingGroupRecord : Record
    {
        public class FileIdCluster
        {
            public UInt32 DrawingGroupId;
            public UInt32 CSpIdCur;

            public FileIdCluster(BinaryReader reader)
            {
                this.DrawingGroupId = reader.ReadUInt32();
                this.CSpIdCur = reader.ReadUInt32();
            }

            public string ToString(uint depth)
            {
                StringBuilder result = new StringBuilder();

                result.Append(IndentationForDepth(depth));
                result.AppendFormat("FileIdCluster: DrawingGroupId = {0}, CSpIdCur = {1}",
                   this.DrawingGroupId, this.CSpIdCur);

                return result.ToString();
            }
        }

        public UInt32 MaxShapeId;           // Maximum shape ID
        public UInt32 IdClustersCount;      // Number of FileIdClusters
        public UInt32 ShapesSavedCount;     // Total number of shapes saved
        public UInt32 DrawingsSavedCount;   // Total number of drawings saved

        public List<FileIdCluster> Clusters = new List<FileIdCluster>();

        public DrawingGroupRecord(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.MaxShapeId = this.Reader.ReadUInt32();
            this.IdClustersCount = this.Reader.ReadUInt32() - 1; // Office saves the actual value + 1 -- flgr
            this.ShapesSavedCount = this.Reader.ReadUInt32();
            this.DrawingsSavedCount = this.Reader.ReadUInt32();

            for (int i = 0; i < this.IdClustersCount; i++)
            {
                Clusters.Add(new FileIdCluster(this.Reader));
            }
        }

        override public string ToString(uint depth)
        {
            StringBuilder result = new StringBuilder();

            result.AppendLine(base.ToString(depth));

            result.Append(IndentationForDepth(depth + 1));
            result.AppendFormat("MaxShapeId = {0}, IdClustersCount = {1}",
                this.MaxShapeId, this.IdClustersCount);
                
            result.AppendLine();
            result.Append(IndentationForDepth(depth + 1));
            result.AppendFormat("ShapesSavedCount = {0}, DrawingsSavedCount = {1}",
                this.ShapesSavedCount, this.DrawingsSavedCount);

            depth++;

            if (this.Clusters.Count > 0)
            {
                result.AppendLine();
                result.Append(IndentationForDepth(depth));
                result.Append("Clusters:");
            }

            foreach (FileIdCluster cluster in this.Clusters)
            {
                result.AppendLine();
                result.Append(cluster.ToString(depth + 1));
            }

            return result.ToString();
        }
    }

    [OfficeRecordAttribute(TypeCode = 0xF00D)]
    public class ClientTextbox : RegularContainer
    {
        public ClientTextbox(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

    [OfficeRecordAttribute(TypeCode = 0xF011)]
    public class ClientData : RegularContainer
    {
        public ClientData(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

    #endregion
}
