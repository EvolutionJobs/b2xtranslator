using System;
using System.Text;
using System.IO;
using b2xtranslator.OfficeDrawing;

namespace b2xtranslator.PptFileFormat
{
    public class ParagraphRun
    {
        public class TabStop
        {
            public ushort Distance;
            public ushort Alignment;

            public TabStop(BinaryReader reader)
            {
                this.Distance = reader.ReadUInt16();
                this.Alignment = reader.ReadUInt16();
            }
        }

        public uint Length;
        public ushort IndentLevel;
        public ParagraphMask Mask;

        public TabStop[] TabStops;

        #region Presence flag getters
        public bool BulletFlagsFieldPresent
        {
            get { return (this.Mask & ParagraphMask.BulletFlagsFieldExists) != 0; }
        }

        public bool BulletCharPresent
        {
            get { return (this.Mask & ParagraphMask.BulletChar) != 0; }
        }

        public bool BulletFontPresent
        {
            get { return (this.Mask & ParagraphMask.BulletFont) != 0; }
        }

        public bool BulletHasFont
        {
            get { return (this.Mask & ParagraphMask.BulletHasFont) != 0; }
        }

        public bool BulletSizePresent
        {
            get { return (this.Mask & ParagraphMask.BulletSize) != 0; }
        }

        public bool BulletHasSize
        {
            get { return (this.Mask & ParagraphMask.BulletHasSize) != 0; }
        }

        public bool BulletColorPresent
        {
            get {
                    return (this.Mask & ParagraphMask.BulletColor) != 0;
            }
        }

        public bool BulletHasColor
        {
            get { return (this.Mask & ParagraphMask.BulletHasColor) != 0; }
        }

        public bool AlignmentPresent
        {
            get { return (this.Mask & ParagraphMask.Align) != 0; }
        }

        public bool LineSpacingPresent
        {
            get { return (this.Mask & ParagraphMask.LineSpacing) != 0; }
        }

        public bool SpaceBeforePresent
        {
            get { return (this.Mask & ParagraphMask.SpaceBefore) != 0; }
        }

        public bool SpaceAfterPresent
        {
            get { return (this.Mask & ParagraphMask.SpaceAfter) != 0; }
        }

        public bool LeftMarginPresent
        {
            get { return (this.Mask & ParagraphMask.LeftMargin) != 0; }
        }

        public bool IndentPresent
        {
            get { return (this.Mask & ParagraphMask.Indent) != 0; }
        }

        public bool DefaultTabSizePresent
        {
            get { return (this.Mask & ParagraphMask.DefaultTabSize) != 0; }
        }

        public bool TabStopsPresent
        {
            get { return (this.Mask & ParagraphMask.TabStops) != 0; }
        }

        public bool FontAlignPresent
        {
            get { return (this.Mask & ParagraphMask.FontAlign) != 0; }
        }

        public bool LineBreakFlagsFieldPresent
        {
            get { return (this.Mask & ParagraphMask.WrapFlagsFieldExists) != 0; }
        }

        public bool TextDirectionPresent
        {
            get { return (this.Mask & ParagraphMask.TextDirection) != 0; }
        }
        #endregion

        public ushort? BulletFlags;
        public char? BulletChar;
        public ushort? BulletTypefaceIdx;
        public short? BulletSize;
        public GrColorAtom BulletColor;
        public short? Alignment;
        public short? LineSpacing;
        public short? SpaceBefore;
        public short? SpaceAfter;
        public short? LeftMargin;
        public short? Indent;
        public short? DefaultTabSize;
        public ushort? FontAlign;
        public ushort? LineBreakFlags;
        public ushort? TextDirection;

        public ParagraphRun(BinaryReader reader, bool noIndentField)
        {
            try
            {
            
                this.IndentLevel = noIndentField ? (ushort)0 : reader.ReadUInt16();
                this.Mask = (ParagraphMask)reader.ReadUInt32();

                // Note: These appear in Mask as well -- there they are true
                // when the flag differs from the Master style.
                // The actual value for the differing flags is stored here.
                // (TODO: This is still a guess. Verify.)
                if (this.BulletFlagsFieldPresent)
                    this.BulletFlags = reader.ReadUInt16();

                if (this.BulletCharPresent)
                    this.BulletChar = (char)reader.ReadUInt16();

                if (this.BulletFontPresent)
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
                    ushort tabStopsCount = reader.ReadUInt16();
                    this.TabStops = new TabStop[tabStopsCount];

                    for (int i = 0; i < tabStopsCount; i++)
                    {
                        this.TabStops[i] = new TabStop(reader);
                    }
                }

                if (this.FontAlignPresent)
                    this.FontAlign = reader.ReadUInt16();

                if (this.LineBreakFlagsFieldPresent)
                    this.LineBreakFlags = reader.ReadUInt16();

                if (this.TextDirectionPresent)
                    this.TextDirection = reader.ReadUInt16();

            }
            catch (Exception e)
            {
                string s = e.ToString();
            }
            //if (this.TabStopsPresent)
            //{
            //    UInt16 tabStopsCount = reader.ReadUInt16();
            //    this.TabStops = new TabStop[tabStopsCount];

            //    for (int i = 0; i < tabStopsCount; i++)
            //    {
            //        this.TabStops[i] = new TabStop(reader);
            //    }
            //}
        }

        public string ToString(uint depth)
        {
            var result = new StringBuilder();

            string indent = Record.IndentationForDepth(depth);

            result.Append(indent);
            result.Append(base.ToString());

            depth++;
            indent = Record.IndentationForDepth(depth);

            result.AppendFormat("\n{0}Length = {1}", indent, this.Length);
            result.AppendFormat("\n{0}IndentLevel = {1}", indent, this.IndentLevel);
            result.AppendFormat("\n{0}Mask = {1}", indent, this.Mask);

            if (this.BulletFlags != null)
                result.AppendFormat("\n{0}BulletFlags = {1}", indent, this.BulletFlags);

            if (this.BulletChar != null)
                result.AppendFormat("\n{0}BulletChar = {1}", indent, this.BulletChar);

            if (this.BulletTypefaceIdx != null)
                result.AppendFormat("\n{0}BulletTypefaceIdx = {1}", indent, this.BulletTypefaceIdx);

            if (this.BulletSize != null)
                result.AppendFormat("\n{0}BulletSize = {1}", indent, this.BulletSize);

            if (this.BulletColor != null)
                result.AppendFormat("\n{0}BulletColor = {1}", indent, this.BulletColor);

            if (this.Alignment != null)
                result.AppendFormat("\n{0}Alignment = {1}", indent, this.Alignment);

            if (this.LineSpacing != null)
                result.AppendFormat("\n{0}LineSpacing = {1}", indent, this.LineSpacing);

            if (this.SpaceBefore != null)
                result.AppendFormat("\n{0}SpaceBefore = {1}", indent, this.SpaceBefore);

            if (this.SpaceAfter != null)
                result.AppendFormat("\n{0}SpaceAfter = {1}", indent, this.SpaceAfter);

            if (this.LeftMargin != null)
                result.AppendFormat("\n{0}LeftMargin = {1}", indent, this.LeftMargin);

            if (this.Indent != null)
                result.AppendFormat("\n{0}Indent = {1}", indent, this.Indent);

            if (this.DefaultTabSize != null)
                result.AppendFormat("\n{0}DefaultTabSize = {1}", indent, this.DefaultTabSize);

            if (this.FontAlign != null)
                result.AppendFormat("\n{0}FontAlign = {1}", indent, this.FontAlign);

            if (this.LineBreakFlags != null)
                result.AppendFormat("\n{0}LineBreakFlags = {1}", indent, this.LineBreakFlags);

            if (this.TextDirection != null)
                result.AppendFormat("\n{0}TextDirection = {1}", indent, this.TextDirection);

            return result.ToString();
        }

        public override string ToString()
        {
            return this.ToString(0);
        }
    }

    [FlagsAttribute]
    public enum ParagraphMask : uint
    {
        None = 0,

        /// <summary>
        /// A bit that specifies whether the bulletFlags field of the TextPFException structure that
        /// contains this PFMasks exists and whether bulletFlags.fHasBullet is valid.
        /// </summary>
        HasBullet = 1 << 0,

        /// <summary>
        /// A bit that specifies whether the bulletFlags field of the TextPFException structure that
        /// contains this PFMasks exists and whether bulletFlags.fBulletHasFont is valid.
        /// </summary>
        BulletHasFont = 1 << 1,

        /// <summary>
        /// A bit that specifies whether the bulletFlags field of the TextPFException structure that
        /// contains this PFMasks exists and whether bulletFlags. fBulletHasColor is valid.
        /// </summary>
        BulletHasColor = 1 << 2,

        /// <summary>
        /// A bit that specifies whether the bulletFlags field of the TextPFException structure that
        /// contains this PFMasks exists and whether bulletFlags.fBulletHasSize is valid.
        /// </summary>
        BulletHasSize = 1 << 3,

        BulletFlagsFieldExists = HasBullet | BulletHasFont | BulletHasColor | BulletHasSize,

        /// <summary>
        /// A bit that specifies whether the bulletFontRef field of the TextPFException structure that
        /// contains this PFMasks exists.
        /// </summary>
        BulletFont = 1 << 4,

        /// <summary>
        /// A bit that specifies whether the bulletColor field of the TextPFException structure that
        /// contains this PFMasks exists.
        /// </summary>
        BulletColor = 1 << 5,

        /// <summary>
        /// A bit that specifies whether the bulletSize field of the TextPFException structure that
        /// contains this PFMasks exists.
        /// </summary>
        BulletSize = 1 << 6,

        /// <summary>
        /// A bit that specifies whether the bulletChar field of the TextPFException structure that
        /// contains this PFMasks exists.
        /// </summary>
        BulletChar = 1 << 7,

        /// <summary>
        /// A bit that specifies whether the leftMargin field of the TextPFException structure that
        /// contains this PFMasks exists.
        /// </summary>
        LeftMargin = 1 << 8,

        unused9 = 1 << 9, // Bit 9 is reserved

        /// <summary>
        /// A bit that specifies whether the indent field of the TextPFException structure that
        /// contains this PFMasks exists.
        /// </summary>
        Indent = 1 << 10,

        /// <summary>
        /// A bit that specifies whether the textAlignment field of the TextPFException structure
        /// that contains this PFMasks exists.
        /// </summary>
        Align = 1 << 11,

        /// <summary>
        /// A bit that specifies whether the lineSpacing field of the TextPFException structure
        /// that contains this PFMasks exists.
        /// </summary>
        LineSpacing = 1 << 12,

        /// <summary>
        /// A bit that specifies whether the spaceBefore field of the TextPFException that
        /// contains this PFMasks exists.
        /// </summary>
        SpaceBefore = 1 << 13,

        /// <summary>
        /// A bit that specifies whether the spaceAfter field of the TextPFException
        /// structure that contains this PFMasks exists.
        /// </summary>
        SpaceAfter = 1 << 14,

        /// <summary>
        /// A bit that specifies whether the defaultTabSize field of the TextPFException
        /// structure that contains this PFMasks exists.
        /// </summary>
        DefaultTabSize = 1 << 15,

        /// <summary>
        /// A bit that specifies whether the fontAlign field of the TextPFException
        /// structure that contains this PFMasks exists.
        /// </summary>
        FontAlign = 1 << 16,

        /// <summary>
        /// A bit that specifies whether the wrapFlags field of the TextPFException
        /// structure that contains this PFMasks exists and whether wrapFlags.charWrap is valid.
        /// </summary>
        CharWrap = 1 << 17,

        /// <summary>
        /// A bit that specifies whether the wrapFlags field of the TextPFException
        /// structure that contains this PFMasks exists and whether wrapFlags.wordWrap is valid.
        /// </summary>
        WordWrap = 1 << 18,

        /// <summary>
        /// A bit that specifies whether the wrapFlags field of the TextPFException
        /// structure that contains this PFMasks exists and whether wrapFlags.overflow is valid.
        /// </summary>
        Overflow = 1 << 19,

        WrapFlagsFieldExists = CharWrap | WordWrap | Overflow,

        /// <summary>
        /// A bit that specifies whether the tabStops field of the TextPFException
        /// structure that contains this PFMasks exists.
        /// </summary>
        TabStops = 1 << 20,

        /// <summary>
        /// A bit that specifies whether the textDirection field of the TextPFException
        /// structure that contains this PFMasks exists.
        /// </summary>
        TextDirection = 1 << 21,

        unused22 = 1 << 22, // Bit 22 is reserved

        /// <summary>
        /// A bit that specifies whether the bulletBlipRef field of the TextPFException9
        /// structure that contains this PFMasks exists.
        /// </summary>
        BulletBlip = 1 << 23,

        /// <summary>
        /// A bit that specifies whether the bulletAutoNumberScheme field of the TextPFException9
        /// structure that contains this PFMasks exists.
        /// </summary>
        BulletScheme = 1 << 24,

        /// <summary>
        /// A bit that specifies whether the fBulletHasAutoNumber field of the TextPFException9
        /// structure that contains this PFMasks exists.
        /// </summary>
        BulletHasScheme = 1 << 25
    }
}
