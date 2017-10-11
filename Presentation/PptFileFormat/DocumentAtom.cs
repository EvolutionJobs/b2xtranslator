

using System;
using System.IO;
using b2xtranslator.OfficeDrawing;

namespace b2xtranslator.PptFileFormat
{
    public enum SlideSizeType : short
    {
        OnScreen,
        LetterSizedPaper,
        A4Paper,
        Size35mm,
        Overhead,
        Banner,
        Custom
    }

    [OfficeRecord(1001)]
    public class DocumentAtom : Record
    {
        public GPointAtom SlideSize;
        public GPointAtom NotesSize;
        public GRatioAtom ServerZoom;

        public uint NotesMasterPersist;
        public uint HandoutMasterPersist;
        public ushort FirstSlideNum;
        public SlideSizeType SlideSizeType;

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
            this.SlideSizeType = (SlideSizeType) this.Reader.ReadInt16();

            this.SaveWithFonts = this.Reader.ReadByte() != 0;
            this.OmitTitlePlace = this.Reader.ReadByte() != 0;
            this.RightToLeft = this.Reader.ReadByte() != 0;
            this.ShowComments = this.Reader.ReadByte() != 0;
        }

        override public string ToString(uint depth)
        {
            return string.Format("{0}\n{1}SlideSize = {2}, NotesSize = {3}, ServerZoom = {4}\n{1}" +
                "NotesMasterPersist = {5}, HandoutMasterPersist = {6}, FirstSlideNum = {7}, SlideSizeType = {8}\n{1}" +
                "SaveWithFonts = {9}, OmitTitlePlace = {10}, RightToLeft = {11}, ShowComments = {12}",

                base.ToString(depth), IndentationForDepth(depth + 1),

                this.SlideSize, this.NotesSize, this.ServerZoom,

                this.NotesMasterPersist, this.HandoutMasterPersist, this.FirstSlideNum, this.SlideSizeType,

                this.SaveWithFonts, this.OmitTitlePlace, this.RightToLeft, this.ShowComments);
        }
    }

}
