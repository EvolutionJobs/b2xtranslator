

using System;
using System.IO;
using b2xtranslator.OfficeDrawing;

namespace b2xtranslator.PptFileFormat
{
    [OfficeRecord(1007)]
    public class SlideAtom : Record
    {
        public SSlideLayoutAtom Layout;
        public uint MasterId;
        public int NotesId;
        public ushort Flags;

        public SlideAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.Layout = new SSlideLayoutAtom(this.Reader);
            this.MasterId = this.Reader.ReadUInt32();
            this.NotesId = this.Reader.ReadInt32();
            this.Flags = this.Reader.ReadUInt16();
            this.Reader.ReadUInt16(); // Throw away undocumented data
        }

        override public string ToString(uint depth)
        {
            return string.Format("{0}\n{1}Layout = {2}\n{1}MasterId = {3}, NotesId = {4}, Flags = {5})",
                base.ToString(depth), IndentationForDepth(depth + 1),
                this.Layout, this.MasterId, this.NotesId, this.Flags);
        }
    }

}
