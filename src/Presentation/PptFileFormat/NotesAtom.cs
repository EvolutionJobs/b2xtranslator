

using System.IO;
using b2xtranslator.OfficeDrawing;

namespace b2xtranslator.PptFileFormat
{
    [OfficeRecord(1009)]
    public class NotesAtom : Record
    {
        public uint SlideIdRef;
        public ushort Flags;

        public NotesAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.SlideIdRef = this.Reader.ReadUInt32();
            this.Flags = this.Reader.ReadUInt16();
            this.Reader.ReadUInt16(); // Throw away undocumented data
        }
    }

}
