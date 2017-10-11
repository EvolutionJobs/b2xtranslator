using b2xtranslator.OfficeDrawing;
using System.IO;

namespace b2xtranslator.PptFileFormat
{
    [OfficeRecord(4055)]
    public class ExHyperlinkContainer : RegularContainer
    {
        public ExHyperlinkContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

    [OfficeRecord(4051)]
    public class ExHyperlinkAtom : Record
    {
        public uint exHyperlinkId;

        public ExHyperlinkAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.exHyperlinkId = this.Reader.ReadUInt32();
        }
    }

    [OfficeRecord(4068)]
    public class ExHyperlink9Container : RegularContainer
    {
        public ExHyperlink9Container(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }
}
