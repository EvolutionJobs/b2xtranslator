

using System.IO;
using b2xtranslator.OfficeDrawing;

namespace b2xtranslator.PptFileFormat
{

    [OfficeRecord(4090)]
    public class FooterMCAtom: Record
    {
        public int Position;
        public FooterMCAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.Position = this.Reader.ReadInt32();
        }
    }

    [OfficeRecord(4087)]
    public class DateTimeMCAtom : Record
    {
        public int Position;
        public byte index;
        public DateTimeMCAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.Position = this.Reader.ReadInt32();
            this.index = this.Reader.ReadByte();
        }
    }


    [OfficeRecord(4089)]
    public class HeaderMCAtom : Record
    {
        public int Position;
        public HeaderMCAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.Position = this.Reader.ReadInt32();
        }
    }

    [OfficeRecord(4056)]
    public class SlideNumberMCAtom : Record
    {
        public int Position;
        public SlideNumberMCAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.Position = this.Reader.ReadInt32();
        }
    }

    [OfficeRecord(4088)]
    public class GenericDateMCAtom : Record
    {
        public int Position;
        public GenericDateMCAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.Position = this.Reader.ReadInt32();
        }
    }

}
