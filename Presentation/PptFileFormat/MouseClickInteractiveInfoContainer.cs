using b2xtranslator.OfficeDrawing;
using System.IO;

namespace b2xtranslator.PptFileFormat
{
    [OfficeRecord(4082)]
    public class MouseClickInteractiveInfoContainer : RegularContainer
    {
        public MouseClickTextInteractiveInfoAtom Range;

        public MouseClickInteractiveInfoContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) {
        
        }
    }

    [OfficeRecord(4063)]
    public class MouseClickTextInteractiveInfoAtom : Record
    {
        public int begin;
        public int end;

        public MouseClickTextInteractiveInfoAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.begin = this.Reader.ReadInt32();
            this.end = this.Reader.ReadInt32();
        }
    }

    [OfficeRecord(4083)]
    public class InteractiveInfoAtom : Record
    {
        public uint SoundIdRef;
        public uint exHyperlinkIdRef;
        public InteractiveInfoActionEnum action;
        public byte oleVerb;
        public byte jump;

        public bool fAnimated;
        public bool fStopSound;
        public bool fCustomShowReturn;
        public bool fVisited;

        public byte hyperlinkType;

        public InteractiveInfoAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.SoundIdRef = this.Reader.ReadUInt32();
            this.exHyperlinkIdRef = this.Reader.ReadUInt32();

            this.action = (InteractiveInfoActionEnum)this.Reader.ReadByte();
            this.oleVerb = this.Reader.ReadByte();
            this.jump = this.Reader.ReadByte();

            byte mask = this.Reader.ReadByte();
            this.fAnimated = ((mask & (1)) != 0);
            this.fStopSound = ((mask & (1 << 1)) != 0);
            this.fCustomShowReturn = ((mask & (1 << 2)) != 0);
            this.fVisited = ((mask & (1 << 3)) != 0);
        }       
    }


    public enum InteractiveInfoActionEnum
    {
        No = 0,
        Macro,
        RunProgram,
        Jump,
        Hyperlink,
        OLE,
        Media,
        CustomShow
    }

}
