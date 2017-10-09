using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    [OfficeRecordAttribute(4082)]
    public class MouseClickInteractiveInfoContainer : RegularContainer
    {
        public MouseClickTextInteractiveInfoAtom Range;

        public MouseClickInteractiveInfoContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) {
        
        }
    }

    [OfficeRecordAttribute(4063)]
    public class MouseClickTextInteractiveInfoAtom : Record
    {
        public Int32 begin;
        public Int32 end;

        public MouseClickTextInteractiveInfoAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            begin = this.Reader.ReadInt32();
            end = this.Reader.ReadInt32();
        }
    }

    [OfficeRecordAttribute(4083)]
    public class InteractiveInfoAtom : Record
    {
        public UInt32 SoundIdRef;
        public UInt32 exHyperlinkIdRef;
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
            SoundIdRef = this.Reader.ReadUInt32();
            exHyperlinkIdRef = this.Reader.ReadUInt32();

            action = (InteractiveInfoActionEnum)this.Reader.ReadByte();
            oleVerb = this.Reader.ReadByte();
            jump = this.Reader.ReadByte();

            byte mask = this.Reader.ReadByte();
            fAnimated = ((mask & (1)) != 0);
            fStopSound = ((mask & (1 << 1)) != 0);
            fCustomShowReturn = ((mask & (1 << 2)) != 0);
            fVisited = ((mask & (1 << 3)) != 0);
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
