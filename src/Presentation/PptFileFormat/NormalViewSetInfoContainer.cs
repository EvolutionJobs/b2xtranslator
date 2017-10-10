using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    [OfficeRecord(1044)]
    public class NormalViewSetInfoContainer : RegularContainer
    {
        public NormalViewSetInfoContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

    [OfficeRecord(1018)]
    public class SlideViewInfoContainer : RegularContainer
    {
        public SlideViewInfoContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

    [OfficeRecord(1045)]
    public class NormalViewSetInfoAtom : Record
    {
        public RatioStruct leftPortion;
        public RatioStruct topPortion;
        public byte vertBarState;
        public byte horizBarState;
        public byte fPreferSingleSet;
        public bool fHideThumbnails;
        public bool fBarSnapped;

        public NormalViewSetInfoAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            leftPortion = new RatioStruct(Reader);
            topPortion = new RatioStruct(Reader);
            vertBarState = Reader.ReadByte();
            horizBarState = Reader.ReadByte();

            byte temp = Reader.ReadByte();
            fHideThumbnails = Tools.Utils.BitmaskToBool(temp, 0x1);
            fBarSnapped = Tools.Utils.BitmaskToBool(temp, 0x1 << 1);

        }
    }

    [OfficeRecord(1021)]
    public class ZoomViewInfoAtom : Record
    {

        public ScalingStruct curScale;
        public PointStruct origin;
        public byte fUserVarScale;
        public byte fDraftMode;

        public ZoomViewInfoAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            curScale = new ScalingStruct(Reader);
            Reader.ReadBytes(24); //unused
            origin = new PointStruct(Reader);
            fUserVarScale = Reader.ReadByte();
            fDraftMode = Reader.ReadByte();
        }
    }

    public class PointStruct
    {
        public int x;
        public int y;

        public PointStruct(BinaryReader _reader)
        {
            x = _reader.ReadInt32();
            y = _reader.ReadInt32();
        }
    }

    public class RatioStruct
    {
        public int numer;
        public int denom;

        public RatioStruct(BinaryReader _reader)
        {
            numer = _reader.ReadInt32();
            denom = _reader.ReadInt32();
        }
    }

    public class ScalingStruct
    {
        public RatioStruct x;
        public RatioStruct y;

        public ScalingStruct(BinaryReader _reader)
        {
            x = new RatioStruct(_reader);
            y = new RatioStruct(_reader);
        }
    }

}
