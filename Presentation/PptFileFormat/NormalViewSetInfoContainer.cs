using b2xtranslator.OfficeDrawing;
using System.IO;

namespace b2xtranslator.PptFileFormat
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
            this.leftPortion = new RatioStruct(this.Reader);
            this.topPortion = new RatioStruct(this.Reader);
            this.vertBarState = this.Reader.ReadByte();
            this.horizBarState = this.Reader.ReadByte();

            byte temp = this.Reader.ReadByte();
            this.fHideThumbnails = Tools.Utils.BitmaskToBool(temp, 0x1);
            this.fBarSnapped = Tools.Utils.BitmaskToBool(temp, 0x1 << 1);

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
            this.curScale = new ScalingStruct(this.Reader);
            this.Reader.ReadBytes(24); //unused
            this.origin = new PointStruct(this.Reader);
            this.fUserVarScale = this.Reader.ReadByte();
            this.fDraftMode = this.Reader.ReadByte();
        }
    }

    public class PointStruct
    {
        public int x;
        public int y;

        public PointStruct(BinaryReader _reader)
        {
            this.x = _reader.ReadInt32();
            this.y = _reader.ReadInt32();
        }
    }

    public class RatioStruct
    {
        public int numer;
        public int denom;

        public RatioStruct(BinaryReader _reader)
        {
            this.numer = _reader.ReadInt32();
            this.denom = _reader.ReadInt32();
        }
    }

    public class ScalingStruct
    {
        public RatioStruct x;
        public RatioStruct y;

        public ScalingStruct(BinaryReader _reader)
        {
            this.x = new RatioStruct(_reader);
            this.y = new RatioStruct(_reader);
        }
    }

}
