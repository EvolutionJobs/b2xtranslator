

using System;
using System.IO;
using b2xtranslator.OfficeDrawing;

namespace b2xtranslator.PptFileFormat
{
    [OfficeRecord(1055)]
    public class RoundTripShapeId12 : Record
    {
        public uint ShapeId;

        public RoundTripShapeId12(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.ShapeId = this.Reader.ReadUInt32();
        }

        override public string ToString(uint depth)
        {
            return string.Format("{0}\n{1}ShapeId = {2}",
                base.ToString(depth), IndentationForDepth(depth + 1),
                this.ShapeId);
        }
    }

}
