using System;
using b2xtranslator.OfficeDrawing;
using System.IO;

namespace b2xtranslator.PptFileFormat
{
    [OfficeRecord(1052)]
    public class OriginalMainMasterId : Record
    {
        public uint MainMasterId;

        public OriginalMainMasterId(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.MainMasterId = this.Reader.ReadUInt32();
        }

        override public string ToString(uint depth)
        {
            return string.Format("{0}\n{1}MainMasterId = {2}",
                base.ToString(depth), IndentationForDepth(depth + 1),
                this.MainMasterId);
        }
    }
}
