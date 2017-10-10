using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using DIaLOGIKa.b2xtranslator.Tools;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    [OfficeRecord(1052)]
    public class OriginalMainMasterId : Record
    {
        public UInt32 MainMasterId;

        public OriginalMainMasterId(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.MainMasterId = this.Reader.ReadUInt32();
        }

        override public string ToString(uint depth)
        {
            return String.Format("{0}\n{1}MainMasterId = {2}",
                base.ToString(depth), IndentationForDepth(depth + 1),
                this.MainMasterId);
        }
    }
}
