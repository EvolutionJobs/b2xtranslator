using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    [OfficeRecordAttribute(0x03FF)]
    public class VBAInfoContainer : Record
    {
        public UInt32 objStgDataRef;
        public UInt32 hasMacros;
        public UInt32 vbaVersion;

        public VBAInfoContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            objStgDataRef = System.BitConverter.ToUInt32(this.RawData, 8);
            hasMacros = System.BitConverter.ToUInt32(this.RawData, 12);
            vbaVersion = System.BitConverter.ToUInt32(this.RawData, 16);
        }
    }
}
