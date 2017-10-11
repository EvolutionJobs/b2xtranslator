using b2xtranslator.OfficeDrawing;
using System.IO;

namespace b2xtranslator.PptFileFormat
{
    [OfficeRecord(0x03FF)]
    public class VBAInfoContainer : Record
    {
        public uint objStgDataRef;
        public uint hasMacros;
        public uint vbaVersion;

        public VBAInfoContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.objStgDataRef = System.BitConverter.ToUInt32(this.RawData, 8);
            this.hasMacros = System.BitConverter.ToUInt32(this.RawData, 12);
            this.vbaVersion = System.BitConverter.ToUInt32(this.RawData, 16);
        }
    }
}
