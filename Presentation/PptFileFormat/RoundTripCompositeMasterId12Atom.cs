

using System.IO;
using b2xtranslator.OfficeDrawing;

namespace b2xtranslator.PptFileFormat
{
    [OfficeRecord(1053)]
    public class RoundTripCompositeMasterId12Atom : Record
    {
        public uint compositeMasterId;

        public RoundTripCompositeMasterId12Atom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.compositeMasterId = this.Reader.ReadUInt32();
        }
    }

}
