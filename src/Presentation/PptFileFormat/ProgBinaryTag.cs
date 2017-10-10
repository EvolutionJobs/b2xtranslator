using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    [OfficeRecord(5002)]
    public class ProgBinaryTag : RegularContainer
    {
        public ProgBinaryTag(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }
}
