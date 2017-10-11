using b2xtranslator.OfficeDrawing;
using System.IO;

namespace b2xtranslator.PptFileFormat
{
    [OfficeRecord(5000)]
    public class ProgTags : RegularContainer
    {
        public ProgTags(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) {       
        }
    }
}
