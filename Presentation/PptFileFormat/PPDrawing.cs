

using System.IO;
using b2xtranslator.OfficeDrawing;

namespace b2xtranslator.PptFileFormat
{
    [OfficeRecord(1036)]
    public class PPDrawing : RegularContainer
    {
        public PPDrawing(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

}
