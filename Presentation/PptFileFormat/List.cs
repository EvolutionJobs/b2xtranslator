

using System.IO;
using b2xtranslator.OfficeDrawing;

namespace b2xtranslator.PptFileFormat
{
    // Wrongly listed in documentation as 1016
    [OfficeRecord(2000)]
    public class List : RegularContainer
    {
        public List(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

}
