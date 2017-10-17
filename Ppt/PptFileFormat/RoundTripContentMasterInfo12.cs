

using System.IO;
using b2xtranslator.OfficeDrawing;

namespace b2xtranslator.PptFileFormat
{
    [OfficeRecord(1054)]
    public class RoundTripContentMasterInfo12 : XmlContainer
    {
        public RoundTripContentMasterInfo12(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
        }
    }

}
