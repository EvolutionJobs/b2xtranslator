using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.DocFileFormat
{
    public class CommandIdentifier: ByteStructure
    {
        public enum CidType
        {
            cmtFci = 0x1,
            cmtMacro = 0x2,
            cmtAllocated = 0x3,
            cmtNil = 0x7
        }

        public const int CID_LENGTH = 4;

        public short ibstMacro;

        public CommandIdentifier(VirtualStreamReader reader)
            : base(reader, CID_LENGTH)
        {
            var bytes = reader.ReadBytes(4);

            var type = (CidType)Utils.BitmaskToInt((int)bytes[0], 0x07);

            switch (type)
            {   
                case CidType.cmtFci:
                    break;
                case CidType.cmtMacro:
                    this.ibstMacro = System.BitConverter.ToInt16(bytes, 2);
                    break;
                case CidType.cmtAllocated:
                    break;
                case CidType.cmtNil:
                    break;
                default:
                    break;
            }
        }
    }
}
