using b2xtranslator.OfficeDrawing;
using System.IO;

namespace b2xtranslator.PptFileFormat
{
    [OfficeRecord(1010)]
    public class Environment : RegularContainer
    {
        public Environment(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) {

                foreach (var rec in this.Children)
                {
                    switch (rec.TypeCode)
                    {
                        case 0x7d5: //FontCollectionContainer
                            break;
                        case 0xfa3: //TextMasterStyleAtom
                            var a = (TextMasterStyleAtom)rec;
                            break;
                        case 0xfa4: //TextCFExceptionAtom
                            var ce = (TextCFExceptionAtom)rec;
                            break;
                        case 0xfa5: //TextPFExceptionAtom
                            var e = (TextPFExceptionAtom)rec;
                            break;
                        case 0xfa9: //TextSIEExceptionAtom
                            break;
                        case 0xfc8: //KinsokuContainer
                            break;
                        default:
                            break;
                    }
                }
        
        }
    }
}
