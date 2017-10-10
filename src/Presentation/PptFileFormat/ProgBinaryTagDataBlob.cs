using b2xtranslator.OfficeDrawing;
using System.IO;

namespace b2xtranslator.PptFileFormat
{
    [OfficeRecord(5003)]
    public class ProgBinaryTagDataBlob : RegularContainer
    {
        public ProgBinaryTagDataBlob(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) {

                foreach (var rec in this.Children)
                {
                    switch (rec.TypeCode)
                    {
                        case 0x2b00: //HashCode10Atom
                            break;
                        case 0x2b02: //BuildListContainer
                            break;
                        case 0x36b1: //DocToolbarStates10Atom
                            break;
                        case 0x40d: //GridSpacing10Atom
                            break;
                        case 0x7f8: //BlipCollection9
                            break;
                        case 0x2eeb: //SlideTime10Atom
                            break;
                        case 0xf144: //ExtTimeNodeContainer
                            break;
                        case 0xfad: //TextMasterStyle9Atom
                            break;
                        default:
                            break;
                    }
                }                    
        }
    }
}
