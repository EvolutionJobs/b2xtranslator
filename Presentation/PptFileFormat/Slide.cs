

using System.IO;
using b2xtranslator.OfficeDrawing;

namespace b2xtranslator.PptFileFormat
{
    [OfficeRecord(1006)]
    public class Slide : RegularContainer
    {
        public Slide(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) {
                foreach (var rec in this.Children)
                {
                    switch (rec.TypeCode)
                    {
                        case 0x3ef: //SlideAtom
                        case 0x3f9: //SlideShowSlideInfoAtom
                        case 0x40c: //DrawingContainer
                            break;
                        case 0x422: //RoundTripContentMasterId12
                            var id = (RoundTripContentMasterId12)rec;
                            break;
                        case 0x7f0: //SlideSchemeColorSchemeAtom
                        case 0xfa3: //TextMasterStyleAtom
                        case 0xfd9: //SlideHeadersFootersContainer
                            break;
                        case 0x1388: //SlideProgTagsContainer
                            var con = (RegularContainer)rec;
                            break;
                        default:
                            break;
                    }
                }
        }

        public SlidePersistAtom PersistAtom;
    }

}
