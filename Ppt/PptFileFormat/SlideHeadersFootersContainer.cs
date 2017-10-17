using b2xtranslator.OfficeDrawing;
using System.IO;

namespace b2xtranslator.PptFileFormat
{
    [OfficeRecord(4057)]
    public class SlideHeadersFootersContainer : RegularContainer
    {
        public SlideHeadersFootersContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) {

                foreach (var rec in this.Children)
                {
                    
                }
        
        }
    }

    [OfficeRecord(4058)]
    public class HeadersFootersAtom : Record
    {
        public short formatId;
        public bool fHasDate;
        public bool fHasTodayDate;
        public bool fHasUserDate;
        public bool fHasSlideNumber;
        public bool fHasHeader;
        public bool fHasFooter;
        public HeadersFootersAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.formatId = this.Reader.ReadInt16();

            short mask = this.Reader.ReadInt16();
            this.fHasDate = ((mask & (1)) != 0);
            this.fHasTodayDate = ((mask & (1 << 1)) != 0);
            this.fHasUserDate = ((mask & (1 << 2)) != 0);
            this.fHasSlideNumber = ((mask & (1 << 3)) != 0);
            this.fHasHeader = ((mask & (1 << 4)) != 0);
            this.fHasFooter = ((mask & (1 << 5)) != 0);
        }       
    }

}
