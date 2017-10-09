using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    [OfficeRecordAttribute(4057)]
    public class SlideHeadersFootersContainer : RegularContainer
    {
        public SlideHeadersFootersContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) {

                foreach (Record rec in Children)
                {
                    
                }
        
        }
    }

    [OfficeRecordAttribute(4058)]
    public class HeadersFootersAtom : Record
    {
        public Int16 formatId;
        public bool fHasDate;
        public bool fHasTodayDate;
        public bool fHasUserDate;
        public bool fHasSlideNumber;
        public bool fHasHeader;
        public bool fHasFooter;
        public HeadersFootersAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            formatId = this.Reader.ReadInt16();

            Int16 mask = this.Reader.ReadInt16();
            fHasDate = ((mask & (1)) != 0);
            fHasTodayDate = ((mask & (1 << 1)) != 0);
            fHasUserDate = ((mask & (1 << 2)) != 0);
            fHasSlideNumber = ((mask & (1 << 3)) != 0);
            fHasHeader = ((mask & (1 << 4)) != 0);
            fHasFooter = ((mask & (1 << 5)) != 0);
        }       
    }

}
