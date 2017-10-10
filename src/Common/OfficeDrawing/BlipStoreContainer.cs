using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing
{
    [OfficeRecord(0xF001)]
    public class BlipStoreContainer : RegularContainer
    {
        public BlipStoreContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) 
        { 
            
        }
    }
}
