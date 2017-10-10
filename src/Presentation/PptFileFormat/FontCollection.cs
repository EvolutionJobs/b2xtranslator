

using System.Collections.Generic;
using System.IO;
using b2xtranslator.OfficeDrawing;

namespace b2xtranslator.PptFileFormat
{
    [OfficeRecord(2005)]
    public class FontCollection : RegularContainer
    {
        public List<FontEntityAtom> entities = new List<FontEntityAtom>();

        public FontCollection(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) {

                foreach (var rec in this.Children)
                {
                    if (rec is FontEntityAtom)
                    {
                    this.entities.Add((FontEntityAtom)rec);
                    }
                    else
                    {
                        uint type = rec.TypeCode;
                    }
                }

                return;
        }
    }

}
