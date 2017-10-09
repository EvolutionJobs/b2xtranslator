using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.OlePropertySet
{
    public class PropertySet : List<object>
    {
        private UInt32 size;
        private UInt32 numProperties;
        private UInt32[] identifiers;
        private UInt32[] offsets;

        public PropertySet(VirtualStreamReader stream)
        {
            long pos = stream.BaseStream.Position;

            //read size and number of properties
            this.size = stream.ReadUInt32();
            this.numProperties = stream.ReadUInt32();

            //read the identifier and offsets
            this.identifiers = new UInt32[this.numProperties];
            this.offsets = new UInt32[this.numProperties];
            for (int i = 0; i < this.numProperties; i++)
            {
                this.identifiers[i] = stream.ReadUInt32();
                this.offsets[i] = stream.ReadUInt32();
            }

            //read the properties
            for (int i = 0; i < this.numProperties; i++)
            {
                if (this.identifiers[i] == 0)
                {
                    // dictionary property
                    throw new NotImplementedException("Dictionary Properties are not yet implemented!");
                }
                else
                {
                    // value property
                    this.Add(new ValueProperty(stream));
                }
            }

            // seek to the end of the property set to avoid crashes
            stream.BaseStream.Seek(pos + size, SeekOrigin.Begin);
        }
    }
}
