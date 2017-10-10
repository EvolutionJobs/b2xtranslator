

using System.Collections.Generic;
using System.IO;
using b2xtranslator.OfficeDrawing;

namespace b2xtranslator.PptFileFormat
{
    //[OfficeRecordAttribute(XXXX)]
    public class Pictures : Record
    {
        public Dictionary<long, Record> _pictures = new Dictionary<long, Record>();

        public Pictures(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.Reader.BaseStream.Position = 0;
            long pos;
            while (this.Reader.BaseStream.Position < this.Reader.BaseStream.Length)
            {
                pos = this.Reader.BaseStream.Position;
                var r = Record.ReadRecord(this.Reader);
                switch (r.TypeCode)
                {
                    case 0:
                        this.Reader.BaseStream.Position = this.Reader.BaseStream.Length;
                        break;
                    case 0xF01A:
                    case 0xF01B:
                    case 0xF01C:
                        var mb = (MetafilePictBlip)r;
                        this._pictures.Add(pos, mb);
                        break;
                    case 0xF01D:
                    case 0xF01E:
                    case 0xF01F:
                    case 0xF020:
                    case 0xF021:
                        var b = (BitmapBlip)r;
                        this._pictures.Add(pos, b);
                        break;
                    default:
                        break;
                }
                
            }
            pos = 1;
        }
    }
}
