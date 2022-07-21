

using System.Text;
using System.IO;
using b2xtranslator.OfficeDrawing;

namespace b2xtranslator.PptFileFormat
{
    [OfficeRecord(4023)]
    public class FontEntityAtom : Record
    {
        public string TypeFace = "";

        public FontEntityAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
                        
            var facename = this.Reader.ReadBytes(64);
            this.TypeFace = Encoding.Unicode.GetString(facename);
            this.TypeFace = this.TypeFace.TrimEnd('\0');

            //TODO: read other flags

            byte lfCharSet = this.Reader.ReadByte();
            byte firstbyte = this.Reader.ReadByte();
            byte secondbyte = this.Reader.ReadByte();
            byte lfPitchAndFamily = this.Reader.ReadByte();           
        }

    }
}

