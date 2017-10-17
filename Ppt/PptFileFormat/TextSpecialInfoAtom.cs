

using System.Collections.Generic;
using b2xtranslator.OfficeDrawing;
using System.IO;
using b2xtranslator.Tools;

namespace b2xtranslator.PptFileFormat
{
    [OfficeRecord(4009)]
    public class TextSIExceptionAtom : Record
    {
        public TextSIException si;
        public TextSIExceptionAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.si = new TextSIException(this.Reader);
        }
    }

    [OfficeRecord(4010)]
    public class TextSpecialInfoAtom : Record
    {
        public List<TextSIRun> Runs = new List<TextSIRun>();

        public TextSpecialInfoAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {

            while (this.Reader.BaseStream.Position < this.Reader.BaseStream.Length)
            {
                var run = new TextSIRun(this.Reader);
                this.Runs.Add(run);
            }

        }       
    }

    public class TextSIRun
    {
        public uint count;
        public TextSIException si;

        public TextSIRun(BinaryReader reader)
        {
            this.count = reader.ReadUInt32();
            this.si = new TextSIException(reader);
        }
    }

    public class TextSIException
    {
        private uint flags;
        public bool spell;
        public bool lang;
        public bool altLang;
        public bool fPp10ext;
        public bool fBidi;
        public bool smartTag;
        public ushort spellInfo;
        public ushort lid;
        public ushort bidi;
        public ushort altLid;
        
        public TextSIException(BinaryReader reader)
        {
            this.flags = reader.ReadUInt32();

            this.spell = Utils.BitmaskToBool(this.flags, 0x1);
            this.lang = Utils.BitmaskToBool(this.flags, 0x1 << 1);
            this.altLang = Utils.BitmaskToBool(this.flags, 0x1 << 2);
            this.fPp10ext = Utils.BitmaskToBool(this.flags, 0x1 << 5);
            this.fBidi = Utils.BitmaskToBool(this.flags, 0x1 << 6);
            this.smartTag = Utils.BitmaskToBool(this.flags, 0x1 << 9);

            if (this.spell) this.spellInfo = reader.ReadUInt16();
            if (this.lang) this.lid = reader.ReadUInt16();
            if (this.altLang) this.altLid = reader.ReadUInt16();
            if (this.fBidi) this.bidi = reader.ReadUInt16();
            uint dummy;
            if (this.fPp10ext) dummy = reader.ReadUInt32();
            byte[] smartTags;
            if (this.smartTag) smartTags = reader.ReadBytes((int)(reader.BaseStream.Length - reader.BaseStream.Position));

        }
    }
}
