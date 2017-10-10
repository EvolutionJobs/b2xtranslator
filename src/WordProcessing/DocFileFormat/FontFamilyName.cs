using System.Text;
using b2xtranslator.Tools;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.DocFileFormat
{
    public class FontFamilyName : ByteStructure
    {
        public struct FontSignature
        {
            public uint UnicodeSubsetBitfield0;
            public uint UnicodeSubsetBitfield1;
            public uint UnicodeSubsetBitfield2;
            public uint UnicodeSubsetBitfield3;
            public uint CodePageBitfield0;
            public uint CodePageBitfield1;
        }

        /// <summary>
        /// When true, font is a TrueType font
        /// </summary>
        public bool fTrueType;

        /// <summary>
        /// Font family id
        /// </summary>
        public byte ff;

        /// <summary>
        /// Base weight of font
        /// </summary>
        public short wWeight;

        /// <summary>
        /// Character set identifier
        /// </summary>
        public byte chs;

        /// <summary>
        /// Pitch request
        /// </summary>
        public byte prq;

        /// <summary>
        /// Name of font
        /// </summary>
        public string xszFtn;

        /// <summary>
        /// Alternative name of the font
        /// </summary>
        public string xszAlt;

        /// <summary>
        /// Panose
        /// </summary>
        public byte[] panose;

        /// <summary>
        /// Font sinature
        /// </summary>
        public FontSignature fs;


        public FontFamilyName(VirtualStreamReader reader, int length) : base(reader, length)
        {
            long startPos = this._reader.BaseStream.Position;

            //FFID
            int ffid = (int)this._reader.ReadByte();

            int req = ffid;
            req = req << 6;
            req = req >> 6;
            this.prq = (byte)req;

            this.fTrueType = Utils.BitmaskToBool(ffid, 0x04);

            int family = ffid;
            family = family << 1;
            family = family >> 4;
            this.ff = (byte)family;

            this.wWeight = this._reader.ReadInt16();

            this.chs = this._reader.ReadByte();

            //skip byte 5
            this._reader.ReadByte();

            //read the 10 bytes panose
            this.panose = this._reader.ReadBytes(10);

            //read the 24 bytes FontSignature
            this.fs = new FontSignature
            {
                UnicodeSubsetBitfield0 = this._reader.ReadUInt32(),
                UnicodeSubsetBitfield1 = this._reader.ReadUInt32(),
                UnicodeSubsetBitfield2 = this._reader.ReadUInt32(),
                UnicodeSubsetBitfield3 = this._reader.ReadUInt32(),
                CodePageBitfield0 = this._reader.ReadUInt32(),
                CodePageBitfield1 = this._reader.ReadUInt32()
            };

            //read the next \0 terminated string
            long strStart = reader.BaseStream.Position;
            long strEnd = searchTerminationZero(this._reader);
            this.xszFtn = Encoding.Unicode.GetString(this._reader.ReadBytes((int)(strEnd - strStart)));
            this.xszFtn = this.xszFtn.Replace("\0", "");

            long readBytes = this._reader.BaseStream.Position - startPos;
            if(readBytes < this._length)
            {
                //read the next \0 terminated string
                strStart = reader.BaseStream.Position;
                strEnd = searchTerminationZero(this._reader);
                this.xszAlt = Encoding.Unicode.GetString(this._reader.ReadBytes((int)(strEnd - strStart)));
                this.xszAlt = this.xszAlt.Replace("\0", "");
            }
        }

        private long searchTerminationZero(VirtualStreamReader reader)
        {
            long strStart = reader.BaseStream.Position;
            while (reader.ReadInt16() != 0)
            {
                ;
            }
            long pos = reader.BaseStream.Position;
            reader.BaseStream.Seek(strStart, System.IO.SeekOrigin.Begin);
            return pos;
        }
    }
}
