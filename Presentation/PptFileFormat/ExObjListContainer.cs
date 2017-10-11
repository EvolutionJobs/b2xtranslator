using b2xtranslator.OfficeDrawing;
using System.IO;
using System.IO.Compression;
using b2xtranslator.CommonTranslatorLib;

namespace b2xtranslator.PptFileFormat
{
    [OfficeRecord(1033)]
    public class ExObjListContainer : RegularContainer
    {
        public ExObjListContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }


    [OfficeRecord(1034)]
    public class ExObjListAtom : Record
    {
        public ExObjListAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {

        }
    }

    [OfficeRecord(3009)]
    public class ExObjRefAtom : Record
    {
        public int exObjIdRef;

        public ExObjRefAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.exObjIdRef = this.Reader.ReadInt32();
        }
    }

    [OfficeRecord(4035)]
    public class ExOleObjAtom : Record
    {
        public uint persistIdRef;
        public int exObjId;

        public ExOleObjAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            int drawAspect = this.Reader.ReadInt32();
            int type = this.Reader.ReadInt32();
            this.exObjId = this.Reader.ReadInt32();
            int subType = this.Reader.ReadInt32();
            this.persistIdRef = this.Reader.ReadUInt32();
            int unused = this.Reader.ReadInt32();            
        }
    }

    [OfficeRecord(4044)]
    public class ExOleEmbedContainer : RegularContainer
    {
        public ExOleObjStgAtom stgAtom;

        public ExOleEmbedContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

    [OfficeRecord(4045)]
    public class ExOleEmbedAtom : Record
    {
        public ExOleEmbedAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            int exColorFollow = this.Reader.ReadInt32();
            byte fCantLockServer = this.Reader.ReadByte();
            byte fNoSizeToServer = this.Reader.ReadByte();
            byte fIsTable = this.Reader.ReadByte();

        }
    }

    [OfficeRecord(4113)]
    public class ExOleObjStgAtom : Record, IVisitable
    {
        public uint len = 0;
        public uint decompressedSize = 0;
        public byte[] data;

        public ExOleObjStgAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            if (instance == 0)
            {
                //uncompressed
                this.data = this.Reader.ReadBytes((int)size);
            }
            else
            {
                //compressed
                this.decompressedSize = this.Reader.ReadUInt32();
                this.len = size - 4;
                this.data = this.Reader.ReadBytes((int)this.len);
            }
        }

        public byte[] DecompressData()
        {
            // create memory stream to the data
            var msCompressed = new MemoryStream(this.data);
            
            // skip the first 2 bytes
            msCompressed.ReadByte();
            msCompressed.ReadByte();

            // decompress the bytes
            var decompressedBytes = new byte[this.decompressedSize];
            var deflateStream = new DeflateStream(msCompressed, CompressionMode.Decompress, true);
            deflateStream.Read(decompressedBytes, 0, decompressedBytes.Length);

            return decompressedBytes;
        }

        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<ExOleObjStgAtom>)mapping).Apply(this);
        }

        #endregion
    }
}
