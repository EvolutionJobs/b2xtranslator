using System;
using System.Collections.Generic;
using System.Text;
using DIaLOGIKa.b2xtranslator.OfficeDrawing;
using System.IO;
using System.IO.Compression;
using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    [OfficeRecordAttribute(1033)]
    public class ExObjListContainer : RegularContainer
    {
        public ExObjListContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }


    [OfficeRecordAttribute(1034)]
    public class ExObjListAtom : Record
    {
        public ExObjListAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {

        }
    }

    [OfficeRecordAttribute(3009)]
    public class ExObjRefAtom : Record
    {
        public Int32 exObjIdRef;

        public ExObjRefAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            exObjIdRef = this.Reader.ReadInt32();
        }
    }

    [OfficeRecordAttribute(4035)]
    public class ExOleObjAtom : Record
    {
        public UInt32 persistIdRef;
        public Int32 exObjId;

        public ExOleObjAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            Int32 drawAspect = this.Reader.ReadInt32();
            Int32 type = this.Reader.ReadInt32();
            exObjId = this.Reader.ReadInt32();
            Int32 subType = this.Reader.ReadInt32();
            persistIdRef = this.Reader.ReadUInt32();
            Int32 unused = this.Reader.ReadInt32();            
        }
    }

    [OfficeRecordAttribute(4044)]
    public class ExOleEmbedContainer : RegularContainer
    {
        public ExOleObjStgAtom stgAtom;

        public ExOleEmbedContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

    [OfficeRecordAttribute(4045)]
    public class ExOleEmbedAtom : Record
    {
        public ExOleEmbedAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            Int32 exColorFollow = this.Reader.ReadInt32();
            byte fCantLockServer = this.Reader.ReadByte();
            byte fNoSizeToServer = this.Reader.ReadByte();
            byte fIsTable = this.Reader.ReadByte();

        }
    }

    [OfficeRecordAttribute(4113)]
    public class ExOleObjStgAtom : Record, IVisitable
    {
        public uint len = 0;
        public UInt32 decompressedSize = 0;
        public byte[] data;

        public ExOleObjStgAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            if (instance == 0)
            {
                //uncompressed
                data = this.Reader.ReadBytes((int)size);
            }
            else
            {
                //compressed
                decompressedSize = this.Reader.ReadUInt32();
                len = size - 4;
                data = this.Reader.ReadBytes((int)len);
            }
        }

        public byte[] DecompressData()
        {
            // create memory stream to the data
            MemoryStream msCompressed = new MemoryStream(data);
            
            // skip the first 2 bytes
            msCompressed.ReadByte();
            msCompressed.ReadByte();

            // decompress the bytes
            byte[] decompressedBytes = new byte[decompressedSize];
            DeflateStream deflateStream = new DeflateStream(msCompressed, CompressionMode.Decompress, true);
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
