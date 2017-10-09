using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.OlePropertySet
{
    public class ValueProperty
    {
        public enum PropertyType
        {
            Empty = 0x0000,
            Null = 0x0001,
            SignedInt16 = 0x0002,
            SignedInt32 = 0x0003,
            FloatingPoint32 = 0x0004,
            FloatingPoint64 = 0x0005,
            Currency = 0x0006,
            Date = 0x0007,
            CodePageString1 = 0x0008,
            HResult = 0x000A,
            Boolean = 0x000B,
            Decimal = 0x000E,
            SignedByte = 0x0010,
            UnsignedByte = 0x0011,
            UnsignedInt16 = 0x0012,
            UnsignedInt32 = 0x0013,
            SignedInt64 = 0x0014,
            UsignedInt64 = 0x0015,
            NewSignedInt32 = 0x0016,
            NewUnsignedInt32 = 0x0017,
            CodePageString2 = 0x001E,
            UnicodeString = 0x001F,
            FileTime = 0x0040,
            Blob = 0x0041,
            Stream = 0x0042,
            Storage = 0x0043,
            StreamObject = 0x0044,
            StorageObject = 0x0045,
            BlobObject = 0x0046,
            PropertyId = 0x0047,
            ClassId = 0x0048,
            VersionedStream = 0x0049,
            VectorOfSignedInt16 = 0x1002,
            VectorOfSignedInt32 = 0x1003,
            VectorOfFloatingPoint32 = 0x1004,
            VectorOfFloatingPoint64 = 0x1005,
            VectorOfCurrency = 0x1006,
            VectorOfDate = 0x1007,
            VectorOfCodePageString1 = 0x1008,
            VectorOfHResult = 0x100A,
            VectorOfBoolean = 0x100B,
            VectorOfVariables = 0x100C,
            VectorOfSignedByte = 0x1010,
            VectorOfUnsignedByte = 0x1011,
            VectorOfUnsignedInt16 = 0x1012,
            VectorOfUnsignedInt32 = 0x1013,
            VectorOfSignedInt64 = 0x1014,
            VectorOfUsignedInt64 = 0x1015,
            VectorOfCodePageString2 = 0x101E,
            VectorOfUnicodeString = 0x101F,
            VectorOfFileTime = 0x1040,
            VectorOfPropertyId = 0x1047,
            VectorOfClassId = 0x1048,
            ArrayOfSignedInt16 = 0x2002,
            ArrayOfSignedInt32 = 0x2003,
            ArrayOfFloatingPoint32 = 0x2004,
            ArrayOfFloatingPoint64 = 0x2005,
            ArrayOfCurrency = 0x2006,
            ArrayOfDate = 0x2007,
            ArrayOfCodePageString1 = 0x2008,
            ArrayOfHResult = 0x200A,
            ArrayOfBoolean = 0x200B,
            ArrayOfVariables = 0x200C,
            ArrayOfDecimal = 0x200E,
            ArrayOfSignedByte = 0x2010,
            ArrayOfUnsignedByte = 0x2011,
            ArrayOfUnsignedInt16 = 0x2012,
            ArrayOfUnsignedInt32 = 0x2013,
            ArrayOfSignedInt4Byte = 0x2016,
            ArrayOfUnsignedInt4Byte = 0x2017
        }

        public PropertyType Type;

        public byte[] Data;

        public ValueProperty(VirtualStreamReader stream)
        {
            //read type
            this.Type = (PropertyType)stream.ReadUInt16();

            //skip padding
            stream.ReadBytes(2);

            //read data
            if(
                Type == PropertyType.SignedInt16 ||
                Type == PropertyType.UnsignedInt16
                )
            {
                // 2 bytes data
                this.Data = stream.ReadBytes(2);
            }
            else if (
                Type == PropertyType.SignedInt32 ||
                Type == PropertyType.UnsignedInt32 ||
                Type == PropertyType.FloatingPoint32 ||
                Type == PropertyType.NewSignedInt32 ||
                Type == PropertyType.NewUnsignedInt32 ||
                Type == PropertyType.HResult ||
                Type == PropertyType.Boolean)
            {
                // 4 bytes data
                this.Data = stream.ReadBytes(4);
            }
            else if(
                Type == PropertyType.FloatingPoint64 ||
                Type == PropertyType.SignedInt64 ||
                Type == PropertyType.UsignedInt64 ||
                Type == PropertyType.Currency ||
                Type == PropertyType.Date
                )
            {
                // 8 bytes data
                this.Data = stream.ReadBytes(8);
            }
            else if (
                Type == PropertyType.Decimal
                )
            {
                // 16 bytes data
                this.Data = stream.ReadBytes(16);
            }
            else
            {
                // not yet implemented
                this.Data = new byte[0];
            }
        }
    }
}
