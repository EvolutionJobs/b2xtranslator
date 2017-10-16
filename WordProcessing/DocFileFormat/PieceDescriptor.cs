using b2xtranslator.Tools;
using System;
using System.Text;

namespace b2xtranslator.DocFileFormat
{
    public class PieceDescriptor
    {
        /// <summary>File offset of beginning of piece. <br/>
        /// This is relative to the beginning of the WordDocument stream.</summary>
        public uint fc;

        /// <summary>The encoding of the piece</summary>
        public Encoding encoding;

        public int cpStart;

        public int cpEnd;

        /// <summary>Parses the bytes to retrieve a PieceDescriptor</summary>
        /// <param name="bytes">The bytes</param>
        public PieceDescriptor(byte[] bytes)
        {
            if (bytes.Length != 8)
                throw new ByteParseException("Cannot parse the struct PCD, the length of the struct doesn't match");

            //get the fc value
            uint fcValue = BitConverter.ToUInt32(bytes, 2);

            //get the flag
            bool flag = Utils.BitmaskToBool((int)fcValue, 0x40000000);

            //delete the flag
            fcValue = fcValue & 0xBFFFFFFF;

            //find encoding and offset
            if (flag)
            {
                this.encoding = Encoding.GetEncoding("ISO-8859-1"); // windows-1252 not supported by platform
                this.fc = fcValue / 2;
            }
            else
            {
                this.encoding = Encoding.Unicode;
                this.fc = fcValue;
            }
        }
    }
}
