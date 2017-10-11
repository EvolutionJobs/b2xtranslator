using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;
using System.IO;

namespace b2xtranslator.DocFileFormat
{
    public abstract class ByteStructure
    {
        protected VirtualStreamReader _reader;
        protected int _length;
        protected byte[] _rawBytes;
        public const int VARIABLE_LENGTH = int.MaxValue;

        public byte[] RawBytes => this._rawBytes;


        public ByteStructure(VirtualStreamReader reader, int length) 
        {
            this._reader = reader;
            this._length = length;

            //read the raw bytes
            if (this._length != VARIABLE_LENGTH)
            {
                this._rawBytes = this._reader.ReadBytes(this._length);
                this._reader.BaseStream.Seek(-1 * this._length, SeekOrigin.Current);
            }
        }

        public override string ToString() => 
            Utils.GetHashDump(this._rawBytes);
    }
}
