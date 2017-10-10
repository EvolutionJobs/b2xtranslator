namespace b2xtranslator.PptFileFormat
{
    //[OfficeRecordAttribute(0x1011)]
    //public class VbaProjectAtom : Record
    //{
    //    /// <summary>
    //    /// The VBA OLE Storage
    //    /// </summary>
    //    public StructuredStorageReader Storage;

    //    public VbaProjectAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
    //        : base(_reader, size, typeCode, version, instance)
    //    {
    //        if (instance == 0)
    //        {
    //            // it's a uncompressed vba project
    //            // the body is the storage
    //            this.Storage = new StructuredStorageReader(_reader.BaseStream);
    //        }
    //        else if (instance == 1)
    //        {
    //            // it's a compressed vba project
    //            UInt32 decompSize = System.BitConverter.ToUInt32(this.RawData, 0);

    //            // get the compressed bytes
    //            byte[] compressedBytes = new byte[size - 4];
    //            Array.Copy(this.RawData, 4, compressedBytes, 0, compressedBytes.Length);
    //            MemoryStream msCompressed = new MemoryStream(compressedBytes);

    //            byte[] decompressedBytes = new byte[decompSize];

    //            #region DecompressionDotNet
    //            //decompress the bytes using .NET DeflateStream class.
    //            DeflateStream deflateStream = new DeflateStream(msCompressed, CompressionMode.Decompress, true);
    //            deflateStream.Read(decompressedBytes, 0, decompressedBytes.Length);
    //            #endregion

    //            //#region DecompressionIonicZlib
    //            //// decompress the bytes using Ionic zip library.
    //            //Ionic.Zlib.DeflateStream deflateStream = new Ionic.Zlib.DeflateStream(msCompressed, Ionic.Zlib.CompressionMode.Decompress);
    //            //deflateStream.Read(decompressedBytes, 0, decompressedBytes.Length);
    //            //#endregion

    //            // create the storage
    //            MemoryStream msDecompressed = new MemoryStream(decompressedBytes);
    //            this.Storage = new StructuredStorageReader(msDecompressed);
    //        }
    //    }

    //    //private int readAllBytesFromStream(DeflateStream stream, byte[] buffer, int blocksize)
    //    //{
    //    //    int offset = 0;
    //    //    int totalCount = 0;
    //    //    while (true)
    //    //    {
    //    //        int bytesRead = stream.Read(buffer, offset, blocksize);
    //    //        if (bytesRead == 0)
    //    //        {
    //    //            break;
    //    //        }
    //    //        offset += bytesRead;
    //    //        totalCount += bytesRead;
    //    //    }
    //    //    return totalCount;
    //    //} 

    //}
}
