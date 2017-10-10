
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.ExternSheet)] 
    public class ExternSheet : BiffRecord
    {
        public const RecordType ID = RecordType.ExternSheet;

        public ushort cXTI;

        public ushort[] iSUPBOOK; 

        public ushort[] itabFirst;

        public ushort[] itabLast;

        public ExternSheet(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            this.cXTI = this.Reader.ReadUInt16();

            this.iSUPBOOK = new ushort[this.cXTI];
            this.itabFirst = new ushort[this.cXTI]; 
            this.itabLast = new ushort[this.cXTI]; 

            for (int i = 0; i < this.cXTI; i++)
            {                
                this.iSUPBOOK[i] = this.Reader.ReadUInt16();
                this.itabFirst[i] = this.Reader.ReadUInt16();
                this.itabLast[i] = this.Reader.ReadUInt16(); 
            }


            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
