
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.String)] 
    public class STRING : BiffRecord
    {
        public const RecordType ID = RecordType.String;

        public string value;

        public int cch;

        public int grbit; 

        public STRING(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            this.cch = reader.ReadUInt16();

            this.grbit = reader.ReadByte();

            this.value = ExcelHelperClass.getStringFromBiffRecord(reader, this.cch, this.grbit); 
	

            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
