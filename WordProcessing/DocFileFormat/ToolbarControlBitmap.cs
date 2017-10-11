using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.DocFileFormat
{
    public class ToolbarControlBitmap : ByteStructure
    {
        /// <summary>
        /// Signed integer that specifies the count of total bytes, excluding this field, 
        /// in the TBCBitmap structure plus 10. Value is given by the following formula: <br/>
        /// cbDIB = sizeOf(biHeader) + sizeOf(colors) + sizeOf(bitmapData) + 10<br/>
        /// MUST be greater or equal to 40, and MUST be less or equal to 65576.
        /// </summary>
        public int cbDIB;

        public ToolbarControlBitmap(VirtualStreamReader reader)
            : base(reader, ByteStructure.VARIABLE_LENGTH)
        {
            this.cbDIB = reader.ReadInt32();

            //ToDo: Read TBCBitmap
            reader.ReadBytes(this.cbDIB - 10);
        }
    }
}
