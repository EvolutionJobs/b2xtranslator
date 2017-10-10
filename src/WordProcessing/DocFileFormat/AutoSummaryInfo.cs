using System.Collections;
using b2xtranslator.Tools;

namespace b2xtranslator.DocFileFormat
{
    public class AutoSummaryInfo
    {
        /// <summary>
        /// True if the ASUMYI is valid
        /// </summary>
        public bool fValid;

        /// <summary>
        /// True if AutoSummary View is active
        /// </summary>
        public bool fView;

        /// <summary>
        /// Display method for AutoSummary View: <br/>
        /// 0 = Emphasize in current doc<br/>
        /// 1 = Reduce doc to summary<br/>
        /// 2 = Insert into doc<br/>
        /// 3 = Show in new document
        /// </summary>
        public short iViewBy;

        /// <summary>
        /// True if File Properties summary information 
        /// should be updated after the next summarization
        /// </summary>
        public bool fUpdateProps;

        /// <summary>
        /// Dialog summary level
        /// </summary>
        public short wDlgLevel;

        /// <summary>
        /// Upper bound for lLevel for sentences in this document
        /// </summary>
        public int lHighestLevel;

        /// <summary>
        /// Show document sentences at or below this level
        /// </summary>
        public int lCurrentLevel;

        /// <summary>
        /// Parses the bytes to retrieve a AutoSummaryInfo
        /// </summary>
        /// <param name="bytes">The bytes</param>
        public AutoSummaryInfo(byte[] bytes)
        {
            if (bytes.Length == 12)
            {
                //split byte 0 and 1 into bits
                var bits = new BitArray(new byte[] { bytes[0], bytes[1] });
                this.fValid = bits[0];
                this.fView = bits[1];
                this.iViewBy = (short)Utils.BitArrayToUInt32(Utils.BitArrayCopy(bits, 2, 2));
                this.fUpdateProps = bits[4];

                this.wDlgLevel = System.BitConverter.ToInt16(bytes, 2);
                this.lHighestLevel = System.BitConverter.ToInt32(bytes, 4);
                this.lCurrentLevel = System.BitConverter.ToInt32(bytes, 8);
            }
            else
            {
                throw new ByteParseException("Cannot parse the struct ASUMYI, the length of the struct doesn't match");
            }
        }
    }
}
