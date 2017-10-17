using b2xtranslator.Tools;

namespace b2xtranslator.DocFileFormat
{
    public class ParagraphHeight
    {
        /// <summary>
        /// Complex shape layout in this paragraph
        /// </summary>
        public bool fVolatile;

        /// <summary>
        /// ParagraphHeight is valid when fUnk is true
        /// </summary>
        public bool fUnk;

        /// <summary>
        /// When true, total height of paragraph is known but lines in 
        /// paragraph have different heights
        /// </summary>
        public bool fDiffLines;

        /// <summary>
        /// When fDiffLines is 0, is number of lines in paragraph
        /// </summary>
        public short clMac;

        /// <summary>
        /// Width of lines in paragraph
        /// </summary>
        public int dxaCol;

        /// <summary>
        /// When fDiffLines is true, is height of every line in paragraph in pixels
        /// </summary>
        public int dymLine;

        /// <summary>
        /// When fDiffLines is true, is the total height in pixels of the paragraph
        /// </summary>
        public int dymHeight;

        /// <summary>
        /// If not == 0, used as a hint when finding the next row.<br/>
        /// (this value is only set if the PHE is stored in a PAP whose fTtp field is set)
        /// </summary>
        public short dcpTtpNext;

        /// <summary>
        /// Height of table row.<br/>
        /// (this value is only set if the PHE is stored in a PAP whose fTtp field is set)
        /// </summary>
        public int dymTableHeight;

        /// <summary>
        /// Reserved
        /// </summary>
        public bool fSpare;

        /// <summary>
        /// Creates a new empty ParagraphHeight with default values
        /// </summary>
        public ParagraphHeight()
        {
            //set default values
            setDefaultValues();
        }

        /// <summary>
        /// Parses the bytes to retrieve a ParagraphHeight
        /// </summary>
        /// <param name="bytes">The bytes</param>
        /// <param name="fTtpMode">
        /// The flag which indicates if the 
        /// ParagraphHeight is stored in a ParagraphProperties whose fTtp field is set
        /// </param>
        public ParagraphHeight(byte[] bytes, bool fTtpMode)
        {
            //set default values
            setDefaultValues();

            if (bytes.Length == 12)
            {
                // The ParagraphHeight is placed in a ParagraphProperties whose fTtp field is set, 
                //so used another bit setting
                if (fTtpMode)
                {
                    this.fSpare = Utils.BitmaskToBool(System.BitConverter.ToInt16(bytes, 0), 0x0001);
                    this.fUnk = Utils.BitmaskToBool(System.BitConverter.ToInt16(bytes, 0), 0x0002);
                    this.dcpTtpNext = System.BitConverter.ToInt16(bytes, 0);
                    this.dxaCol = System.BitConverter.ToInt32(bytes, 4);
                    this.dymTableHeight = System.BitConverter.ToInt32(bytes, 8);
                }
                else
                {
                    this.fVolatile = Utils.BitmaskToBool(System.BitConverter.ToInt16(bytes, 0), 0x0001);
                    this.fUnk = Utils.BitmaskToBool(System.BitConverter.ToInt16(bytes, 0), 0x0002);
                    this.fDiffLines = Utils.BitmaskToBool(System.BitConverter.ToInt16(bytes, 0), 0x0004);
                    this.clMac = System.Convert.ToInt16(((int)System.BitConverter.ToUInt16(bytes, 0)) & 0x00FF);

                    this.dxaCol = System.BitConverter.ToInt32(bytes, 4);
                    this.dymLine = System.BitConverter.ToInt32(bytes, 8);
                    this.dymHeight = System.BitConverter.ToInt32(bytes, 8);
                }
            }
            else
            {
                throw new ByteParseException("Cannot parse the struct ParagraphHeight, the length of the struct doesn't match");
            }
        }

        private void setDefaultValues()
        {
            this.clMac = 0;
            this.dcpTtpNext = 0;
            this.dxaCol = 0;
            this.dymHeight = 0;
            this.dymLine = 0;
            this.dymTableHeight = 0;
            this.fDiffLines = false;
            this.fSpare = false;
            this.fUnk = false;
            this.fVolatile = false;
        }
    }
}
