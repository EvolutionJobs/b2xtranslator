using System;
using b2xtranslator.Tools;

namespace b2xtranslator.DocFileFormat
{
    public class OutlineLiSTData
    {
        /// <summary>
        /// An array of  ANLV structures describing how heading numbers 
        /// should be displayed fpr each of Word's 0 outline heading levels
        /// </summary>
        public AutoNumberLevelDescriptor[] rganlv;

        /// <summary>
        /// When true, restart heading on section break
        /// </summary>
        public bool fRestartHdr;

        /// <summary>
        /// Reserved
        /// </summary>
        public bool fSpareOlst2;

        /// <summary>
        /// Reserved
        /// </summary>
        public bool fSpareOlst3;

        /// <summary>
        /// Reserved
        /// </summary>
        public bool fSpareOlst4;

        /// <summary>
        /// Text before/after number
        /// </summary>
        public char[] rgxch;

        /// <summary>
        /// Creates a new OutlineLiSTData with default values
        /// </summary>
        public OutlineLiSTData()
        {
            setDefaultValues();
        }

        /// <summary>
        /// Parses the bytes to retrieve a OutlineLiSTData
        /// </summary>
        /// <param name="bytes">The bytes</param>
        public OutlineLiSTData(byte[] bytes)
        {
            if (bytes.Length == 248)
            {
                //Fill the rganlv array
                int j = 0;
                for (int i = 0; i < 180; i += 20)
                {
                    //copy the 20 byte pages
                    var page = new byte[20];
                    Array.Copy(bytes, i, page, 0, 20);
                    this.rganlv[j] = new AutoNumberLevelDescriptor(page);
                    j++;
                }

                //Set the flags
                this.fRestartHdr = Utils.IntToBool((int)bytes[180]);
                this.fSpareOlst2 = Utils.IntToBool((int)bytes[181]);
                this.fSpareOlst3 = Utils.IntToBool((int)bytes[182]);
                this.fSpareOlst4 = Utils.IntToBool((int)bytes[183]);

                //Fill the rgxch array
                j = 0;
                for (int i = 184; i < 64; i += 2)
                {
                    this.rgxch[j] = Convert.ToChar(System.BitConverter.ToInt16(bytes, i));
                    j++;
                }
            }
            else
            {
                throw new ByteParseException("Cannot parse the struct OLST, the length of the struct doesn't match");
            }
        }

        private void setDefaultValues()
        {
            this.fRestartHdr = false;
            this.fSpareOlst2 = false;
            this.fSpareOlst3 = false;
            this.fSpareOlst4 = false;
            this.rganlv = new AutoNumberLevelDescriptor[9];
            for (int i = 0; i < 9; i++)
            {
                this.rganlv[i] = new AutoNumberLevelDescriptor();
            }
            this.rgxch = Utils.ClearCharArray(new char[32]);
        }
    }
}
