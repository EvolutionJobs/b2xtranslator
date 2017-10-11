using System;
using b2xtranslator.Tools;

namespace b2xtranslator.DocFileFormat
{
    public class AutoNumberedListDataDescriptor
    {
        public AutoNumberLevelDescriptor anlv;

        /// <summary>
        /// Number only 1 item per table cell
        /// </summary>
        public bool fNumber1;

        /// <summary>
        /// Number across cells in table rows
        /// </summary>
        public bool fNumberAcross;

        /// <summary>
        /// Restart heading number on section boundary
        /// </summary>
        public bool fRestartHdn;

        /// <summary>
        /// Not used
        /// </summary>
        public bool fSpareX;
	
        /// <summary>
        /// Characters displayed before/after auto number
        /// </summary>
        public char[] rgxch;

        /// <summary>
        /// Creates a new AutoNumberedListDataDescriptor with defaut values
        /// </summary>
        public AutoNumberedListDataDescriptor()
        {
            setDefaultValues();
        }

        /// <summary>
        /// Parses the bytes to retrieve a AutoNumberedListDataDescriptor
        /// </summary>
        /// <param name="bytes">The bytes</param>
        public AutoNumberedListDataDescriptor(byte[] bytes)
        {
            if (bytes.Length == 88)
            {
                //copies the first 20 bytes into a new array
                var anlvArray = new byte[20];
                Array.Copy(bytes, anlvArray, anlvArray.Length);
                //parse these bytes 
                this.anlv = new AutoNumberLevelDescriptor(anlvArray);

                //parse the rest
                if (bytes[20] == 1)
                    this.fNumber1 = true;
                if (bytes[21] == 1)
                    this.fNumberAcross = true;
                if (bytes[22] == 1)
                    this.fRestartHdn = true;
                if (bytes[23] == 1)
                    this.fSpareX = true;

                this.rgxch = new char[32];
                int j=0;
                for (int i = 24; i <= 88; i += 2)
                {
                    this.rgxch[j] = Convert.ToChar(BitConverter.ToInt16(bytes, i));
                    j++;
                }
            }
            else
            {
                throw new ByteParseException(
                    "Cannot parse the struct ANLD, the length of the struct doesn't match");
            }
        }

        private void setDefaultValues()
        {
            this.fNumber1 = false;
            this.fNumberAcross = false;
            this.fRestartHdn = false;
            this.fSpareX = false;
            this.rgxch = Utils.ClearCharArray(new char[32]);
        }
    }
}
