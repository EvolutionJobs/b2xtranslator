using System;

namespace b2xtranslator.DocFileFormat
{
    public class DropCapSpecifier
    {
        /// <summary>
        /// drop cap type can be:<br/>
        /// 0 no drop cap
        /// 1 normal drop cap
        /// 2 drop cap in margin
        /// </summary>
        public byte Type;

        /// <summary>
        /// Count of lines to drop
        /// </summary>
        public byte Count;

        /// <summary>
        /// Creates a new DropCapSpecifier with default values
        /// </summary>
        public DropCapSpecifier()
        {
            setDefaultValues();
        }

        /// <summary>
        /// Parses the bytes to retrieve a DropCapSpecifier
        /// </summary>
        /// <param name="bytes"></param>
        public DropCapSpecifier(byte[] bytes)
        {
            if (bytes.Length == 2)
            {
                byte val = bytes[0];
                this.Type = Convert.ToByte((int)val & (int)0x0007);
                this.Count = Convert.ToByte((int)val & (int)0x00F8);
            }
            else
            {
                throw new ByteParseException(
                    "Cannot parse the struct DCS, the length of the struct doesn't match");
            }
        }

        private void setDefaultValues()
        {
            this.Count = 0;
            this.Type = 0;
        }
    }
}
