namespace b2xtranslator.DocFileFormat
{
    public class LineSpacingDescriptor
    {
        /// <summary>
        /// 
        /// </summary>
        public short dyaLine;

        /// <summary>
        /// 
        /// </summary>
        public bool fMultLinespace;

        /// <summary>
        /// Creates a new LineSpacingDescriptor with empty values
        /// </summary>
        public LineSpacingDescriptor()
        {
            setDefaultValues();
        }

        /// <summary>
        /// Parses the bytes to retrieve a LineSpacingDescriptor
        /// </summary>
        /// <param name="bytes"></param>
        public LineSpacingDescriptor(byte[] bytes)
        {
            if (bytes.Length == 4)
            {
                this.dyaLine = System.BitConverter.ToInt16(bytes, 0);

                if (System.BitConverter.ToInt16(bytes, 2) == 1)
                {
                    this.fMultLinespace = true;
                }
            }
            else
            {
                throw new ByteParseException("Cannot parse the struct LSPD, the length of the struct doesn't match");
            }
        }

        private void setDefaultValues()
        {
            this.dyaLine = 0;
            this.fMultLinespace = false;
        }
    }
}
