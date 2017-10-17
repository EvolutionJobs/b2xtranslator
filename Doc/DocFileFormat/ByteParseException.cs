using System;

namespace b2xtranslator.DocFileFormat
{
    public class ByteParseException : Exception
    {
        public ByteParseException() : base()
        {
        }

        public ByteParseException(string message) 
            : base(message)
        {
        }

        public ByteParseException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}
