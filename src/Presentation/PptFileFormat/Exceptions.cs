using System;

namespace DIaLOGIKa.b2xtranslator.PptFileFormat
{
    public class InvalidStreamException : Exception
    {
        public InvalidStreamException()
            : base() { }

        public InvalidStreamException(string msg)
            : base(msg) { }

        public InvalidStreamException(string msg, Exception innerException)
            : base(msg, innerException) { }
    }
}
