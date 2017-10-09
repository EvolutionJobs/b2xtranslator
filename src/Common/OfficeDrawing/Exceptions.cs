using System;
using System.Collections.Generic;
using System.Text;

namespace DIaLOGIKa.b2xtranslator.OfficeDrawing
{
    public class InvalidRecordException : Exception
    {
        public InvalidRecordException()
            : base() { }

        public InvalidRecordException(string msg)
            : base(msg) { }

        public InvalidRecordException(string msg, Exception innerException)
            : base(msg, innerException) { }
    }

}
