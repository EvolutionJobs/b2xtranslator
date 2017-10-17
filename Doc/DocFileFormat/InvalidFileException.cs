using System;

namespace b2xtranslator.DocFileFormat
{
    public class UnspportedFileVersionException :Exception
    {
        private const string MSG = "DocFileFormat does not support .doc files that have been created with Word versions older than Word 97.";

        public UnspportedFileVersionException()
            : base(MSG)
        {
        }

        public UnspportedFileVersionException(string text)
            : base(text)
        {
        }
    }
}
