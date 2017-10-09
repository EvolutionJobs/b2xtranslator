using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class BiffRecordSequence
    {
        IStreamReader _reader;
        public IStreamReader Reader
        {
            get { return _reader; }
            set { this._reader = value; }
        }

        public BiffRecordSequence(IStreamReader reader)
        {
            _reader = reader;
        }
    }
}
