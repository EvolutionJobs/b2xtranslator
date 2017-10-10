using System.Collections.Generic;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class ShapePropsSequence : BiffRecordSequence
    {
        public ShapePropsStream ShapePropsStream;

        public List<ContinueFrt12> ContinueFrt12s;

        public ShapePropsSequence(IStreamReader reader)
            : base(reader)
        {
            // SHAPEPROPS = ShapePropsStream *ContinueFrt12

            // ShapePropsStream
            this.ShapePropsStream = (ShapePropsStream)BiffRecord.ReadRecord(reader);

            // *ContinueFrt12
            this.ContinueFrt12s = new List<ContinueFrt12>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.ContinueFrt12)
            {
                this.ContinueFrt12s.Add((ContinueFrt12)BiffRecord.ReadRecord(reader));
            }
        }
    }
}
