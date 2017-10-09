using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using DIaLOGIKa.b2xtranslator.StructuredStorage.Reader;

namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class ChartSheetSequence : BiffRecordSequence, IVisitable
    {
        public BOF BOF;

        public ChartFrtInfo ChartFrtInfo;

        public ChartSheetContentSequence ChartSheetContentSequence;

        public ChartSheetSequence(IStreamReader reader) : base(reader)
        {
            //BOF 
            this.BOF = (BOF)BiffRecord.ReadRecord(reader);

            // [ChartFrtInfo] (not specified)
            if (BiffRecord.GetNextRecordType(reader) == RecordType.ChartFrtInfo)
            {
                this.ChartFrtInfo = (ChartFrtInfo)BiffRecord.ReadRecord(reader);
            }

            //CHARTSHEETCONTENT
            this.ChartSheetContentSequence = new ChartSheetContentSequence(reader);
        }

        #region IVisitable Members

        public void Convert<T>(T mapping)
        {
            ((IMapping<ChartSheetSequence>)mapping).Apply(this);
        }

        #endregion
    }
}