using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
namespace DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class ChartSheetData : SheetData
    {
        public ChartSheetSequence ChartSheetSequence;

        public ChartSheetData()
        {
        }

        public override void Convert<T>(T mapping)
        {
            ((IMapping<ChartSheetSequence>)mapping).Apply(this.ChartSheetSequence);
        }
    }
}
