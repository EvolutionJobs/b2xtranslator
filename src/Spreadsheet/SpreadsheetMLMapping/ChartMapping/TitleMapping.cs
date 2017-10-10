

using DIaLOGIKa.b2xtranslator.CommonTranslatorLib;
using DIaLOGIKa.b2xtranslator.OpenXmlLib.DrawingML;
using DIaLOGIKa.b2xtranslator.Spreadsheet.XlsFileFormat;

namespace DIaLOGIKa.b2xtranslator.SpreadsheetMLMapping
{
    public class TitleMapping : AbstractChartMapping,
          IMapping<AttachedLabelSequence>
    {
        public TitleMapping(ExcelContext workbookContext, ChartContext chartContext)
            : base(workbookContext, chartContext)
        {
        }

        #region IMapping<AttachedLabelSequence> Members

        public void Apply(AttachedLabelSequence attachedLabelSequence)
        {
            // c:title
            this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElTitle, Dml.Chart.Ns);
            {
                // c:tx

                // c:layout

                // c:overlay

                // c:spPr

                // c:txPr

            }
            this._writer.WriteEndElement(); // c:title
        }
        #endregion
    }
}
